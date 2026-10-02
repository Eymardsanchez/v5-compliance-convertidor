Attribute VB_Name = "PapelTrabajo"
Option Explicit

' ==========================================================================
'  Papel de trabajo automatico - conciliacion manual dentro del reporte
'  El usuario captura el UUID en la columna G de una hoja de banco y ejecuta
'  ConciliarManuales (Alt+F8 o boton de cinta). La macro:
'   - busca el UUID en Recibidos / Emitidos / Nomina / Comp. Pago (REP tipo P)
'   - acepta VARIOS UUID en la misma celda separados por "|" o "+" (un
'     movimiento que paga dos o mas CFDI); ver PartesUUID
'   - trae los datos del CFDI a la fila de banco (Serie, Folio, RFC, Nombre, Fecha)
'   - marca Conciliado (manual) en ambas hojas y escribe los datos bancarios
'     de vuelta en la hoja del CFDI.
'  No toca lo ya auto-conciliado (estatus que empieza con "Conciliado").
' ==========================================================================

Private Const HDR    As Long = 3     ' fila de encabezados en hojas de banco
' Col A = ID (llave de sync con la BD). Todo lo demas se corrio una a la derecha.
Private Const C_ID   As Long = 1     ' columna A - ID ESTABLE (llave de sync BD)
Private Const C_UUID As Long = 8     ' columna H - UUID que captura el usuario
' Columnas usadas por el respaldo/restauracion del ID (hoja de banco):
Private Const C_FEC  As Long = 2     ' B - FECHA
Private Const C_DES  As Long = 4     ' D - DESCRIPCION
Private Const C_RET  As Long = 5     ' E - RETIRO
Private Const C_DEP  As Long = 6     ' F - DEPOSITO
Private Const RESPHOJA As String = "_RESPALDO_IDS"

Public Sub ConciliarManuales()
    Dim ws As Worksheet, ult As Long, f As Long, n As Long
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo fin
    AsignarIDs                       ' asigna ID a filas nuevas antes de conciliar
    For Each ws In ThisWorkbook.Worksheets
        If EsBanco(ws) Then
            ult = ws.Cells(ws.Rows.Count, C_UUID).End(xlUp).Row
            For f = HDR + 1 To ult
                If Len(Trim(ws.Cells(f, C_UUID).Value)) > 0 _
                   And (UCase(ws.Cells(f, 16).Value) Like "CONCILIADO*") = False Then
                    If Conciliar(ws, f) Then n = n + 1
                End If
            Next f
        End If
    Next ws
fin:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox n & " movimiento(s) conciliado(s) manualmente.", vbInformation, "Papel de trabajo"
End Sub

' ==========================================================================
'  AsignarIDs - genera el ID ESTABLE (col A) a las filas de banco que tengan
'  datos pero no tengan ID todavia (p.ej. movimientos agregados a mano en el
'  Excel). El ID es el MAXIMO global existente + 1: NUNCA reusa un ID borrado
'  (los huecos se conservan), asi la reimportacion casa por ID sin sobrescribir.
'  Se ejecuta solo al ABRIR el archivo (Workbook_Open) y al conciliar (arriba).
' ==========================================================================
Public Sub AsignarIDs()
    Dim ws As Worksheet, ult As Long, f As Long, maxID As Long
    Dim v As Variant
    Application.ScreenUpdating = False
    On Error Resume Next
    ' 1) Maximo ID actual en TODAS las hojas de banco.
    maxID = 0
    For Each ws In ThisWorkbook.Worksheets
        If EsBanco(ws) Then
            ult = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row   ' col B = FECHA
            For f = HDR + 1 To ult
                v = ws.Cells(f, C_ID).Value
                If IsNumeric(v) Then
                    If CLng(v) > maxID Then maxID = CLng(v)
                End If
            Next f
        End If
    Next ws
    ' 2) Asignar max+1 a las filas con datos (col B) pero sin ID.
    For Each ws In ThisWorkbook.Worksheets
        If EsBanco(ws) Then
            ult = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
            For f = HDR + 1 To ult
                If Len(Trim(ws.Cells(f, 2).Value)) > 0 _
                   And Len(Trim(ws.Cells(f, C_ID).Value)) = 0 Then
                    maxID = maxID + 1
                    ws.Cells(f, C_ID).Value = maxID
                End If
            Next f
        End If
    Next ws
    Application.ScreenUpdating = True
End Sub

Private Function EsBanco(ByVal Sh As Object) As Boolean
    ' Deteccion por firma estructural, no por nombre (hay nombres dinamicos: SANTANDER MAY, etc.)
    On Error Resume Next
    EsBanco = (UCase(Trim(Sh.Cells(HDR, 2).Value)) = "FECHA") _
          And (InStr(UCase(Sh.Cells(HDR, C_UUID).Value), "UUID") > 0)
End Function

' ==========================================================================
'  PartesUUID - un mismo movimiento de banco puede pagar VARIOS CFDI. En la
'  hoja de BANCO (col H) el usuario captura los UUID en la MISMA celda
'  separados por "|" o "+" (tambien se aceptan "," ";" tabulador, espacio y
'  salto de linea). Devuelve cada UUID por separado; arreglo vacio si no hay
'  nada capturado. Un UUID nunca contiene estos caracteres, asi que cortar por
'  ellos no puede partir un UUID valido a la mitad.
' ==========================================================================
Private Function PartesUUID(ByVal s As String) As Variant
    Dim t As String, sep As Variant, p() As String, o() As String
    Dim i As Long, n As Long
    t = s
    For Each sep In Array("|", "+", ",", ";", vbCrLf, vbLf, vbCr, vbTab, " ")
        t = Replace(t, CStr(sep), Chr(1))
    Next sep
    p = Split(t, Chr(1))
    ReDim o(0 To UBound(p))
    For i = 0 To UBound(p)
        If Len(Trim$(p(i))) > 0 Then
            o(n) = Trim$(p(i))
            n = n + 1
        End If
    Next i
    If n = 0 Then
        PartesUUID = Split("", Chr(1))      ' arreglo vacio: UBound = -1
    Else
        ReDim Preserve o(0 To n - 1)
        PartesUUID = o
    End If
End Function

' Acumula valores de varios CFDI en un solo texto, sin separadores sueltos.
Private Function Unir(ByVal acum As String, ByVal v As String, _
                      ByVal sep As String) As String
    If Len(Trim$(v)) = 0 Then
        Unir = acum
    ElseIf Len(acum) = 0 Then
        Unir = Trim$(v)
    Else
        Unir = acum & sep & Trim$(v)
    End If
End Function

' Layout de la hoja donde vive el CFDI encontrado. Cada una numera distinto:
'   0 = Recibidos / Emitidos   1 = Nomina (tipo N)   2 = Comp. Pago (REP tipo P)
Private Function TipoHojaCFDI(ByVal nombre As String) As Long
    Dim n As String: n = UCase$(Trim$(nombre))
    If n = "NOMINA" Then
        TipoHojaCFDI = 1
    ElseIf InStr(n, "COMP. PAGO") > 0 Or InStr(n, "COMPLEMENTO") > 0 Then
        TipoHojaCFDI = 2
    Else
        TipoHojaCFDI = 0
    End If
End Function

Private Function Conciliar(ByVal Sh As Object, ByVal f As Long) As Boolean
    Dim lista As Variant
    lista = PartesUUID(CStr(Sh.Cells(f, C_UUID).Value))
    If UBound(lista) < 0 Then Exit Function          ' celda vacia
    Dim hc As Worksheet, fc As Long, i As Long, k As Long, tipoH As Long
    Dim serie As String, folio As String, rfc As String
    Dim nom As String, femi As String, faltan As String
    For i = 0 To UBound(lista)
        If Buscar(CStr(lista(i)), hc, fc) Then
            tipoH = TipoHojaCFDI(hc.Name)
            Select Case tipoH
            Case 1
                ' Nomina: 1=UUID 2=FOLIO 3=FECHA EMISION 4=RFC EMPLEADO 5=NOMBRE EMPLEADO
                folio = Unir(folio, CStr(hc.Cells(fc, 2).Value), " | ")
                rfc = Unir(rfc, CStr(hc.Cells(fc, 4).Value), " | ")
                nom = Unir(nom, CStr(hc.Cells(fc, 5).Value), " | ")
                femi = Unir(femi, CStr(hc.Cells(fc, 3).Value), " | ")
            Case 2
                ' Comp. Pago (REP): 1=UUID COMPLEMENTO P 2=FECHA PAGO
                ' 6=RFC EMISOR 7=RAZON EMISOR 8=RFC RECEPTOR 9=RAZON RECEPTOR
                ' El REP no tiene serie ni folio; la contraparte es el EMISOR en
                ' los recibidos y el RECEPTOR en los emitidos.
                Dim colRFC As Long
                If InStr(UCase$(hc.Name), "EMITID") > 0 Then colRFC = 8 Else colRFC = 6
                folio = Unir(folio, "REP", " | ")
                rfc = Unir(rfc, CStr(hc.Cells(fc, colRFC).Value), " | ")
                nom = Unir(nom, CStr(hc.Cells(fc, colRFC + 1).Value), " | ")
                femi = Unir(femi, CStr(hc.Cells(fc, 2).Value), " | ")
            Case Else
                serie = Unir(serie, CStr(hc.Cells(fc, 2).Value), " | ")
                folio = Unir(folio, CStr(hc.Cells(fc, 3).Value), " | ")
                rfc = Unir(rfc, CStr(hc.Cells(fc, 5).Value), " | ")
                nom = Unir(nom, CStr(hc.Cells(fc, 6).Value), " | ")
                femi = Unir(femi, CStr(hc.Cells(fc, 4).Value), " | ")
            End Select
            ' CFDI  <- Banco   (hc: hoja de CFDI SIN columna ID, no cambia)
            Select Case tipoH
            Case 1
                ' Nomina (16 cols): 11=BANCO 12=FOLIO/REF 13=FECHA PAGO BANCO 14=ESTATUS
                hc.Cells(fc, 11).Value = Sh.Name                 ' Banco
                hc.Cells(fc, 12).Value = Sh.Cells(f, 3).Value    ' Folio / Ref bancaria (col C)
                hc.Cells(fc, 13).Value = Sh.Cells(f, 2).Value    ' Fecha pago banco (col B)
                hc.Cells(fc, 14).Value = "Conciliado (manual)"   ' Estatus
                Pintar hc, fc, 1, 16
            Case 2
                ' Comp. Pago (16 cols): 13=BANCO 14=ESTATUS. No hay columna de
                ' folio ni de fecha bancaria, y los montos no se tocan.
                hc.Cells(fc, 13).Value = Sh.Name                 ' Banco
                hc.Cells(fc, 14).Value = "Conciliado (manual)"   ' Estatus
                Pintar hc, fc, 1, 16
            Case Else
                ' Recibidos/Emitidos: 14=BANCO 15=FOLIO/REF 16=FECHA PAGO BANCO 18=ESTATUS
                hc.Cells(fc, 14).Value = Sh.Name                 ' Banco
                hc.Cells(fc, 15).Value = Sh.Cells(f, 3).Value    ' Folio / Ref bancaria (col C)
                hc.Cells(fc, 16).Value = Sh.Cells(f, 2).Value    ' Fecha pago banco (col B)
                hc.Cells(fc, 18).Value = "Conciliado (manual)"   ' Estatus
                Pintar hc, fc, 1, 18
            End Select
            k = k + 1
        Else
            faltan = Unir(faltan, CStr(lista(i)), ", ")
        End If
    Next i
    If k = 0 Then
        Sh.Cells(f, 16).Value = "UUID no encontrado"
        Sh.Cells(f, 16).Interior.Color = RGB(250, 220, 218)
        Exit Function
    End If
    ' Banco  <- CFDI   (columnas de banco corridas +1 por el ID en col A)
    Sh.Cells(f, 10).Value = serie
    Sh.Cells(f, 11).Value = folio
    Sh.Cells(f, 12).Value = rfc
    Sh.Cells(f, 13).Value = nom
    Sh.Cells(f, 14).Value = femi
    Sh.Cells(f, 15).Value = Sh.Cells(f, 2).Value     ' Fecha pago = fecha banco (col B)
    Sh.Cells(f, 17).Value = "Manual"
    Pintar Sh, f, 8, 17
    If Len(faltan) > 0 Then
        ' A proposito NO empieza con "Conciliado": ConciliarManuales lo vuelve a
        ' intentar en la siguiente corrida, cuando el UUID faltante ya este en
        ' el reporte. Volver a correr reconstruye los textos, no los duplica.
        Sh.Cells(f, 16).Value = "Parcial (manual) " & k & " de " & _
                                (UBound(lista) + 1) & " - falta: " & faltan
        Sh.Cells(f, 16).Interior.Color = RGB(255, 242, 204)
    ElseIf k > 1 Then
        Sh.Cells(f, 16).Value = "Conciliado (manual) " & k & " CFDI"
    Else
        Sh.Cells(f, 16).Value = "Conciliado (manual)"
    End If
    Conciliar = (Len(faltan) = 0)
End Function

Private Function Buscar(ByVal u As String, ByRef hoja As Worksheet, ByRef fila As Long) As Boolean
    Dim nm As Variant, ws As Worksheet, r As Variant
    ' Comp. Pago Recibidos/Emitidos: ahi viven los REP (CFDI tipo P). Sin ellos,
    ' capturar el UUID de un REP en la hoja de banco daba "UUID no encontrado".
    For Each nm In Array("Recibidos", "Emitidos", "Nomina", _
                         "Comp. Pago Recibidos", "Comp. Pago Emitidos")
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(nm))
        On Error GoTo 0
        If Not ws Is Nothing Then
            r = Application.Match(u, ws.Columns(1), 0)   ' Match de texto es case-insensitive
            If Not IsError(r) Then
                Set hoja = ws: fila = CLng(r): Buscar = True: Exit Function
            End If
        End If
    Next nm
End Function

Private Sub Pintar(ByVal Sh As Object, ByVal f As Long, ByVal c1 As Long, ByVal c2 As Long)
    Sh.Range(Sh.Cells(f, c1), Sh.Cells(f, c2)).Interior.Color = RGB(220, 240, 220)
End Sub

Private Function ExcluirIVA(ByVal estatus As String) As Boolean
    Dim e As String: e = UCase$(Trim$(estatus))
    ExcluirIVA = (e = "SIN BANCO" Or e = "NO CONCILIADO" Or e = "NO ENCONTRADO")
End Function


' ==========================================================================
'  Respaldo / auto-restauracion del ID (col A) — protege la llave de sync.
'  Al abrir: se restaura el ID desde una hoja OCULTA de respaldo (que guarda,
'  por cada fila, su ID junto con su informacion contigua: fecha|descripcion|
'  retiro|deposito). Si el usuario cambia un ID, se vuelve a poner el correcto
'  buscando la fila por su contenido. Las filas NUEVAS (agregadas a mano) se
'  agregan al respaldo con el ID que les asigna AsignarIDs.
'  Columnas de la hoja de banco: A=ID(1) B=FECHA(2) C=FOLIO(3) D=DESC(4)
'                                E=RETIRO(5) F=DEPOSITO(6)
' ==========================================================================
Private Function HojaResp() As Worksheet
    Dim h As Worksheet
    On Error Resume Next
    Set h = ThisWorkbook.Worksheets(RESPHOJA)
    On Error GoTo 0
    If h Is Nothing Then
        Set h = ThisWorkbook.Worksheets.Add
        h.Name = RESPHOJA
        h.Cells(1, 1).Value = "HOJA"
        h.Cells(1, 2).Value = "ID"
        h.Cells(1, 3).Value = "CLAVE"
    End If
    h.Visible = 2   ' xlSheetVeryHidden: no aparece ni en "Mostrar hojas"
    Set HojaResp = h
End Function

Private Function ClaveFila(ByVal Sh As Object, ByVal f As Long) As String
    ClaveFila = UCase$(Trim$(CStr(Sh.Cells(f, C_FEC).Value))) & "|" & _
             UCase$(Trim$(CStr(Sh.Cells(f, C_DES).Value))) & "|" & _
             CStr(Sh.Cells(f, C_RET).Value) & "|" & _
             CStr(Sh.Cells(f, C_DEP).Value)
End Function

' Restaura el ID de cada fila de banco desde el respaldo (por contenido).
Public Sub RestaurarIDs()
    Dim h As Worksheet, ws As Worksheet, ult As Long, f As Long, r As Long
    Dim d As Object, k As String, idBak As Variant
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo salir
    Set h = HojaResp()
    ult = h.Cells(h.Rows.Count, 2).End(xlUp).Row
    If ult < 2 Then GoTo salir          ' respaldo vacio: nada que restaurar
    Set d = CreateObject("Scripting.Dictionary")
    For r = 2 To ult
        k = CStr(h.Cells(r, 1).Value) & Chr(1) & CStr(h.Cells(r, 3).Value)
        If Not d.Exists(k) Then d(k) = h.Cells(r, 2).Value
    Next r
    For Each ws In ThisWorkbook.Worksheets
        If EsBanco(ws) Then
            ult = ws.Cells(ws.Rows.Count, C_FEC).End(xlUp).Row
            For f = HDR + 1 To ult
                If Len(Trim(ws.Cells(f, C_FEC).Value)) > 0 Then
                    k = ws.Name & Chr(1) & ClaveFila(ws, f)
                    If d.Exists(k) Then
                        idBak = d(k)
                        If CStr(ws.Cells(f, C_ID).Value) <> CStr(idBak) Then
                            ws.Cells(f, C_ID).Value = idBak
                        End If
                    End If
                End If
            Next f
        End If
    Next ws
salir:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

' Agrega al respaldo cualquier (hoja, clave) que aun no tenga ID guardado.
' NUNCA sobrescribe un ID ya respaldado (conserva el original).
Public Sub RespaldarIDs()
    Dim h As Worksheet, ws As Worksheet, ult As Long, f As Long, r As Long
    Dim d As Object, k As String, prox As Long
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo salir
    Set h = HojaResp()
    Set d = CreateObject("Scripting.Dictionary")
    ult = h.Cells(h.Rows.Count, 2).End(xlUp).Row
    For r = 2 To ult
        k = CStr(h.Cells(r, 1).Value) & Chr(1) & CStr(h.Cells(r, 3).Value)
        d(k) = True
    Next r
    prox = h.Cells(h.Rows.Count, 2).End(xlUp).Row + 1
    If prox < 2 Then prox = 2
    For Each ws In ThisWorkbook.Worksheets
        If EsBanco(ws) Then
            ult = ws.Cells(ws.Rows.Count, C_FEC).End(xlUp).Row
            For f = HDR + 1 To ult
                If Len(Trim(ws.Cells(f, C_FEC).Value)) > 0 _
                   And Len(Trim(ws.Cells(f, C_ID).Value)) > 0 Then
                    k = ws.Name & Chr(1) & ClaveFila(ws, f)
                    If Not d.Exists(k) Then
                        h.Cells(prox, 1).Value = ws.Name
                        h.Cells(prox, 2).Value = ws.Cells(f, C_ID).Value
                        h.Cells(prox, 3).Value = ClaveFila(ws, f)
                        d(k) = True
                        prox = prox + 1
                    End If
                End If
            Next f
        End If
    Next ws
salir:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

' Secuencia al abrir: restaurar IDs -> asignar a filas nuevas -> respaldar.
Public Sub InicializarIDs()
    On Error Resume Next
    RestaurarIDs
    AsignarIDs
    RespaldarIDs
End Sub

' ==========================================================================
'  Calculo Previo de IVA - busca hojas "IVA Enero"..."IVA Diciembre" o
'  "Calculo Previo IVA" y recalcula cada una filtrando por mes y estatus.
'  Columnas Recibidos/Emitidos: 4=FECHA EMISION, 7=SUBTOTAL, 9=IVA,
'  10=IVA RETENIDO, 18=ESTATUS, 21=TIPO CFDI.
'  Excluye filas con estatus "Sin banco", "No conciliado", "No encontrado".
' ==========================================================================

Private Function MesDeHojaIVA(ByVal nombre As String) As Long
    Dim n As String: n = Trim$(nombre)
    Dim meses As Variant, i As Long
    meses = Array("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", _
                  "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")
    For i = 0 To 11
        If UCase$(n) = UCase$("IVA " & meses(i)) Then
            MesDeHojaIVA = i + 1: Exit Function
        End If
    Next i
    If UCase$(n) = "CALCULO PREVIO IVA" Then
        MesDeHojaIVA = 0: Exit Function
    End If
    MesDeHojaIVA = -1
End Function

Private Function MesDeFecha(ByVal v As Variant) As Long
    On Error Resume Next
    If IsDate(v) Then MesDeFecha = Month(CDate(v)): Exit Function
    Dim s As String: s = CStr(v)
    If Len(s) >= 7 And Mid$(s, 5, 1) = "-" Then
        MesDeFecha = Val(Mid$(s, 6, 2))
    ElseIf Len(s) >= 5 And Mid$(s, 3, 1) = "/" Then
        MesDeFecha = Val(Mid$(s, 4, 2))
    End If
End Function

Private Sub CalcularIVAHoja(ByVal wsIVA As Worksheet, ByVal mesFiltro As Long)
    ' Cols desglose IVA (ocultas, generadas por Python):
    '   33=BASE IVA 16%, 34=BASE IVA 8%, 35=IVA IMPORTE 8%,
    '   36=BASE IVA 0%, 37=BASE EXENTA, 38=TIPO TASA
    ' Ya vienen con signo (negativo para Egreso).
    Dim ing16 As Double, ing8 As Double, ing0 As Double, ingEx As Double
    Dim ivaTrasl16 As Double, ivaTrasl8 As Double, ivaRetPM As Double
    Dim gas16 As Double, gas8 As Double, gas0 As Double, gasEx As Double
    Dim ivaAcred16 As Double, ivaAcred8 As Double, ivaRetEnt As Double
    Dim mesCFDI As Long
    Dim b16 As Double, b8 As Double, b0 As Double, bEx As Double
    Dim tt As String, s7 As Double

    Dim wsR As Worksheet
    On Error Resume Next: Set wsR = ThisWorkbook.Worksheets("Recibidos"): On Error GoTo 0
    If Not wsR Is Nothing Then
        Dim rLast As Long: rLast = wsR.Cells(wsR.Rows.Count, 1).End(xlUp).Row
        Dim rr As Long
        For rr = 3 To rLast
            If ExcluirIVA(CStr(wsR.Cells(rr, 18).Value)) Then GoTo SigR
            If mesFiltro > 0 Then
                mesCFDI = MesDeFecha(wsR.Cells(rr, 4).Value)
                If mesCFDI <> mesFiltro Then GoTo SigR
            End If
            b16 = CDbl("0" & wsR.Cells(rr, 33).Value)
            b8 = CDbl("0" & wsR.Cells(rr, 34).Value)
            b0 = CDbl("0" & wsR.Cells(rr, 36).Value)
            bEx = CDbl("0" & wsR.Cells(rr, 37).Value)
            If Abs(b16) + Abs(b8) + Abs(b0) + Abs(bEx) < 0.01 Then
                tt = UCase$(Trim$(CStr(wsR.Cells(rr, 38).Value)))
                s7 = CDbl("0" & wsR.Cells(rr, 7).Value)
                Select Case tt
                    Case "16%": b16 = s7
                    Case "8%": b8 = s7
                    Case "0%": b0 = s7
                    Case "EXENTO": bEx = s7
                End Select
            End If
            ing16 = ing16 + b16
            ing8 = ing8 + b8
            ing0 = ing0 + b0
            ingEx = ingEx + bEx
            ivaTrasl16 = ivaTrasl16 + CDbl("0" & wsR.Cells(rr, 9).Value) _
                       - CDbl("0" & wsR.Cells(rr, 35).Value)
            ivaTrasl8 = ivaTrasl8 + CDbl("0" & wsR.Cells(rr, 35).Value)
            ivaRetPM = ivaRetPM + CDbl("0" & wsR.Cells(rr, 10).Value)
SigR:
        Next rr
    End If

    Dim wsE As Worksheet
    On Error Resume Next: Set wsE = ThisWorkbook.Worksheets("Emitidos"): On Error GoTo 0
    If Not wsE Is Nothing Then
        Dim eLast As Long: eLast = wsE.Cells(wsE.Rows.Count, 1).End(xlUp).Row
        Dim ee As Long
        For ee = 3 To eLast
            If ExcluirIVA(CStr(wsE.Cells(ee, 18).Value)) Then GoTo SigE
            If mesFiltro > 0 Then
                mesCFDI = MesDeFecha(wsE.Cells(ee, 4).Value)
                If mesCFDI <> mesFiltro Then GoTo SigE
            End If
            b16 = CDbl("0" & wsE.Cells(ee, 33).Value)
            b8 = CDbl("0" & wsE.Cells(ee, 34).Value)
            b0 = CDbl("0" & wsE.Cells(ee, 36).Value)
            bEx = CDbl("0" & wsE.Cells(ee, 37).Value)
            If Abs(b16) + Abs(b8) + Abs(b0) + Abs(bEx) < 0.01 Then
                tt = UCase$(Trim$(CStr(wsE.Cells(ee, 38).Value)))
                s7 = CDbl("0" & wsE.Cells(ee, 7).Value)
                Select Case tt
                    Case "16%": b16 = s7
                    Case "8%": b8 = s7
                    Case "0%": b0 = s7
                    Case "EXENTO": bEx = s7
                End Select
            End If
            gas16 = gas16 + b16
            gas8 = gas8 + b8
            gas0 = gas0 + b0
            gasEx = gasEx + bEx
            ivaAcred16 = ivaAcred16 + CDbl("0" & wsE.Cells(ee, 9).Value) _
                       - CDbl("0" & wsE.Cells(ee, 35).Value)
            ivaAcred8 = ivaAcred8 + CDbl("0" & wsE.Cells(ee, 35).Value)
            ivaRetEnt = ivaRetEnt + CDbl("0" & wsE.Cells(ee, 10).Value)
SigE:
        Next ee
    End If

    Dim totalCobrados As Double: totalCobrados = ing16 + ing8 + ing0 + ingEx
    Dim totalIvaTrasl As Double: totalIvaTrasl = ivaTrasl16 + ivaTrasl8
    Dim totalPagados As Double: totalPagados = gas16 + gas8 + gas0 + gasEx
    Dim totalIvaAcred As Double: totalIvaAcred = ivaAcred16 + ivaAcred8
    Dim saldo As Double: saldo = totalIvaTrasl - ivaRetPM - totalIvaAcred - ivaRetEnt

    wsIVA.Cells(4, 2).Value = ing16
    wsIVA.Cells(5, 2).Value = ing8
    wsIVA.Cells(6, 2).Value = ing0
    wsIVA.Cells(7, 2).Value = ingEx
    wsIVA.Cells(8, 2).Value = totalCobrados
    wsIVA.Cells(10, 2).Value = ivaTrasl16
    wsIVA.Cells(11, 2).Value = ivaTrasl8
    wsIVA.Cells(12, 2).Value = totalIvaTrasl
    wsIVA.Cells(14, 2).Value = ivaRetPM
    wsIVA.Cells(17, 2).Value = gas16
    wsIVA.Cells(19, 2).Value = gas16
    wsIVA.Cells(20, 2).Value = gas8
    wsIVA.Cells(22, 2).Value = gas8
    wsIVA.Cells(23, 2).Value = gas0
    wsIVA.Cells(24, 2).Value = gasEx
    wsIVA.Cells(26, 2).Value = gasEx
    wsIVA.Cells(27, 2).Value = totalPagados
    wsIVA.Cells(29, 2).Value = ivaAcred16
    wsIVA.Cells(31, 2).Value = ivaAcred16
    wsIVA.Cells(32, 2).Value = ivaAcred8
    wsIVA.Cells(34, 2).Value = ivaAcred8
    wsIVA.Cells(35, 2).Value = totalIvaAcred
    wsIVA.Cells(37, 2).Value = ivaRetEnt
    wsIVA.Cells(39, 2).Value = saldo
    If saldo < 0 Then
        wsIVA.Cells(39, 2).Font.Color = RGB(183, 28, 28)
    Else
        wsIVA.Cells(39, 2).Font.Color = RGB(27, 94, 32)
    End If
    wsIVA.Cells(41, 2).Value = 0
End Sub

Public Sub CalcularIVA()
    Dim ws As Worksheet, n As Long, mesHoja As Long, hojas As String
    Application.ScreenUpdating = False
    For Each ws In ThisWorkbook.Worksheets
        mesHoja = MesDeHojaIVA(ws.Name)
        If mesHoja >= 0 Then
            CalcularIVAHoja ws, mesHoja
            n = n + 1
            hojas = hojas & ws.Name & ", "
        End If
    Next ws
    Application.ScreenUpdating = True
    If n = 0 Then
        MsgBox "No se encontraron hojas de IVA (IVA Enero, IVA Febrero, etc.).", vbExclamation
    Else
        MsgBox n & " hoja(s) de IVA recalculada(s): " & Left$(hojas, Len(hojas) - 2), _
               vbInformation, "Calculo IVA"
    End If
End Sub
