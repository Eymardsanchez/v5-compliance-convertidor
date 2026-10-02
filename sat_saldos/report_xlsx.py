"""Generación del papel de trabajo .xlsx. Referencia: blueprint sección 13.

Tres hojas con openpyxl y FÓRMULAS VIVAS (nunca totales precalculados en Python).
"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .analyzer_iva import AnalisisIVA
from .analyzer_iva_acumulado import AnalisisIVAAcumulado
from .config import MES_ABR
from .models import AcusePago, OrigenSaldo

# --- Estilos --------------------------------------------------------------- #
FUENTE = "Arial"
FMT_MONEDA = '$#,##0;($#,##0);-'
AZUL = "2E5496"
VERDE = "C6EFCE"
ROJO = "FFC7CE"
GRIS = "D9D9D9"

_thin = Side(style="thin", color="BFBFBF")
BORDE = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)


def _f(bold=False, color="000000", size=10):
    return Font(name=FUENTE, bold=bold, color=color, size=size)


def _header_cell(ws, row, col, texto):
    c = ws.cell(row=row, column=col, value=texto)
    c.font = _f(bold=True, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor=AZUL)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = BORDE
    return c


def _cell(ws, row, col, value, *, moneda=False, bold=False, fill=None,
          align="left", border=True):
    c = ws.cell(row=row, column=col, value=value)
    c.font = _f(bold=bold)
    if moneda:
        c.number_format = FMT_MONEDA
        align = "right"
    c.alignment = Alignment(horizontal=align, vertical="center")
    if fill:
        c.fill = PatternFill("solid", fgColor=fill)
    if border:
        c.border = BORDE
    return c


# --------------------------------------------------------------------------- #
def generar_reporte(
    ruta_salida: Path,
    *,
    razon_social: str,
    rfc: str,
    regimen: str,
    origenes_isr: list[OrigenSaldo],
    analisis_iva: AnalisisIVA,
    iva_acumulado: AnalisisIVAAcumulado | None = None,
    saldo_inicial_iva: float | None = None,
    pagos: list[AcusePago] | None = None,
) -> Path:
    wb = Workbook()

    coords = None
    if iva_acumulado is not None and iva_acumulado.renglones:
        coords = _hoja_iva_acumulado(wb, iva_acumulado, saldo_inicial_iva)   # crea la hoja y devuelve coords
    _hoja_resumen(wb, razon_social, rfc, regimen, origenes_isr, analisis_iva,
                  iva_acumulado, coords)
    _hoja_isr(wb, origenes_isr)
    _hoja_iva(wb, analisis_iva)
    if pagos:
        _hoja_pagos(wb, pagos)

    # Elimina la hoja por defecto si quedó vacía.
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]

    # Orden de pestañas.
    orden = ["Resumen", "ISR-Aplicacion", "IVA-Detalle", "IVA-Acumulado", "Pagos"]
    wb._sheets.sort(key=lambda s: orden.index(s.title) if s.title in orden else 99)

    ruta_salida = Path(ruta_salida)
    ruta_salida.parent.mkdir(parents=True, exist_ok=True)
    wb.save(ruta_salida)
    return ruta_salida


# --------------------------------------------------------------------------- #
def _hoja_resumen(wb, razon_social, rfc, regimen, origenes, iva: AnalisisIVA,
                  acum: AnalisisIVAAcumulado | None = None, coords: dict | None = None):
    ws = wb.active
    ws.title = "Resumen"
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCDEF", (34, 16, 16, 16, 22, 40)):
        ws.column_dimensions[col].width = w

    ws["A1"] = "Papel de trabajo — Saldos a favor ISR / IVA"
    ws["A1"].font = _f(bold=True, size=14, color=AZUL)
    ws["A2"] = f"Contribuyente: {razon_social}"
    ws["A3"] = f"RFC: {rfc}"
    ws["A4"] = f"Régimen: {regimen}"
    ws["A5"] = "Fuente: Acuses de Declaraciones y Pagos (SAT)"
    ws["A6"] = f"Elaborado: {datetime.now().strftime('%d/%m/%Y')}"
    for r in range(2, 7):
        ws[f"A{r}"].font = _f()

    # -- Tabla ISR ------------------------------------------------------------
    row = 8
    ws.cell(row=row, column=1, value="ISR — Saldos a favor y remanente").font = _f(bold=True, size=12)
    row += 1
    encabezados = ["Origen", "Monto original", "Aplicado", "Remanente", "Estatus", "Referencia"]
    for j, h in enumerate(encabezados, start=1):
        _header_cell(ws, row, j, h)
    row += 1
    fila_ini = row
    for o in origenes:
        _cell(ws, row, 1, o.etiqueta)
        _cell(ws, row, 2, o.monto_original, moneda=True)
        _cell(ws, row, 3, o.total_aplicado, moneda=True)
        # Remanente = fórmula viva (B - C)
        c = _cell(ws, row, 4, f"=B{row}-C{row}", moneda=True, bold=True)
        _cell(ws, row, 5, o.estatus)
        _cell(ws, row, 6, "Confirmar en el aplicativo del SAT")
        row += 1
    fila_fin = row - 1
    # Fila total (SUM vivo)
    _cell(ws, row, 1, "TOTAL", bold=True, fill=GRIS)
    for col in (2, 3, 4):
        L = get_column_letter(col)
        _cell(ws, row, col,
              f"=SUM({L}{fila_ini}:{L}{fila_fin})" if fila_fin >= fila_ini else 0,
              moneda=True, bold=True, fill=GRIS)
    _cell(ws, row, 5, "", fill=GRIS)
    _cell(ws, row, 6, "", fill=GRIS)
    row += 2

    # Nota técnica ISR
    ws.cell(row=row, column=1,
            value=("Notas ISR: (1) El remanente nominal = monto original − Σ compensado "
                   "(solo declaraciones vigentes). (2) El SAT actualiza el saldo por INPC "
                   "(Art. 17-A CFF); el remanente actualizado del último acuse es "
                   "informativo. (3) Una complementaria puede revertir una compensación; "
                   "no se cuenta. (4) Confirmar el remanente exacto en el aplicativo del SAT.")
            ).font = _f(size=9)
    ws.row_dimensions[row].height = 60
    row += 2

    # -- Tabla IVA ------------------------------------------------------------
    ws.cell(row=row, column=1, value="IVA — Resultados por periodo").font = _f(bold=True, size=12)
    row += 1
    _cell(ws, row, 1, "Meses con saldo a favor", bold=True)
    _cell(ws, row, 2, "=COUNTIF('IVA-Detalle'!C:C,\">0\")", bold=True)
    row += 1
    _cell(ws, row, 1, "Suma de saldos a favor (aritmética)", bold=True)
    _cell(ws, row, 2, "=SUM('IVA-Detalle'!C:C)", moneda=True, bold=True, fill=VERDE)
    row += 1
    _cell(ws, row, 1, "IVA a cargo (pagado)", bold=True)
    _cell(ws, row, 2, "=SUM('IVA-Detalle'!D:D)", moneda=True, bold=True, fill=ROJO)
    row += 2

    # -- Bloque IVA ACUMULADO (declaraciones detalladas) ----------------------
    if acum is not None and acum.renglones and coords:
        ws.cell(row=row, column=1,
                value="IVA — Remanente acumulado (declaraciones detalladas)").font = _f(bold=True, size=12)
        row += 1
        _cell(ws, row, 1, "Saldo a favor generado en el periodo", bold=True)
        _cell(ws, row, 2, f"='IVA-Acumulado'!{coords['gen']}", moneda=True, bold=True, fill=VERDE)
        row += 1
        _cell(ws, row, 1, "Acreditado a periodos anteriores", bold=True)
        _cell(ws, row, 2, f"='IVA-Acumulado'!{coords['acr']}", moneda=True, bold=True, fill=ROJO)
        row += 1
        _cell(ws, row, 1, "Cambio neto del acumulado (gen − acred)", bold=True)
        _cell(ws, row, 2, f"='IVA-Acumulado'!{coords['gen']}-'IVA-Acumulado'!{coords['acr']}",
              moneda=True, bold=True)
        row += 1
        _cell(ws, row, 1, "Saldo a favor inicial requerido (mín., del año previo)", bold=True)
        _cell(ws, row, 2, f"='IVA-Acumulado'!{coords['inicial']}", moneda=True, bold=True, fill=GRIS)
        row += 1
        _cell(ws, row, 1, "Remanente al cierre (con saldo inicial capturado)", bold=True)
        _cell(ws, row, 2, f"='IVA-Acumulado'!{coords['cierre']}", moneda=True, bold=True, fill=VERDE)
        row += 2

    # Caja de ALCANCE / NOTA IVA
    if acum is not None and acum.renglones:
        texto = ("✔ IVA determinable con declaraciones detalladas. " + acum.nota
                 + " Captura el saldo a favor inicial en 'IVA-Acumulado' para el "
                   "remanente exacto; confírmalo en el aplicativo del SAT.")
        color_texto, color_fondo = "1F6B1F", VERDE
    else:
        texto = "⚠ " + iva.texto_alcance
        color_texto, color_fondo = "9C0006", ROJO
    ac = ws.cell(row=row, column=1, value=texto)
    ac.font = _f(bold=True, color=color_texto)
    ac.fill = PatternFill("solid", fgColor=color_fondo)
    ac.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
    ws.row_dimensions[row].height = 110


def _hoja_isr(wb, origenes: list[OrigenSaldo]):
    ws = wb.create_sheet("ISR-Aplicacion")
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCDEFGH", (16, 14, 18, 14, 14, 20, 14, 34)):
        ws.column_dimensions[col].width = w

    row = 1
    cols = ["Periodo aplicado", "Fecha present.", "Tipo declaración", "ISR a cargo",
            "Compensado", "Remanente actualiz. (acuse)", "Cantidad a pagar", "Nota"]

    for o in origenes:
        # Encabezado del bloque (gris)
        titulo = (f"{o.etiqueta} = ${o.monto_original:,.0f}"
                  f"  (fecha: {o.fecha_origen or 's/f'}, op: {o.op_origen or 's/n'})")
        hc = ws.cell(row=row, column=1, value=titulo)
        hc.font = _f(bold=True)
        hc.fill = PatternFill("solid", fgColor=GRIS)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
        row += 1

        for j, h in enumerate(cols, start=1):
            _header_cell(ws, row, j, h)
        row += 1

        ini = row
        for a in o.aplicaciones:
            nota = ""
            if a.get('subtipo') and a['subtipo'] != 'normal':
                nota = f"Vigente: {a['subtipo']}"
            if a.get('compensado') in (0, 0.0):
                nota = (nota + " · " if nota else "") + "Compensación revertida/anulada"
            _cell(ws, row, 1, a.get('periodo'))
            _cell(ws, row, 2, a.get('fecha'))
            _cell(ws, row, 3, a.get('tipo'))
            _cell(ws, row, 4, a.get('isr_a_cargo'), moneda=True)
            _cell(ws, row, 5, a.get('compensado'), moneda=True)
            _cell(ws, row, 6, a.get('remanente_act'), moneda=True)
            _cell(ws, row, 7, a.get('cantidad_pagar'), moneda=True)
            _cell(ws, row, 8, nota)
            row += 1
        fin = row - 1

        # TOTAL COMPENSADO (SUM vivo)
        _cell(ws, row, 3, "TOTAL COMPENSADO", bold=True, fill=GRIS)
        _cell(ws, row, 5,
              f"=SUM(E{ini}:E{fin})" if fin >= ini else 0,
              moneda=True, bold=True, fill=GRIS)
        row += 1
        # Monto original
        _cell(ws, row, 3, "Monto original", bold=True)
        _cell(ws, row, 5, o.monto_original, moneda=True, bold=True)
        fila_orig = row
        fila_total = row - 1
        row += 1
        # REMANENTE NOMINAL = original - total (verde)
        _cell(ws, row, 3, "REMANENTE NOMINAL", bold=True, fill=VERDE)
        _cell(ws, row, 5, f"=E{fila_orig}-E{fila_total}", moneda=True, bold=True, fill=VERDE)
        row += 2

    if not origenes:
        ws.cell(row=1, column=1, value="Sin orígenes de saldo a favor de ISR detectados.")


def _hoja_iva(wb, iva: AnalisisIVA):
    ws = wb.create_sheet("IVA-Detalle")
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCDEF", (12, 10, 16, 16, 16, 24)):
        ws.column_dimensions[col].width = w

    ws.cell(row=1, column=1, value="IVA — Resultados por periodo (acuses)").font = _f(bold=True, size=12)
    row = 2
    cols = ["Ejercicio", "Mes", "Saldo a favor", "IVA a cargo", "Cant. a pagar", "Tipo · Op."]
    for j, h in enumerate(cols, start=1):
        _header_cell(ws, row, j, h)
    row += 1

    ini = row
    for r in iva.resultados:
        _cell(ws, row, 1, r.ejercicio, align="center")
        _cell(ws, row, 2, MES_ABR[r.periodo] if 1 <= r.periodo <= 12 else r.periodo,
              align="center")
        _cell(ws, row, 3, r.a_favor, moneda=True, fill=VERDE if r.a_favor > 0 else None)
        _cell(ws, row, 4, r.a_cargo, moneda=True, fill=ROJO if r.a_cargo > 0 else None)
        _cell(ws, row, 5, r.cantidad_pagar, moneda=True)
        _cell(ws, row, 6, f"{r.tipo} · {r.operacion}")
        row += 1
    fin = row - 1

    _cell(ws, row, 2, "TOTAL", bold=True, fill=GRIS)
    for col in (3, 4, 5):
        L = get_column_letter(col)
        _cell(ws, row, col,
              f"=SUM({L}{ini}:{L}{fin})" if fin >= ini else 0,
              moneda=True, bold=True, fill=GRIS)
    _cell(ws, row, 6, "", fill=GRIS)
    row += 2

    ws.cell(row=row, column=1,
            value=("Nota: la suma de la columna 'Saldo a favor' es ARITMÉTICA y NO "
                   "equivale al remanente acumulado de IVA. " + iva.texto_alcance)
            ).font = _f(size=9, color="9C0006")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
    ws.row_dimensions[row].height = 75


def _hoja_iva_acumulado(wb, acum: AnalisisIVAAcumulado, saldo_inicial: float | None = None) -> dict:
    """Reconstrucción del remanente acumulado de IVA con fórmulas vivas.

    La celda de 'Saldo a favor inicial' es editable: al capturar el arrastre del
    ejercicio anterior, toda la columna de remanente acumulado recalcula.
    Devuelve las coordenadas de las celdas de totales para enlazarlas en Resumen.
    """
    ws = wb.create_sheet("IVA-Acumulado")
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCDEFGHI", (12, 16, 18, 16, 18, 18, 18, 12, 26)):
        ws.column_dimensions[col].width = w

    ws.cell(row=1, column=1, value="IVA — Remanente acumulado (declaraciones detalladas)").font = _f(bold=True, size=12)

    # Celda editable: saldo a favor inicial del ejercicio anterior.
    # Por defecto = mínimo requerido detectado; si se captura el cierre real del
    # año previo (--saldo-inicial-iva), se usa ese y el remanente cierra exacto.
    b2 = saldo_inicial if saldo_inicial is not None else acum.saldo_inicial_requerido
    _cell(ws, 2, 1, "Saldo a favor inicial (ejercicio anterior):", bold=True, border=False)
    cell_ini = _cell(ws, 2, 2, b2, moneda=True, bold=True, fill=GRIS)
    coord_ini = f"$B$2"
    _cell(ws, 3, 1,
          f"(mínimo requerido detectado: se editable; captura el saldo real al cierre "
          f"del ejercicio {(acum.ejercicio or 1) - 1})", border=False).font = _f(size=8, color="808080")

    encab = ["Periodo", "No. operación", "Fecha present.", "IVA a cargo periodo",
             "Saldo a favor generado", "Acreditado (per. ant.)", "Remanente acumulado",
             "Tipo", "Nota"]
    hrow = 5
    for j, h in enumerate(encab, start=1):
        _header_cell(ws, hrow, j, h)

    row = hrow + 1
    ini_datos = row
    prev_e = None
    ejercicio_actual = None
    for r in acum.renglones:
        # Banda de encabezado por ejercicio (agrupa los meses de cada año).
        if r.ejercicio != ejercicio_actual:
            ejercicio_actual = r.ejercicio
            bc = ws.cell(row=row, column=1, value=f"Ejercicio {r.ejercicio}")
            bc.font = _f(bold=True, color=AZUL)
            bc.fill = PatternFill("solid", fgColor=GRIS)
            bc.alignment = Alignment(horizontal="left", vertical="center")
            for cc in range(2, 10):
                ws.cell(row=row, column=cc).fill = PatternFill("solid", fgColor=GRIS)
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=9)
            row += 1
        etiqueta = f"{MES_ABR[r.periodo]}-{r.ejercicio}" if 1 <= r.periodo <= 12 else f"{r.periodo}-{r.ejercicio}"
        fecha = r.fecha_presentacion.strftime("%d/%m/%Y %H:%M") if r.fecha_presentacion else ""
        _cell(ws, row, 1, etiqueta, align="center")
        _cell(ws, row, 2, r.operacion or "", align="center")
        _cell(ws, row, 3, fecha, align="center")
        _cell(ws, row, 4, r.a_cargo_periodo, moneda=True)
        _cell(ws, row, 5, r.a_favor_generado, moneda=True, fill=VERDE if r.a_favor_generado > 0 else None)
        _cell(ws, row, 6, r.acreditamiento, moneda=True, fill=ROJO if r.acreditamiento > 0 else None)
        # Remanente acumulado = (inicial o remanente previo) + generado − acreditado  [fórmula viva]
        base = coord_ini if prev_e is None else f"G{prev_e}"
        _cell(ws, row, 7, f"={base}+E{row}-F{row}", moneda=True, bold=True)
        _cell(ws, row, 8, r.subtipo, align="center")
        nota = "Complementaria vigente" if r.subtipo not in ("normal", "") else ""
        _cell(ws, row, 9, nota)
        prev_e = row
        row += 1
    fin_datos = row - 1

    # Fila TOTAL (SUM vivo de generado y acreditado).
    _cell(ws, row, 1, "TOTAL", bold=True, fill=GRIS)
    _cell(ws, row, 2, "", fill=GRIS)
    _cell(ws, row, 3, "", fill=GRIS)
    _cell(ws, row, 4, "", fill=GRIS)
    _cell(ws, row, 5, f"=SUM(E{ini_datos}:E{fin_datos})", moneda=True, bold=True, fill=GRIS)
    _cell(ws, row, 6, f"=SUM(F{ini_datos}:F{fin_datos})", moneda=True, bold=True, fill=GRIS)
    _cell(ws, row, 7, f"=G{fin_datos}", moneda=True, bold=True, fill=VERDE)  # remanente al cierre
    _cell(ws, row, 8, "", fill=GRIS)
    _cell(ws, row, 9, "Remanente al cierre →", fill=GRIS)
    total_row = row
    row += 2

    ws.cell(row=row, column=1, value=("Mecánica (Art. 6 LIVA): remanente[m] = remanente[m-1] "
            "+ saldo a favor generado[m] − acreditamiento aplicado[m]. " + acum.nota)
            ).font = _f(size=9)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=9)
    ws.row_dimensions[row].height = 95
    row += 2

    # -- Origen del remanente al cierre (PEPS) --------------------------------
    if acum.capas_cierre:
        ws.cell(row=row, column=1,
                value="Origen del remanente al cierre (PEPS — lo más antiguo se aplica primero)"
                ).font = _f(bold=True, size=11)
        row += 1
        for j, h in enumerate(["Mes de origen", "No. operación", "Fecha present.",
                               "Generado", "Sigue vivo al cierre"], start=1):
            _header_cell(ws, row, j, h)
        row += 1
        ini_cap = row
        for c in acum.capas_cierre:
            fecha = c.fecha_presentacion.strftime("%d/%m/%Y %H:%M") if c.fecha_presentacion else ""
            _cell(ws, row, 1, c.origen, align="center")
            _cell(ws, row, 2, c.operacion or "", align="center")
            _cell(ws, row, 3, fecha, align="center")
            _cell(ws, row, 4, c.generado, moneda=True)
            _cell(ws, row, 5, c.restante, moneda=True, fill=VERDE)
            row += 1
        _cell(ws, row, 1, "TOTAL", bold=True, fill=GRIS)
        _cell(ws, row, 2, "", fill=GRIS)
        _cell(ws, row, 3, "", fill=GRIS)
        _cell(ws, row, 4, "", fill=GRIS)
        _cell(ws, row, 5, f"=SUM(E{ini_cap}:E{row - 1})", moneda=True, bold=True, fill=GRIS)

    return {
        "gen": f"E{total_row}",
        "acr": f"F{total_row}",
        "cierre": f"G{total_row}",
        "inicial": "B2",
    }


def _hoja_pagos(wb, pagos: list[AcusePago]):
    """Verificación de pago por declaración: importe pagado, actualización y recargos.

    Prueba si cada declaración realmente se pagó (sección 'Información del pago
    recibido' del acuse de línea de captura). Fórmulas vivas para el importe a
    pagar (contribución + actualización + recargos) y las diferencias.
    """
    ws = wb.create_sheet("Pagos")
    ws.sheet_view.showGridLines = False
    anchos = (11, 13, 14, 13, 13, 13, 14, 14, 13, 13, 18, 14)
    for col, w in zip("ABCDEFGHIJKL", anchos):
        ws.column_dimensions[col].width = w

    ws.cell(row=1, column=1,
            value="Pagos — Verificación de declaraciones pagadas (línea de captura)"
            ).font = _f(bold=True, size=12)

    encab = ["Periodo", "Tipo", "No. operación", "Vencimiento", "Contribución",
             "Actualización", "Recargos", "Importe a pagar", "Importe pagado",
             "Fecha de pago", "Institución", "Estatus"]
    hrow = 2
    for j, h in enumerate(encab, start=1):
        _header_cell(ws, hrow, j, h)

    row = hrow + 1
    ini = row
    for p in pagos:
        etiqueta = f"{MES_ABR[p.periodo]}-{p.ejercicio}" if p.periodo and 1 <= p.periodo <= 12 else f"{p.periodo}-{p.ejercicio}"
        pagado = bool(p.fecha_pago) and (p.importe_pagado or 0) > 0
        _cell(ws, row, 1, etiqueta, align="center")
        _cell(ws, row, 2, p.tipo, align="center")
        _cell(ws, row, 3, p.operacion or "", align="center")
        _cell(ws, row, 4, p.vencimiento or "", align="center")
        _cell(ws, row, 5, p.contribucion, moneda=True)
        _cell(ws, row, 6, p.actualizacion, moneda=True,
              fill=ROJO if p.actualizacion > 0 else None)
        _cell(ws, row, 7, p.recargos, moneda=True, fill=ROJO if p.recargos > 0 else None)
        # Importe a pagar = contribución + actualización + recargos (fórmula viva).
        _cell(ws, row, 8, f"=E{row}+F{row}+G{row}", moneda=True, bold=True)
        _cell(ws, row, 9, p.importe_pagado if p.importe_pagado is not None else 0,
              moneda=True, fill=VERDE if pagado else ROJO)
        _cell(ws, row, 10, p.fecha_pago or "", align="center")
        _cell(ws, row, 11, p.institucion or "")
        estatus = "✓ Pagado" if pagado else "⚠ Sin comprobante"
        _cell(ws, row, 12, estatus, align="center",
              fill=VERDE if pagado else ROJO)
        row += 1
    fin = row - 1

    # Fila TOTAL (SUM vivo).
    _cell(ws, row, 1, "TOTAL", bold=True, fill=GRIS)
    for col in (2, 3, 4):
        _cell(ws, row, col, "", fill=GRIS)
    for col in (5, 6, 7, 8, 9):
        L = get_column_letter(col)
        _cell(ws, row, col, f"=SUM({L}{ini}:{L}{fin})" if fin >= ini else 0,
              moneda=True, bold=True, fill=GRIS)
    for col in (10, 11, 12):
        _cell(ws, row, col, "", fill=GRIS)
    total_row = row
    row += 1

    # Diferencia pagado vs a pagar (verificación).
    _cell(ws, row, 8, "Dif. pagado − a pagar →", bold=True)
    _cell(ws, row, 9, f"=I{total_row}-H{total_row}", moneda=True, bold=True)
    row += 2

    ws.cell(row=row, column=1, value=(
        "Notas: (1) 'Importe pagado' proviene de la sección 'Información del pago "
        "recibido' del acuse; si está en rojo/⚠, no se detectó comprobante de pago. "
        "(2) Actualización = Art. 17-A CFF; Recargos = Art. 21 CFF. (3) Importe a "
        "pagar = contribución + actualización + recargos. (4) Es la declaración "
        "vigente por periodo (complementaria de mayor fecha). Confirma en el SAT.")
        ).font = _f(size=9)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=12)
    ws.row_dimensions[row].height = 60
