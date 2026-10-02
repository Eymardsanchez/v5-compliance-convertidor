"""
exportar_conciliacion.py
Genera el Excel de conciliación banco ↔ CFDIs con hojas:
  • [BANCO]   — movimientos bancarios con columnas de CFDI cruzado
  • Recibidos — CFDIs recibidos con columna ID BANCO / Fecha de Pago
  • Emitidos  — CFDIs emitidos  con columna ID BANCO / Fecha de Pago
  • Resumen   — contadores y totales
"""

from __future__ import annotations

import zipfile
from collections import Counter
from pathlib import Path
from typing import Optional

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import sys

from cfdi_parser    import CFDI
from cfdi_validador import validar_cfdi, validar_lote
from conciliacion   import MovBanco, MovConciliado, ResultadoBanco, _mes_fecha


# ── Plantilla con macro VBA (papel de trabajo automático) ─────────────────────
# openpyxl NO puede sintetizar VBA; solo lo conserva (`keep_vba=True`) si parte de
# una plantilla .xlsm que ya lo contiene. Esa plantilla se construye una vez con
# `construir_plantilla.py` (macro `PapelTrabajo.bas`) y se empaqueta con la app
# (PyInstaller: agregar ('plantilla_papel_trabajo.xlsm', '.') a `datas`). Si la
# plantilla no está presente, el export degrada limpio a .xlsx sin macro.
def _ruta_recurso(nombre: str) -> Path:
    base = getattr(sys, "_MEIPASS", None) or Path(__file__).parent
    return Path(base) / nombre


_PLANTILLA_MACRO = _ruta_recurso("plantilla_papel_trabajo.xlsm")


def _nuevo_workbook():
    """Workbook vacío. Si existe la plantilla con macro, parte de ella
    (conservando el vbaProject.bin firmado) y elimina sus hojas; si no, crea uno
    normal .xlsx."""
    if _PLANTILLA_MACRO.exists():
        wb = openpyxl.load_workbook(_PLANTILLA_MACRO, keep_vba=True)
        for s in list(wb.worksheets):
            wb.remove(s)
        return wb
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    return wb


_PARTES_FIRMA_VBA = ("xl/_rels/vbaProject.bin.rels",)


def _restaurar_rels_firma(destino: Path) -> None:
    """Reinyecta las partes de la firma del VBA que openpyxl tira al guardar.

    openpyxl copia del archivo original solo lo que casa con
    `re.match("xl/vba|xl/ctrlProps|customUI|xl/activeX|...")`. Las partes de
    firma `xl/vbaProjectSignature*.bin` empiezan con "xl/vba" y sí pasan, pero
    **`xl/_rels/vbaProject.bin.rels` no** — y ese es justo el que declara la
    relación entre el proyecto VBA y sus firmas. Sin él Excel no puede asociar
    la firma y el reporte sale con la macro como NO firmada, aunque la plantilla
    estuviera firmada. Verificado comparando las partes de la plantilla contra
    las del reporte generado.
    """
    if not _PLANTILLA_MACRO.exists():
        return
    try:
        with zipfile.ZipFile(_PLANTILLA_MACRO) as zp:
            faltantes = {n: zp.read(n) for n in _PARTES_FIRMA_VBA
                         if n in zp.namelist()}
        if not faltantes:
            return
        with zipfile.ZipFile(destino) as zd:
            if all(n in zd.namelist() for n in faltantes):
                return          # openpyxl ya las conservó: nada que hacer
            items = [(i, zd.read(i.filename)) for i in zd.infolist()]
        tmp = destino.with_suffix(destino.suffix + ".tmp")
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
            for info, data in items:
                if info.filename not in faltantes:
                    zo.writestr(info, data)
            for nombre, data in faltantes.items():
                zo.writestr(nombre, data)
        tmp.replace(destino)
    except Exception:
        # Nunca romper la exportación por esto: sin firma el reporte sirve igual.
        pass


def _guardar_workbook(wb, output_path) -> str:
    """Guarda el workbook. Si lleva macro embebida, fuerza la extensión .xlsm
    (Excel rechaza abrir un .xlsx que contenga vbaProject.bin)."""
    out = Path(output_path)
    lleva_vba = getattr(wb, "vba_archive", None) is not None
    if lleva_vba:
        out = out.with_suffix(".xlsm")
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out))
    if lleva_vba:
        _restaurar_rels_firma(out)
    return str(out)

# ── Paleta ────────────────────────────────────────────────────────────────────
# Los colores están agrupados por el TRABAJO que hacen, no por su tono:
#   · chrome    — superficie, filas, tinta, hairlines
#   · estatus   — estado del renglón (reservados: nunca se usan como categoría)
#   · categoría — nivel de confianza del cruce (los 4 hues validados)
#   · encabezado— identidad de cada hoja (solo se ve una a la vez)
#
# Los 4 hues de categoría se validaron como paleta categórica contra superficie
# blanca en TODOS los pares (no solo adyacentes): separación CVD ΔE 9.2 y de
# visión normal ΔE 16.3, ambas sobre el piso. Un 5º hue rompía el piso, por eso
# las 7 fases se agrupan en 4 niveles de confianza y el DETALLE de la fase lo
# lleva el texto de la celda — el color nunca carga la identidad solo.

# Chrome
C_WHITE   = "FFFFFF"   # superficie
C_ALT     = "F7F7F5"   # fila alterna (banda apenas perceptible)
C_NEUTRO  = "F4F4F2"   # relleno de bloques neutros
C_INK     = "26251F"   # tinta primaria
C_INK_2   = "5A594F"   # tinta secundaria
C_LINEA   = "E1E0D9"   # hairline de rejilla
C_SAL     = "EEF4FC"   # azul muy claro — columna de saldo

# Estatus — RESERVADOS. Van siempre acompañados de su etiqueta de texto.
C_CONC        = "DCF0DC"; C_CONC_INK    = "14571B"   # conciliado
C_REVISAR     = "FBEBC9"; C_REVISAR_INK = "7A5200"   # requiere revisión
C_SIN         = "FADCDA"; C_SIN_INK     = "A32118"   # sin CFDI / crítico
C_PEND        = C_REVISAR                            # pendiente

# Categoría — nivel de confianza del cruce (orden fijo, nunca ciclado)
TIER_DIRECTO   = "2A78D6"   # 1 azul     — Fase 1: 1:1 por identidad + importe
TIER_COMPUESTO = "EB6834"   # 2 naranja  — Fases 2/3/5: sumas y agrupaciones
TIER_COMPLEM   = "1BAF7A"   # 3 aqua     — Fase 4: respaldado por complemento
TIER_DEBIL     = "4A3AA7"   # 4 violeta  — Fases 6/7: sin complemento / solo importe

# Paso oscuro de cada hue, para swatches sólidos con texto blanco (≥5.3:1).
# El hue brillante de arriba solo se usa como marca fina (borde) sobre blanco,
# donde necesita croma; en bloque sólido no daría contraste de texto.
_TIER_OSCURO: dict[str, str] = {
    TIER_DIRECTO:   "1B5FAF",
    TIER_COMPUESTO: "A8480F",
    TIER_COMPLEM:   "0F7A57",
    TIER_DEBIL:     "4A3AA7",
}

# Encabezados por hoja — solo una hoja se ve a la vez, así que aquí no aplica
# el piso de todos-los-pares. Texto blanco sobre todos ellos ≥ 5.3:1.
C_HDR    = "1B5FAF"   # Bancos
C_HDR_R  = "0F7A57"   # Recibidos
C_HDR_E  = "4A3AA7"   # Emitidos
C_HDR_N  = "A8480F"   # Nómina
C_HDR_S  = "33322D"   # Resumen / bloques neutros
C_HDR_P  = "0B6068"   # Complementos de pago (tipo P)

# Validación fiscal CFDI — reutiliza la escala de estatus
C_VAL_VERDE    = C_CONC      # VERDE    — deducible y conciliable
C_VAL_AMARILLO = C_REVISAR   # AMARILLO — requiere revisión
C_VAL_ROJO     = C_SIN       # ROJO     — no deducible
C_VAL_NEGRO    = C_HDR_S     # NEGRO    — sin efecto fiscal (crítico)
C_ART27        = "A8480F"    # Art.27 LIVA efectivo > $2,000


# Fases → nivel de confianza (hue + etiqueta del nivel).
_TIER_FASE: dict[str, tuple[str, str]] = {
    "Exacto":              (TIER_DIRECTO,   "Directo"),
    "Multi-Movimiento":    (TIER_COMPUESTO, "Compuesto"),
    "Combinatorio":        (TIER_COMPUESTO, "Compuesto"),
    "Global Mensual":      (TIER_COMPUESTO, "Compuesto"),
    "PPD c/ Complemento":  (TIER_COMPLEM,   "Con complemento"),
    "Recepción de Pago":   (TIER_COMPLEM,   "Con complemento"),
    "PPD Sin Complemento": (TIER_DEBIL,     "Débil"),
    "Solo Importe":        (TIER_DEBIL,     "Débil"),
}


def _tier(tipo_match: str) -> tuple[str, str]:
    """Devuelve (hue, etiqueta del nivel) para un tipo_match. Los grupos
    configurados ('Grupo: X') caen en Compuesto."""
    if not tipo_match:
        return "", ""
    if tipo_match.startswith("Grupo:"):
        return TIER_COMPUESTO, "Compuesto"
    return _TIER_FASE.get(tipo_match, (C_INK_2, "Otro"))

# Catálogo SAT de Formas de Pago (códigos más comunes)
_FORMA_PAGO_DESC: dict[str, str] = {
    "01": "Efectivo", "02": "Cheque nominativo", "03": "Transferencia electrónica",
    "04": "Tarjeta de crédito", "05": "Monedero electrónico",
    "06": "Dinero electrónico", "08": "Vales de despensa",
    "12": "Dación en pago", "13": "Pago por subrogación",
    "14": "Pago por consignación", "15": "Condonación",
    "17": "Compensación", "23": "Novación", "24": "Confusión",
    "25": "Remisión de deuda", "26": "Prescripción",
    "27": "A satisfacción del acreedor", "28": "Tarjeta de débito",
    "29": "Tarjeta de servicios", "30": "Aplicación de anticipos",
    "31": "Intermediario pagos", "99": "Por definir",
}


def _forma_pago_txt(codigo: str) -> str:
    """Devuelve 'NN - Descripción' o solo el código si no está en el catálogo."""
    if not codigo:
        return ""
    desc = _FORMA_PAGO_DESC.get(codigo, "")
    return f"{codigo} - {desc}" if desc else codigo


# Catálogo SAT de UsoCFDI (tipo de gasto/deducción). Códigos vigentes 4.0.
_USO_CFDI_DESC: dict[str, str] = {
    "G01": "Adquisición de mercancías",
    "G02": "Devoluciones, descuentos o bonificaciones",
    "G03": "Gastos en general",
    "I01": "Construcciones",
    "I02": "Mobiliario y equipo de oficina",
    "I03": "Equipo de transporte",
    "I04": "Equipo de cómputo y accesorios",
    "I05": "Dados, troqueles, moldes, matrices y herramental",
    "I06": "Comunicaciones telefónicas",
    "I07": "Comunicaciones satelitales",
    "I08": "Otra maquinaria y equipo",
    "D01": "Honorarios médicos, dentales y gastos hospitalarios",
    "D02": "Gastos médicos por incapacidad o discapacidad",
    "D03": "Gastos funerales",
    "D04": "Donativos",
    "D05": "Intereses reales de créditos hipotecarios",
    "D06": "Aportaciones voluntarias al SAR",
    "D07": "Primas por seguros de gastos médicos",
    "D08": "Gastos de transportación escolar obligatoria",
    "D09": "Depósitos en cuentas para el ahorro / pensiones",
    "D10": "Pagos por servicios educativos (colegiaturas)",
    "S01": "Sin efectos fiscales",
    "CP01": "Pagos",
    "CN01": "Nómina",
    "P01": "Por definir",
}


def _uso_cfdi_txt(codigo: str) -> str:
    """Devuelve 'G03 - Gastos en general' o solo el código si no está en el catálogo."""
    if not codigo:
        return ""
    desc = _USO_CFDI_DESC.get(codigo.upper(), "")
    return f"{codigo.upper()} - {desc}" if desc else codigo.upper()


# Catálogo SAT de TipoDeComprobante: viene del atributo del propio CFDI.
# "R" no lo emite el parser de comprobantes (los CFDI de retenciones usan otro
# esquema, sin TipoDeComprobante), queda en el catálogo para no perder la clave
# si algún día entra por esa vía.
_TIPO_CFDI_DESC: dict[str, str] = {
    "I": "Ingreso",
    "E": "Egreso",
    "T": "Traslado",
    "N": "Nómina",
    "P": "Pago",
    "R": "Retenciones e Información de Pagos",
}


def _tipo_cfdi_txt(tipo: str) -> str:
    """Devuelve 'I - Ingreso' o solo la clave si no está en el catálogo."""
    t = (tipo or "").strip().upper()
    if not t:
        return ""
    desc = _TIPO_CFDI_DESC.get(t, "")
    return f"{t} - {desc}" if desc else t


# Catálogo SAT c_TipoRelacion (Anexo 20). El 07 es el que usa el comprobante
# que aplica un anticipo previamente facturado.
_TIPO_RELACION_DESC: dict[str, str] = {
    "01": "Nota de crédito de los documentos relacionados",
    "02": "Nota de débito de los documentos relacionados",
    "03": "Devolución de mercancía sobre facturas o traslados previos",
    "04": "Sustitución de los CFDI previos",
    "05": "Traslados de mercancías facturados previamente",
    "06": "Factura generada por los traslados previos",
    "07": "CFDI por aplicación de anticipo",
}


def _rel_txt(rels: list) -> str:
    """'07 - CFDI por aplicación de anticipo: UUID' por cada relación, unidas
    con ' | '. rels es [(TipoRelacion, UUID)] tal cual viene del XML."""
    partes = []
    for tipo, uuid in rels:
        desc = _TIPO_RELACION_DESC.get(tipo, "")
        etiqueta = f"{tipo} - {desc}" if desc else (tipo or "¿?")
        partes.append(f"{etiqueta}: {uuid}")
    return " | ".join(partes)


def _leyenda_relacion(tipo_cfdi: str, rels: list) -> str:
    """Leyenda fiscal para egresos y para la aplicación de anticipos.

    El procedimiento de anticipos (emitir el anticipo, luego el comprobante por
    el total y relacionarlo con TipoRelacion 07) está en el Apéndice 6 de la
    Guía de llenado del CFDI 4.0 (Anexo 20 de la RMF). Un egreso sin
    CfdiRelacionados no permite identificar qué comprobante reduce, así que se
    marca como observación.
    """
    tipos = {t for t, _ in rels if t}
    if "07" in tipos:
        return ("Aplicación de anticipo (TipoRelacion 07). Procedimiento: Guía de "
                "llenado del CFDI 4.0, Apéndice 6 — el comprobante se relaciona "
                "con el CFDI de anticipo para no duplicar el ingreso.")
    if tipo_cfdi != "E":
        return ""
    if not rels:
        return ("Egreso SIN CFDI relacionado: no se identifica el comprobante que "
                "reduce ni el anticipo que aplica (Guía de llenado CFDI 4.0, "
                "Apéndice 6 y catálogo c_TipoRelacion).")
    if tipos & {"01", "03"}:
        return ("Egreso que reduce el importe del CFDI relacionado "
                "(nota de crédito / devolución, TipoRelacion 01 o 03).")
    if "04" in tipos:
        return "Egreso que sustituye al CFDI relacionado (TipoRelacion 04)."
    return "Egreso relacionado: " + ", ".join(sorted(tipos))


MONEY_FMT = '#,##0.00'
DATE_FMT  = 'DD/MM/YYYY'

# Rejilla recesiva: hairline del color de línea, no negro. La rejilla no debe
# competir con los datos.
_S_HAIR = Side(style="thin", color=C_LINEA)
THIN = Border(left=_S_HAIR, right=_S_HAIR, top=_S_HAIR, bottom=_S_HAIR)

FUENTE = "Segoe UI"


def _cell(ws, row: int, col: int, value=None, *,
          bold: bool = False, bg: Optional[str] = None,
          fg: str = C_INK, align: str = "left",
          fmt: Optional[str] = None, wrap: bool = False,
          size: int = 9, accent: Optional[str] = None) -> None:
    """`accent` pinta un borde izquierdo grueso de ese color — es la marca de
    identidad (categoría) sin sacrificar la legibilidad del texto, que se queda
    sobre un relleno claro."""
    c = ws.cell(row=row, column=col, value=value)
    c.font      = Font(name=FUENTE, bold=bold, size=size, color=fg)
    c.alignment = Alignment(horizontal=align, vertical="center",
                            wrap_text=wrap)
    c.border    = (Border(left=Side(style="thick", color=accent),
                          right=_S_HAIR, top=_S_HAIR, bottom=_S_HAIR)
                   if accent else THIN)
    if bg:
        c.fill = PatternFill("solid", fgColor=bg)
    if fmt:
        c.number_format = fmt


def _header_row(ws, row: int, labels: list[str], color: str,
                sizes: Optional[list[int]] = None) -> None:
    borde = Side(style="thin", color=color)
    for col, lbl in enumerate(labels, 1):
        c = ws.cell(row=row, column=col, value=lbl)
        c.font      = Font(name=FUENTE, bold=True, color="FFFFFF", size=9)
        c.fill      = PatternFill("solid", fgColor=color)
        c.alignment = Alignment(horizontal="center", vertical="center",
                                wrap_text=True)
        # Separadores del mismo color que el fondo: el encabezado se lee como
        # una barra continua, no como una cuadrícula.
        c.border    = Border(left=borde, right=borde, top=borde, bottom=borde)
    ws.row_dimensions[row].height = 26
    if sizes:
        for col, w in enumerate(sizes, 1):
            ws.column_dimensions[get_column_letter(col)].width = w


def _alto_fila(*campos: tuple[str, int], base: int = 12) -> int:
    """Alto de fila que SÍ contempla el texto que envuelve.

    Contar solo los '\\n' dejaba las descripciones largas encimadas entre
    renglones: una descripción SPEI de 120 caracteres en una columna de 50
    ocupa 3 líneas visuales sin traer un solo salto de línea.
    """
    n = 1
    for texto, ancho in campos:
        if not texto or ancho <= 0:
            continue
        lineas = sum(max(1, -(-len(seg) // ancho))
                     for seg in str(texto).split("\n"))
        n = max(n, lineas)
    return max(16, n * base + 4)


def _banda(ws, row: int, ancho: str, texto: str,
           color: str = C_HDR_S) -> None:
    """Encabezado de sección dentro de una hoja."""
    ws.merge_cells(f"A{row}:{ancho}{row}")
    c = ws.cell(row=row, column=1, value=texto)
    c.font      = Font(name=FUENTE, bold=True, size=9, color="FFFFFF")
    c.fill      = PatternFill("solid", fgColor=color)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[row].height = 18


def _titulo(ws, ancho: str, texto: str, color: str,
            subtitulo: str = "") -> None:
    """Barra de título de la hoja + línea de subtítulo opcional."""
    ws.merge_cells(f"A1:{ancho}1")
    c = ws["A1"]
    c.value     = texto
    c.font      = Font(name=FUENTE, bold=True, size=14, color="FFFFFF")
    c.fill      = PatternFill("solid", fgColor=color)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 30
    if subtitulo:
        ws.merge_cells(f"A2:{ancho}2")
        s = ws["A2"]
        s.value     = subtitulo
        s.font      = Font(name=FUENTE, size=9, color=C_INK_2)
        s.fill      = PatternFill("solid", fgColor=C_WHITE)
        s.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[2].height = 16


# ── Hoja: Movimientos bancarios ───────────────────────────────────────────────

# La columna A es el ID ESTABLE (llave de sync con la BD Firebird). Todas las
# demás columnas se corren una posición a la derecha; el macro PapelTrabajo y las
# fórmulas del TABLERO ya están ajustados a este corrimiento (UUID=H, TIPO MATCH=Q).
HDRS_BANCO = [
    "ID",
    "FECHA", "FOLIO/REF", "DESCRIPCION", "RETIRO", "DEPOSITO", "SALDO",
    "UUID CFDI / N.FACTURA", "UUID REP", "SERIE", "FOLIO CFDI", "RFC",
    "NOMBRE / RAZON SOCIAL",
    "FECHA EMISION", "FECHA PAGO", "ESTATUS", "TIPO MATCH",
]
WIDTHS_BANCO = [8, 12, 14, 50, 14, 14, 16, 38, 38, 10, 12, 16, 34, 14, 14, 17, 21]


def _hoja_banco(wb, nombre: str,
                conciliados: list[MovConciliado],
                sin_cfdi: list[MovBanco],
                color_banco: str = C_HDR,
                numero_cuenta: str = "") -> int:
    """Devuelve el número de renglones de datos, que el TABLERO necesita para
    armar los rangos de sus fórmulas."""
    ws = wb.create_sheet(nombre[:31])
    # La rejilla de Excel compite con los datos; los bordes hairline ya dan
    # estructura suficiente.
    ws.sheet_view.showGridLines = False

    # Datos: (MovBanco, es_conciliado, revisar, tipo_match)
    filas: list[tuple[MovBanco, bool, bool, str]] = (
        [(mc.banco, True,  mc.revisar, mc.tipo_match) for mc in conciliados] +
        [(mb,       False, False,      "")            for mb in sin_cfdi]
    )
    # NO reordenar: mostrar los movimientos tal cual aparecen en el estado de
    # cuenta. `orden` se estampa en la conciliación (posición original). El sort
    # es estable, así que si `orden` no viniera (todos 0) se conserva el orden
    # de construcción anterior.
    filas.sort(key=lambda f: getattr(f[0], "orden", 0))
    n_datos = len(filas)
    n_conc  = sum(1 for _, c, _, _ in filas if c)
    pct_tot = round(n_conc / n_datos * 100, 1) if n_datos else 0.0

    # Título + cifra principal arriba: el dato que se busca primero no debe
    # exigir bajar hasta el pie de la hoja. Va como FÓRMULA para que siga al día
    # en cuanto el usuario concilie a mano.
    # Sin porcentaje de conciliación (a solicitud): solo conteos.
    _sub = (f'=COUNTIF($Q$4:$Q${n_datos + 3},"<>")&" de {n_datos} movimientos conciliados'
            f'   ·   "&({n_datos}-COUNTIF($Q$4:$Q${n_datos + 3},"<>"))&" sin CFDI"')
    _titulo_txt = f"Conciliación — {nombre}"
    if numero_cuenta:
        _titulo_txt += f"   ·   Cuenta: {numero_cuenta}"
    _titulo(ws, "Q", _titulo_txt, color_banco,
            subtitulo=_sub if n_datos else "Sin movimientos")

    _header_row(ws, 3, HDRS_BANCO, color_banco, WIDTHS_BANCO)

    for idx, (mb, conc, revisar, tipo_match) in enumerate(filas):
        row = idx + 4
        alt = C_ALT if idx % 2 else C_WHITE

        es_suma = mb.estatus == "Conciliado SUMA"
        n_suma  = (len(mb.uuid_suma.split(",")) + 1) if mb.uuid_suma else 0

        # El bloque CFDI se tiñe SOLO cuando es la excepción (sin CFDI). Teñir
        # también lo normal ahogaba justo lo que se anda buscando.
        if conc and es_suma:
            bg_e     = alt
            est_bg, est_fg = C_REVISAR, C_REVISAR_INK
            est_txt  = f"Suma de {n_suma} CFDIs"
            est_bold = True
        elif conc and revisar:
            bg_e     = alt
            est_bg, est_fg = C_REVISAR, C_REVISAR_INK
            est_txt  = "Revisar"
            est_bold = True
        elif conc:
            bg_e     = alt
            est_bg, est_fg = C_CONC, C_CONC_INK
            est_txt  = "Conciliado"
            est_bold = True
        else:
            bg_e     = C_SIN
            est_bg, est_fg = C_SIN, C_SIN_INK
            est_txt  = mb.estatus or "Sin CFDI"
            est_bold = True

        # UUID: listar SIEMPRE todos los CFDIs relacionados (primero + adicionales).
        # No basta con `es_suma`: la Fase 4 (PPD c/ Complemento) deja estatus
        # "Conciliado" pero un REP puede liquidar varias facturas, y todas deben
        # verse en la columna. Igual aplica a Global Mensual (Fase 5).
        uuid_disp = mb.uuid
        if mb.uuid_suma:
            uuid_disp = mb.uuid + "\n+ " + mb.uuid_suma.replace(",", "\n+ ")

        desc = (mb.descripcion or "").replace(" | ", "\n")
        ws.row_dimensions[row].height = _alto_fila(
            (desc, WIDTHS_BANCO[2]), (uuid_disp, WIDTHS_BANCO[6]),
            (mb.nombre or "", WIDTHS_BANCO[11]))

        # Col 1 (A) — ID ESTABLE de la fila (de la BD). Se estampa en el MovBanco
        # como `id_fila`; si no viniera, queda vacío (no rompe nada).
        _cell(ws, row, 1,  getattr(mb, "id_fila", None) or None, bg=alt,
              align="center", fg=C_INK_2)
        _cell(ws, row, 2,  mb.fecha,        bg=alt, align="center")
        _cell(ws, row, 3,  mb.folio,        bg=alt, align="center")
        _cell(ws, row, 4,  desc,            bg=alt, wrap=True)
        # Retiro/depósito son dos sentidos de un flujo, no bueno/malo: tinte
        # neutro cálido vs frío en vez de rojo/verde de estatus.
        _cell(ws, row, 5,  mb.retiro    or None, fmt=MONEY_FMT,
              bg="FDF3EE" if mb.retiro   else alt, align="right")
        _cell(ws, row, 6,  mb.deposito  or None, fmt=MONEY_FMT,
              bg="EDF6F1" if mb.deposito else alt, align="right")
        _cell(ws, row, 7,  mb.saldo     or None, fmt=MONEY_FMT, bg=C_SAL,
              align="right", fg=C_INK_2)
        _cell(ws, row, 8,  uuid_disp,           bg=bg_e, wrap=True)
        _cell(ws, row, 9,  mb.uuid_rep or None, bg=bg_e, wrap=True)
        _cell(ws, row, 10, mb.serie_cfdi or None, bg=bg_e, align="center")
        _cell(ws, row, 11, mb.folio_cfdi or None, bg=bg_e, align="center")
        _cell(ws, row, 12, mb.rfc,              bg=bg_e, align="center")
        _cell(ws, row, 13, mb.nombre,           bg=bg_e, wrap=True)
        _cell(ws, row, 14, mb.fecha_cfdi,       bg=bg_e, align="center")
        _cell(ws, row, 15, mb.fecha_pago_cfdi or None, bg=bg_e, align="center")
        _cell(ws, row, 16, est_txt,             bg=est_bg, align="center",
              bold=est_bold, fg=est_fg)

        # Col 17 (Q) — TIPO MATCH. El nombre de la fase va en texto legible sobre
        # relleno claro; el NIVEL DE CONFIANZA lo marca el borde de color.
        hue, _lvl = _tier(tipo_match)
        _cell(ws, row, 17, tipo_match or None, bg=C_NEUTRO if tipo_match else alt,
              align="left", bold=False, fg=C_INK if tipo_match else C_INK_2,
              accent=hue or None)

        # Col 19 (S) — MES numérico, columna auxiliar OCULTA. Existe para que el
        # bloque de porcentajes sea fórmula viva: la fecha (col B) es texto y su
        # formato cambia por banco ("04/JUN./2026" en Inbursa), así que no se
        # puede sacar el mes con fórmula de forma confiable. Va en la 19 (fuera de
        # la tabla visible) y queda fuera de las columnas 8-17 que toca la macro.
        ws.cell(row=row, column=19,
                value=(mb.mes or _mes_fecha(mb.fecha)) or None)

    ws.freeze_panes = "E4"
    ws.auto_filter.ref = f"A3:{get_column_letter(len(HDRS_BANCO))}{n_datos + 3}"
    ws.cell(row=3, column=19, value="MES (aux)")
    ws.column_dimensions["S"].hidden = True

    # Rangos que usan las fórmulas vivas (aquí y en el TABLERO). Un renglón se
    # cuenta como conciliado si TIPO MATCH (col P) NO está vacío — la macro
    # PapelTrabajo escribe "Manual" ahí, así que las conciliaciones a mano
    # entran solas sin tocar nada más.
    _f, _u = 4, n_datos + 3
    _R_TM  = f"$Q${_f}:$Q${_u}"   # TIPO MATCH (col Q tras insertar ID en A)
    _R_MES = f"$S${_f}:$S${_u}"   # MES aux (col S tras el corrimiento)

    # ── Porcentaje de conciliación por MES (de este estado de cuenta) ─────────
    _MESES_NOM = {1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
                  5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
                  9: "Septiembre", 10: "Octubre", 11: "Noviembre",
                  12: "Diciembre", 0: "Sin fecha"}
    por_mes: dict[int, list[int]] = {}   # mes → [conciliados, total]
    for _mb, _conc, _r, _tm in filas:
        _mnum = _mb.mes or _mes_fecha(_mb.fecha)
        _acc = por_mes.setdefault(_mnum, [0, 0])
        _acc[1] += 1
        if _conc:
            _acc[0] += 1

    # Tabla real (mes · conciliados · total · %) en vez de renglones fusionados:
    # se puede leer en columna y comparar de un vistazo.
    pct_hdr = n_datos + 6
    _banda(ws, pct_hdr, "P", "CONCILIACIÓN POR MES")

    def _fila_pct(r: int, mes: str, cn, tn, pct, bg: str, fg: str,
                  bold: bool = False) -> None:
        # La etiqueta ocupa A:B — "Septiembre" o "TOTAL SANTANDER" no caben en A.
        ws.merge_cells(f"A{r}:B{r}")
        _cell(ws, r, 1, mes, bg=bg, bold=bold, fg=fg)
        _cell(ws, r, 2, None, bg=bg)
        _cell(ws, r, 3, cn,  bg=bg, bold=bold, fg=fg, align="right")
        _cell(ws, r, 4, tn,  bg=bg, bold=bold,
              fg=fg if bold else C_INK_2, align="right")
        _cell(ws, r, 5, pct, bg=bg, bold=True, fg=fg, align="right",
              fmt="0.0%" if pct is not None else None)
        ws.row_dimensions[r].height = 15

    pr = pct_hdr + 1
    _fila_pct(pr, "MES", "CONCILIADOS", "TOTAL", None, C_NEUTRO, C_INK_2, True)
    _cell(ws, pr, 5, "% CONCILIADO", bg=C_NEUTRO, bold=True, fg=C_INK_2,
          align="right", size=8)

    # FÓRMULAS, no números fijos: si el usuario concilia a mano con la macro
    # (o edita la col P), estos porcentajes se recalculan solos al instante.
    fila_ini_meses = pr + 1
    for _mnum in sorted(por_mes):
        pr += 1
        _cn, _tn = por_mes[_mnum]
        _pct = round(_cn / _tn * 100, 1) if _tn else 0.0
        # Semáforo por valor calculado al exportar; el número es fórmula viva.
        if   _pct >= 90: _bg, _fg = C_CONC,    C_CONC_INK
        elif _pct >= 60: _bg, _fg = C_REVISAR, C_REVISAR_INK
        else:            _bg, _fg = C_SIN,     C_SIN_INK
        _fila_pct(pr, _MESES_NOM.get(_mnum, "Sin fecha"),
                  f'=COUNTIFS({_R_MES},{_mnum},{_R_TM},"<>")',
                  f"=COUNTIFS({_R_MES},{_mnum})",
                  f"=IF(D{pr}=0,0,C{pr}/D{pr})", C_WHITE, C_INK)
        _cell(ws, pr, 5, f"=IF(D{pr}=0,0,C{pr}/D{pr})", bg=_bg, bold=True,
              fg=_fg, align="right", fmt="0.0%")
    fila_fin_meses = pr

    pr += 1
    _fila_pct(pr, f"TOTAL {nombre}",
              f"=SUM(C{fila_ini_meses}:C{fila_fin_meses})",
              f"=SUM(D{fila_ini_meses}:D{fila_fin_meses})",
              f"=IF(D{pr}=0,0,C{pr}/D{pr})", C_HDR_S, "FFFFFF", True)
    ws.row_dimensions[pr].height = 17

    # ── Leyenda al pie, agrupada por NIVEL DE CONFIANZA ───────────────────────
    # El color agrupa (4 niveles validados); el texto distingue las 7 fases.
    def _n(tm: str) -> int:
        return sum(1 for mc in conciliados if mc.tipo_match == tm)
    n_grupo = sum(1 for mc in conciliados if mc.tipo_match.startswith("Grupo:"))

    pie_row = pr + 2
    _banda(ws, pie_row, "P", "CÓMO LEER ESTA HOJA")

    fila = pie_row

    def _nivel(hue: str, titulo: str, detalle: str) -> None:
        nonlocal fila
        fila += 1
        # A:B para la etiqueta — "Con complemento" no cabe en A sola.
        oscuro = _TIER_OSCURO.get(hue, C_HDR_S)
        ws.merge_cells(f"A{fila}:B{fila}")
        _cell(ws, fila, 1, titulo, bg=oscuro, bold=True, fg="FFFFFF")
        _cell(ws, fila, 2, None, bg=oscuro)
        ws.merge_cells(f"C{fila}:P{fila}")
        _cell(ws, fila, 3, detalle, bg=C_NEUTRO, fg=C_INK_2)
        ws.row_dimensions[fila].height = 15

    _nivel(TIER_DIRECTO, "Directo",
           f"Fase 1 Exacto ({_n('Exacto')}) — cruce 1:1 por RFC/nombre + importe, "
           f"mismo mes (±$1 MXN). Es la evidencia más fuerte.")
    _nivel(TIER_COMPUESTO, "Compuesto",
           f"Fase 2 Multi-Movimiento ({_n('Multi-Movimiento')}): varios movimientos suman UN CFDI  ·  "
           f"Fase 3 Combinatorio ({_n('Combinatorio')}): un movimiento paga varios CFDIs  ·  "
           f"Fase 5 Agrupado por RFC ({_n('Global Mensual')} auto / {n_grupo} grupos). "
           f"Verificar la asignación individual.")
    _nivel(TIER_COMPLEM, "Con complemento",
           f"Fase 4 PPD c/ Complemento ({_n('PPD c/ Complemento')}) — REP (tipo P) cruzado por "
           f"MontoPagado + RFC/nombre, con FechaPago del REP en el mismo mes del banco o a ±5 días; "
           f"liquida sus facturas.")
    _nivel(TIER_DEBIL, "Débil",
           f"Fase 6 PPD Sin Complemento ({_n('PPD Sin Complemento')}): falta el REP, cruza por RFC/nombre  ·  "
           f"Fase 7 Solo Importe ({_n('Solo Importe')}): SOLO importe del mismo mes, sin validar al tercero. "
           f"Ambas exigen revisión.")

    fila += 1
    ws.merge_cells(f"A{fila}:B{fila}")
    _cell(ws, fila, 1, "Sin CFDI", bg=C_SIN, bold=True, fg=C_SIN_INK)
    _cell(ws, fila, 2, None, bg=C_SIN)
    ws.merge_cells(f"C{fila}:P{fila}")
    _cell(ws, fila, 3, "No se encontró un CFDI que coincida con el movimiento. "
                       "Es lo único resaltado en el cuerpo de la tabla.",
          bg=C_SIN, fg=C_SIN_INK)
    ws.row_dimensions[fila].height = 15

    fila += 2
    ws.merge_cells(f"A{fila}:P{fila}")
    _cell(ws, fila, 1,
          "Trazabilidad — col G: UUID del CFDI (factura tipo I/N; si es suma se listan los adicionales)   ·   "
          "col H: UUID del complemento de pago (REP tipo P), solo cuando el cruce se hizo vía complemento (Fase 4).",
          bg=C_WHITE, fg=C_INK_2, size=8)
    ws.row_dimensions[fila].height = 14

    return n_datos


# ── Hoja: TABLERO (todo con fórmulas, se actualiza en vivo) ──────────────────

def _hoja_tablero(wb, hojas: list[tuple[str, int]]) -> None:
    """Tablero de control con SOLO fórmulas.

    `hojas` = [(nombre_hoja_banco, n_datos), ...].

    Por qué existe: el resto del reporte se calcula en Python y se escribe como
    números fijos, así que en cuanto el contador concilia a mano (macro
    `PapelTrabajo`, que escribe "Manual" en la col P) esos totales quedan
    mintiendo. Aquí todo es COUNTIF/COUNTIFS sobre las hojas de banco: al
    guardar, Excel recalcula y el avance real se ve al instante.

    Un renglón cuenta como conciliado si su TIPO MATCH (col P) no está vacío —
    mismo criterio para lo automático y lo manual.
    """
    ws = wb.create_sheet("TABLERO")
    ws.sheet_view.showGridLines = False
    _titulo(ws, "F", "Tablero de conciliación", C_HDR,
            subtitulo="Todo en esta hoja son fórmulas: al conciliar a mano se "
                      "actualiza solo. Si no cambia, pulsa F9.")

    for col, w in zip("ABCDEF", (30, 14, 14, 14, 16, 34)):
        ws.column_dimensions[col].width = w

    def _ref(hoja: str, col: str, n: int) -> str:
        """Rango absoluto a una hoja de banco, con comillas si el nombre las
        necesita ('SANTANDER MAY'!...)."""
        h = f"'{hoja}'" if any(c in hoja for c in " -()") else hoja
        return f"{h}!${col}$4:${col}${n + 3}"

    # ── Avance por banco ─────────────────────────────────────────────────────
    _banda(ws, 4, "F", "AVANCE POR ESTADO DE CUENTA")
    r = 5
    for c, lbl in zip(range(1, 7),
                      ("ESTADO DE CUENTA", "MOVIMIENTOS", "CONCILIADOS",
                       "PENDIENTES", "% CONCILIADO", "")):
        _cell(ws, r, c, lbl or None, bg=C_NEUTRO, bold=True, fg=C_INK_2,
              align="left" if c == 1 else "right", size=8)

    ini = r + 1
    for hoja, n in hojas:
        r += 1
        tm = _ref(hoja, "Q", n)   # TIPO MATCH (col Q tras insertar ID en A)
        _cell(ws, r, 1, hoja, bg=C_WHITE)
        _cell(ws, r, 2, n, bg=C_WHITE, align="right", fg=C_INK_2)
        _cell(ws, r, 3, f'=COUNTIF({tm},"<>")', bg=C_WHITE, align="right")
        _cell(ws, r, 4, f"=B{r}-C{r}", bg=C_WHITE, align="right")
        _cell(ws, r, 5, f"=IF(B{r}=0,0,C{r}/B{r})", bg=C_WHITE, align="right",
              bold=True, fmt="0.0%")
        # Semáforo por fórmula: se repinta el texto, no el relleno (openpyxl no
        # hace formato condicional de color de fondo sin definir una regla).
        _cell(ws, r, 6, f'=IF(E{r}>=0.9,"Al corriente",'
                        f'IF(E{r}>=0.6,"Requiere atención","Atrasado"))',
              bg=C_WHITE, fg=C_INK_2)
        ws.row_dimensions[r].height = 15
    fin = r

    r += 1
    _cell(ws, r, 1, "TOTAL", bg=C_HDR_S, bold=True, fg="FFFFFF")
    for c, f in ((2, f"=SUM(B{ini}:B{fin})"), (3, f"=SUM(C{ini}:C{fin})"),
                 (4, f"=SUM(D{ini}:D{fin})")):
        _cell(ws, r, c, f, bg=C_HDR_S, bold=True, fg="FFFFFF", align="right")
    _cell(ws, r, 5, f"=IF(B{r}=0,0,C{r}/B{r})", bg=C_HDR_S, bold=True,
          fg="FFFFFF", align="right", fmt="0.0%")
    _cell(ws, r, 6, None, bg=C_HDR_S)
    ws.row_dimensions[r].height = 17
    fila_total = r

    # ── Desglose por nivel de confianza y fase ───────────────────────────────
    r += 2
    _banda(ws, r, "F", "DESGLOSE POR NIVEL DE CONFIANZA")
    r += 1
    for c, lbl in zip(range(1, 6),
                      ("NIVEL / FASE", "MOVIMIENTOS", "% DEL TOTAL", "", "")):
        _cell(ws, r, c, lbl or None, bg=C_NEUTRO, bold=True, fg=C_INK_2,
              align="left" if c == 1 else "right", size=8)

    def _suma_por(criterio: str) -> str:
        """Suma el mismo COUNTIF sobre TODAS las hojas de banco."""
        return "=" + "+".join(
            f'COUNTIF({_ref(h, "Q", n)},{criterio})' for h, n in hojas
        ) if hojas else "=0"

    niveles = [
        (TIER_DIRECTO,   "Directo",         ["Exacto"]),
        (TIER_COMPUESTO, "Compuesto",       ["Multi-Movimiento", "Combinatorio",
                                             "Global Mensual", "Grupo:*"]),
        (TIER_COMPLEM,   "Con complemento", ["PPD c/ Complemento",
                                             "Recepción de Pago"]),
        (TIER_DEBIL,     "Débil",           ["PPD Sin Complemento",
                                             "Solo Importe"]),
        (C_INK_2,        "Manual",          ["Manual"]),
    ]
    for hue, nom, tipos in niveles:
        r += 1
        expr = "=" + "+".join(
            f'COUNTIF({_ref(h, "Q", n)},"{t}")'
            for h, n in hojas for t in tipos
        ) if hojas else "=0"
        # Swatch sólido con el paso oscuro del hue (mismo criterio que la
        # leyenda de la hoja de banco: el borde fino se pierde en la col A).
        _cell(ws, r, 1, nom, bg=_TIER_OSCURO.get(hue, C_HDR_S), bold=True,
              fg="FFFFFF")
        _cell(ws, r, 2, expr, bg=C_WHITE, align="right")
        _cell(ws, r, 3, f"=IF($C${fila_total}=0,0,B{r}/$C${fila_total})",
              bg=C_WHITE, align="right", fmt="0.0%")
        _cell(ws, r, 4, None, bg=C_WHITE)
        _cell(ws, r, 5, None, bg=C_WHITE)
        ws.row_dimensions[r].height = 15

    r += 2
    ws.merge_cells(f"A{r}:F{r}")
    _cell(ws, r, 1,
          "«Manual» son los cruces que hiciste con la macro PapelTrabajo "
          "(Alt+F8 → ConciliarManuales): pega el UUID en la columna G de la hoja "
          "de banco y ejecútala; este tablero lo cuenta solo.",
          bg=C_WHITE, fg=C_INK_2, size=8)
    ws.row_dimensions[r].height = 26

    # El tablero es lo primero que se debe ver.
    wb.move_sheet("TABLERO", offset=-(len(wb.sheetnames) - 1))


# ── Hoja: CFDIs Recibidos / Emitidos ─────────────────────────────────────────

HDRS_RECIBIDOS = [
    "UUID", "SERIE", "FOLIO CFDI", "FECHA EMISION", "RFC EMISOR",
    "RAZON EMISOR", "SUBTOTAL", "DESCUENTO", "IVA 16%",
    "IVA RETENIDO", "ISR RETENIDO", "TOTAL",
    "PERIODO",
    "BANCO", "FOLIO / REF BANCARIA", "FECHA PAGO BANCO", "FECHA PAGO NOMINA", "ESTATUS",
    "VALIDACION FISCAL", "OBSERVACIONES FISCALES",
    "TIPO DE CFDI",                                                  # col 21 — TipoDeComprobante del CFDI
    "FORMA DE PAGO",                                                 # col 22 — Art.27 LIVA
    "MÉTODO PAGO", "COMPLEMENTO DE PAGO", "UUID COMPLEMENTO P",
    "FECHA PAGO COMPLEMENTO",                                        # col 26 — del XML de pago, para cruce íntegro
    "MONTO PAGADO EN COMPLEMENTO DE PAGO",
    "TIPO DE GASTO (USO CFDI)",                                      # col 28 — UsoCFDI G01/G03…
    "UUID RELACIONADO (CfdiRelacionados)",                           # col 29 — tipo E y anticipos
    "LEYENDA FISCAL (EGRESO / ANTICIPO)",                            # col 30
    "CLAVE PROD/SERV SAT",                                           # col 31 — todos los conceptos, " | "
    "CONCEPTOS DEL CFDI",                                            # col 32 — todos los conceptos, " | "
    "BASE IVA 16%",                                                  # col 33 — aux IVA desglose (oculta)
    "BASE IVA 8%",                                                   # col 34 — aux IVA desglose (oculta)
    "IVA IMPORTE 8%",                                                # col 35 — aux IVA desglose (oculta)
    "BASE IVA 0%",                                                   # col 36 — aux IVA desglose (oculta)
    "BASE EXENTA",                                                   # col 37 — aux IVA desglose (oculta)
    "TIPO TASA",                                                     # col 38 — aux: "16%","8%","0%","Exento","Mixta" (oculta)
]
#                  1   2   3   4   5   6   7   8   9  10  11  12  13  14  15  16  17  18  19  20  21  22  23  24  25  26  27  28  29  30  31  32  33  34  35  36  37  38
WIDTHS_RECIBIDOS = [38,  8, 10, 14, 18, 40, 14, 12, 12, 12, 12, 14,  8, 16, 20, 16, 16, 15, 16, 48, 34, 22, 10, 16, 38, 16, 20, 30, 46, 58, 24, 60, 14, 14, 14, 14, 14, 10]

HDRS_EMITIDOS = [
    "UUID", "SERIE", "FOLIO CFDI", "FECHA EMISION", "RFC RECEPTOR",
    "RAZON RECEPTOR", "SUBTOTAL", "DESCUENTO", "IVA 16%",
    "IVA RETENIDO", "ISR RETENIDO", "TOTAL",
    "PERIODO",
    "BANCO", "FOLIO / REF BANCARIA", "FECHA PAGO BANCO", "FECHA PAGO NOMINA", "ESTATUS",
    "VALIDACION FISCAL", "OBSERVACIONES FISCALES",
    "TIPO DE CFDI",                                                  # col 21 — TipoDeComprobante del CFDI
    "FORMA DE PAGO",                                                 # col 22 — Art.27 LIVA
    "MÉTODO PAGO", "COMPLEMENTO DE PAGO", "UUID COMPLEMENTO P",
    "FECHA PAGO COMPLEMENTO",                                        # col 26 — del XML de pago, para cruce íntegro
    "MONTO PAGADO EN COMPLEMENTO DE PAGO",
    "TIPO DE GASTO (USO CFDI)",                                      # col 28 — UsoCFDI G01/G03…
    "UUID RELACIONADO (CfdiRelacionados)",                           # col 29 — tipo E y anticipos
    "LEYENDA FISCAL (EGRESO / ANTICIPO)",                            # col 30
    "CLAVE PROD/SERV SAT",                                           # col 31 — todos los conceptos, " | "
    "CONCEPTOS DEL CFDI",                                            # col 32 — todos los conceptos, " | "
    "BASE IVA 16%",                                                  # col 33 — aux IVA desglose (oculta)
    "BASE IVA 8%",                                                   # col 34 — aux IVA desglose (oculta)
    "IVA IMPORTE 8%",                                                # col 35 — aux IVA desglose (oculta)
    "BASE IVA 0%",                                                   # col 36 — aux IVA desglose (oculta)
    "BASE EXENTA",                                                   # col 37 — aux IVA desglose (oculta)
    "TIPO TASA",                                                     # col 38 — aux: "16%","8%","0%","Exento","Mixta" (oculta)
]
WIDTHS_EMITIDOS = WIDTHS_RECIBIDOS

# ── Hoja: Complementos de Pago (tipo P) ──────────────────────────────────────
HDRS_COMP_PAGO = [
    "UUID COMPLEMENTO P", "FECHA PAGO", "MONTO PAGADO XML",
    "MONTO BANCO", "DIFERENCIA",
    "RFC EMISOR", "RAZON EMISOR", "RFC RECEPTOR", "RAZON RECEPTOR",
    "Nº FACTURAS PAGADAS", "UUIDs FACTURAS PAGADAS",
    "FACTURAS EN CARTERA", "BANCO", "ESTATUS",
    "VALIDACION FISCAL", "OBSERVACIONES FISCALES",
]
WIDTHS_COMP_PAGO = [38, 14, 16, 16, 14, 18, 38, 18, 38, 10, 60, 18, 18, 14, 16, 48]


def _tipo_tasa(cfdi) -> str | None:
    has_16 = getattr(cfdi, "iva_base_16", 0) > 0.005
    has_8  = getattr(cfdi, "iva_base_8", 0) > 0.005
    has_0  = getattr(cfdi, "iva_base_0", 0) > 0.005
    has_ex = getattr(cfdi, "iva_exento_base", 0) > 0.005
    n = has_16 + has_8 + has_0 + has_ex
    if n == 0:
        return None
    if n > 1:
        return "Mixta"
    if has_16: return "16%"
    if has_8:  return "8%"
    if has_0:  return "0%"
    return "Exento"


def _hoja_cfdis(wb, nombre: str, cfdis: list[CFDI],
                usados: set[str], color: str,
                es_emitido: bool = False,
                lookup: Optional[dict] = None,
                val_map: Optional[dict] = None,
                complementos_map: Optional[dict] = None) -> None:
    """
    lookup:          dict uuid_lower → {"folio": str, "fecha_banco": str, "fecha_pago_nom": str}
    val_map:         dict uuid_lower → CFDIValidationResult  (de validar_lote)
    complementos_map:dict uuid_lower → [CFDI tipo P que pagan esta factura]
    """
    ws = wb.create_sheet(nombre[:31])

    ws.merge_cells(f"A1:{get_column_letter(len(HDRS_RECIBIDOS))}1")
    c = ws["A1"]
    c.value     = nombre
    c.font      = Font(bold=True, size=13, color="FFFFFF")
    c.fill      = PatternFill("solid", fgColor=color)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    hdrs = HDRS_EMITIDOS if es_emitido else HDRS_RECIBIDOS
    _header_row(ws, 2, hdrs, color,
                WIDTHS_EMITIDOS if es_emitido else WIDTHS_RECIBIDOS)

    lk       = lookup or {}
    comp_map = complementos_map or {}

    for idx, cfdi in enumerate(cfdis):
        row  = idx + 3
        alt  = C_ALT if idx % 2 else C_WHITE
        conc = cfdi.uuid.lower() in usados
        # Para facturas PPD: si tiene complemento de pago registrado, mostrar como conciliada
        via_comp = False
        if not conc and comp_map.get(cfdi.uuid.lower()):
            conc     = True
            via_comp = True
        bg   = C_CONC if conc else alt

        rfc_cont = cfdi.rfc_receptor if es_emitido else cfdi.rfc_emisor
        nom_cont = cfdi.nombre_receptor if es_emitido else cfdi.nombre_emisor

        # Datos bancarios del movimiento conciliado
        info = lk.get(cfdi.uuid.lower(), {})
        # Fallback: si no hay match directo, buscar banco vía el complemento de pago
        if not info.get("banco") and via_comp:
            for _comp in comp_map.get(cfdi.uuid.lower(), []):
                _fb = lk.get(_comp.uuid.lower(), {})
                if _fb.get("banco"):
                    info = _fb
                    break
        banco_nom    = info.get("banco", "")
        folio_banco  = info.get("folio", "")
        fecha_banco  = info.get("fecha_banco", "")
        fecha_nom    = info.get("fecha_pago_nom", "")   # solo nómina tipo N
        es_suma      = info.get("es_suma", False)

        # Detectar nota de crédito absorbida: tipo "E" no conciliada directamente
        # pero cuya factura relacionada sí fue conciliada.
        es_nc_aplicada = False
        uuid_factura_rel = ""
        if not conc and cfdi.tipo == "E" and cfdi.cfdi_relacionados:
            for uuid_rel in cfdi.cfdi_relacionados:
                if uuid_rel.lower() in usados:
                    es_nc_aplicada = True
                    uuid_factura_rel = uuid_rel.upper()
                    break

        # Color: suma=ambar, nota crédito=lila, conciliado=verde, sin banco=alternado
        if es_nc_aplicada:
            bg = "E1BEE7"          # lila claro
        elif conc and es_suma:
            bg = C_REVISAR
        elif not conc:
            bg = alt

        if es_nc_aplicada:
            est_txt = "N.Credito Aplicada"
            est_fg  = "4A148C"   # morado oscuro
            banco_nom   = "Aplicada a factura"
            folio_banco = uuid_factura_rel
        elif conc and es_suma:
            est_txt = "Conciliado SUMA"
            est_fg  = "7B3F00"
        elif conc and via_comp:
            est_txt = "Conciliado (Comp.)"
            est_fg  = "1B5E20"
        elif conc:
            est_txt = "Conciliado"
            est_fg  = "1B5E20"
        else:
            est_txt = "Sin banco"
            est_fg  = "B71C1C"

        _cell(ws, row,  1, cfdi.uuid,              bg=bg)
        _cell(ws, row,  2, cfdi.serie,             bg=bg, align="center")
        _cell(ws, row,  3, cfdi.folio,             bg=bg, align="center")
        _cell(ws, row,  4, cfdi.fecha_emision[:10],bg=bg, align="center")
        _cell(ws, row,  5, rfc_cont,               bg=bg, align="center")
        _cell(ws, row,  6, nom_cont,               bg=bg)
        # Egresos: signo negativo + fuente roja para distinguirlos visualmente
        es_egreso = cfdi.tipo == "E"
        _s_e = -1 if es_egreso else 1
        _fg_e = "B71C1C" if es_egreso else C_INK
        _cell(ws, row,  7, (_s_e * cfdi.subtotal)    or None, fmt=MONEY_FMT, bg=bg, align="right", fg=_fg_e)
        _cell(ws, row,  8, (_s_e * cfdi.descuento)   or None, fmt=MONEY_FMT, bg=bg, align="right", fg=_fg_e)
        _cell(ws, row,  9, (_s_e * cfdi.iva16)       or None, fmt=MONEY_FMT, bg=bg, align="right", fg=_fg_e)
        _cell(ws, row, 10, (_s_e * cfdi.iva_retenido) or None,fmt=MONEY_FMT, bg=bg, align="right", fg=_fg_e)
        _cell(ws, row, 11, (_s_e * cfdi.isr_retenido) or None,fmt=MONEY_FMT, bg=bg, align="right", fg=_fg_e)
        _cell(ws, row, 12, (_s_e * cfdi.total)       or None, fmt=MONEY_FMT, bg=bg, align="right", fg=_fg_e)
        _cell(ws, row, 13, cfdi.mes,               bg=bg, align="center")
        _cell(ws, row, 14, banco_nom  or None,     bg=bg, align="center")
        _cell(ws, row, 15, folio_banco or None,    bg=bg, align="center")
        _cell(ws, row, 16, fecha_banco or None,    bg=bg, align="center")
        _cell(ws, row, 17, fecha_nom   or None,    bg=bg, align="center")
        _cell(ws, row, 18, est_txt,
              bg=bg, align="center", bold=(conc or es_nc_aplicada), fg=est_fg)

        # ── Columnas 19-20: validación fiscal ─────────────────────────────────
        vm  = val_map or {}
        val = vm.get(cfdi.uuid.lower()) or validar_cfdi(cfdi)
        if val.estatus == "VERDE":      val_bg, val_fg = C_VAL_VERDE,    "1B5E20"
        elif val.estatus == "AMARILLO": val_bg, val_fg = C_VAL_AMARILLO, "827717"
        elif val.estatus == "ROJO":     val_bg, val_fg = C_VAL_ROJO,     "B71C1C"
        else:                           val_bg, val_fg = C_VAL_NEGRO,    "FFFFFF"
        obs_txt = "; ".join(val.errores + val.advertencias) or None
        _cell(ws, row, 19, val.estatus, bg=val_bg, align="center", bold=True, fg=val_fg)
        _cell(ws, row, 20, obs_txt, bg=val_bg, wrap=True)

        # ── Col 21: Tipo de CFDI — TipoDeComprobante del comprobante ─────────
        _cell(ws, row, 21, _tipo_cfdi_txt(cfdi.tipo) or None,
              bg=bg, align="center")

        # ── Col 22: Forma de Pago — Art.27 LIVA (efectivo > $2,000) ─────────
        fp_codigo = (cfdi.forma_pago or "").strip()
        fp_txt    = _forma_pago_txt(fp_codigo)
        art27     = fp_codigo == "01" and (cfdi.total or 0) > 2_000.0
        if art27:
            fp_bg  = C_ART27
            fp_fg  = "FFFFFF"
            fp_txt = fp_txt + " ⚠ Art.27 LIVA"
        else:
            fp_bg  = bg
            fp_fg  = "212121"
        _cell(ws, row, 22, fp_txt or None, bg=fp_bg, align="center", bold=art27, fg=fp_fg)

        # ── Columnas 23-26: Método de pago y complemento de pago tipo P ───────
        metodo = (cfdi.metodo_pago or "").strip()
        if cfdi.tipo == "P":
            met_txt  = "—"
            comp_txt = "Complemento de Pago"
            comp_fg  = "FFFFFF"
            comp_bg  = C_HDR_P
            ref_txt  = cfdi.uuid
            fecha_cp = cfdi.fecha_pago or None
        elif metodo == "PPD":
            comps    = comp_map.get(cfdi.uuid.lower(), [])
            met_txt  = "PPD"
            if comps:
                comp_txt = "✔ Liquidada"
                comp_fg  = "1B5E20"
                comp_bg  = C_VAL_VERDE
                ref_txt  = comps[0].uuid
                fecha_cp = comps[0].fecha_pago or None
            else:
                comp_txt = "✗ Pendiente"
                comp_fg  = "B71C1C"
                comp_bg  = C_VAL_ROJO
                ref_txt  = ""
                fecha_cp = None
        elif metodo == "PUE":
            met_txt  = "PUE"
            comp_txt = "—"
            comp_fg  = "9E9E9E"
            comp_bg  = alt
            ref_txt  = ""
            fecha_cp = None
        else:
            met_txt  = metodo or "—"
            comp_txt = "—"
            comp_fg  = "9E9E9E"
            comp_bg  = alt
            ref_txt  = ""
            fecha_cp = None
        _cell(ws, row, 23, met_txt,  bg=alt,     align="center")
        _cell(ws, row, 24, comp_txt, bg=comp_bg, align="center",
              bold=(cfdi.tipo == "P" or metodo == "PPD"), fg=comp_fg)
        _cell(ws, row, 25, ref_txt,  bg=comp_bg, wrap=True)
        _cell(ws, row, 26, fecha_cp, bg=comp_bg, align="center")

        # ── Columna 27: Monto Pagado en Complemento de Pago ───────────────────
        if cfdi.tipo == "P":
            monto_cp    = cfdi.total or None
            monto_cp_bg = C_HDR_P
            monto_cp_fg = "FFFFFF"
        elif metodo == "PPD":
            comps_v     = comp_map.get(cfdi.uuid.lower(), [])
            total_cp    = sum(c.total for c in comps_v)
            monto_cp    = total_cp if total_cp else None
            monto_cp_bg = C_VAL_VERDE if comps_v else C_VAL_ROJO
            monto_cp_fg = "212121"
        else:
            monto_cp    = None
            monto_cp_bg = alt
            monto_cp_fg = "212121"
        _cell(ws, row, 27, monto_cp, fmt=MONEY_FMT, bg=monto_cp_bg,
              align="right", fg=monto_cp_fg)

        # ── Col 28: Tipo de gasto (UsoCFDI G01/G03…) ──────────────────────────
        _cell(ws, row, 28, _uso_cfdi_txt(cfdi.uso_cfdi) or None, bg=bg, wrap=True)

        # ── Cols 29-30: CFDIs relacionados y leyenda fiscal del egreso ────────
        rels = list(getattr(cfdi, "cfdi_rel_todos", ()) or ())
        _cell(ws, row, 29, _rel_txt(rels) or None, bg=bg, wrap=True)
        ley = _leyenda_relacion(cfdi.tipo, rels)
        # El egreso sin relacionar es una observación, no un dato: se resalta.
        ley_falta = ley.startswith("Egreso SIN")
        _cell(ws, row, 30, ley or None,
              bg=(C_VAL_AMARILLO if ley_falta else bg),
              fg=("827717" if ley_falta else "212121"),
              bold=ley_falta, wrap=True)

        # ── Cols 31-32: conceptos del CFDI, todos separados por " | " ──────────
        _cell(ws, row, 31,
              " | ".join(getattr(cfdi, "conceptos_clave", ()) or ()) or None,
              bg=bg, wrap=True)
        _cell(ws, row, 32,
              " | ".join(getattr(cfdi, "conceptos_desc", ()) or ()) or None,
              bg=bg, wrap=True)

        # ── Cols 33-37: desglose IVA por tasa (ocultas, para VBA CalcularIVA) ──
        _cell(ws, row, 33, (_s_e * getattr(cfdi, "iva_base_16", 0)) or None, fmt=MONEY_FMT, bg=bg, align="right")
        _cell(ws, row, 34, (_s_e * getattr(cfdi, "iva_base_8", 0)) or None, fmt=MONEY_FMT, bg=bg, align="right")
        _cell(ws, row, 35, (_s_e * getattr(cfdi, "iva_importe_8", 0)) or None, fmt=MONEY_FMT, bg=bg, align="right")
        _cell(ws, row, 36, (_s_e * getattr(cfdi, "iva_base_0", 0)) or None, fmt=MONEY_FMT, bg=bg, align="right")
        _cell(ws, row, 37, (_s_e * getattr(cfdi, "iva_exento_base", 0)) or None, fmt=MONEY_FMT, bg=bg, align="right")
        _cell(ws, row, 38, _tipo_tasa(cfdi), bg=bg, align="center")

    # Ocultar columnas auxiliares de desglose IVA
    for col_letter in ("AG", "AH", "AI", "AJ", "AK", "AL"):
        ws.column_dimensions[col_letter].hidden = True

    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:{get_column_letter(len(HDRS_RECIBIDOS))}{len(cfdis) + 2}"


# ── Hoja: Nómina (tipo N) ─────────────────────────────────────────────────────

HDRS_NOMINA = [
    "UUID", "FOLIO CFDI", "FECHA EMISION", "RFC EMPLEADO", "NOMBRE EMPLEADO",
    "PERCEPCIONES", "DEDUCCIONES", "OTROS PAGOS", "NETO PAGO",
    "FECHA PAGO NOMINA", "BANCO", "FOLIO / REF BANCARIA", "FECHA PAGO BANCO", "ESTATUS",
    "VALIDACION FISCAL", "OBSERVACIONES FISCALES",
]
WIDTHS_NOMINA = [38, 10, 14, 18, 42, 15, 15, 13, 15, 16, 16, 20, 16, 14, 16, 48]


def _hoja_nomina(wb, cfdis_nom: list[CFDI],
                 usados: set[str],
                 lookup: Optional[dict] = None,
                 val_map: Optional[dict] = None) -> None:
    if not cfdis_nom:
        return
    ws  = wb.create_sheet("Nomina")
    lk  = lookup or {}
    vm  = val_map or {}
    N   = len(HDRS_NOMINA)

    ws.merge_cells(f"A1:{get_column_letter(N)}1")
    c = ws["A1"]
    c.value     = "Recibos de Nómina (CFDI tipo N)"
    c.font      = Font(bold=True, size=13, color="FFFFFF")
    c.fill      = PatternFill("solid", fgColor=C_HDR_N)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    _header_row(ws, 2, HDRS_NOMINA, C_HDR_N, WIDTHS_NOMINA)

    for idx, cfdi in enumerate(cfdis_nom):
        row  = idx + 3
        alt  = C_ALT if idx % 2 else C_WHITE
        conc = cfdi.uuid.lower() in usados
        bg   = C_CONC if conc else alt

        info        = lk.get(cfdi.uuid.lower(), {})
        banco_nom   = info.get("banco", "")
        folio_banco = info.get("folio", "")
        fecha_banco = info.get("fecha_banco", "")
        fecha_nom   = info.get("fecha_pago_nom", "") or cfdi.fecha_pago

        _cell(ws, row,  1, cfdi.uuid,              bg=bg)
        _cell(ws, row,  2, cfdi.folio,             bg=bg, align="center")
        _cell(ws, row,  3, cfdi.fecha_emision[:10],bg=bg, align="center")
        _cell(ws, row,  4, cfdi.rfc_receptor,      bg=bg, align="center")
        _cell(ws, row,  5, cfdi.nombre_receptor,   bg=bg)
        _cell(ws, row,  6, cfdi.nom_percepciones or None, fmt=MONEY_FMT, bg=bg, align="right")
        _cell(ws, row,  7, cfdi.nom_deducciones  or None, fmt=MONEY_FMT, bg=bg, align="right")
        _cell(ws, row,  8, cfdi.nom_otros_pagos  or None, fmt=MONEY_FMT, bg=bg, align="right")
        _cell(ws, row,  9, cfdi.total            or None, fmt=MONEY_FMT, bg=bg, align="right")
        _cell(ws, row, 10, fecha_nom   or None,    bg=bg, align="center")
        _cell(ws, row, 11, banco_nom   or None,    bg=bg, align="center")
        _cell(ws, row, 12, folio_banco or None,    bg=bg, align="center")
        _cell(ws, row, 13, fecha_banco or None,    bg=bg, align="center")
        _cell(ws, row, 14, "Conciliado" if conc else "Sin banco",
              bg=bg, align="center",
              bold=conc, fg="1B5E20" if conc else "B71C1C")

        # ── Columnas 15-16: validación fiscal (R006 régimen emisor) ──────────
        val = vm.get(cfdi.uuid.lower()) or validar_cfdi(cfdi)
        if val.estatus == "VERDE":      val_bg, val_fg = C_VAL_VERDE,    "1B5E20"
        elif val.estatus == "AMARILLO": val_bg, val_fg = C_VAL_AMARILLO, "827717"
        elif val.estatus == "ROJO":     val_bg, val_fg = C_VAL_ROJO,     "B71C1C"
        else:                           val_bg, val_fg = C_VAL_NEGRO,    "FFFFFF"
        obs_txt = "; ".join(val.errores + val.advertencias) or None
        _cell(ws, row, 15, val.estatus, bg=val_bg, align="center", bold=True, fg=val_fg)
        _cell(ws, row, 16, obs_txt, bg=val_bg, wrap=True)

    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:{get_column_letter(N)}{len(cfdis_nom) + 2}"


# ── Helpers de complementos de pago ──────────────────────────────────────────

def _build_complementos_map(todos_recibidos: list[CFDI],
                             todos_emitidos:  list[CFDI]) -> dict[str, list[CFDI]]:
    """
    Retorna dict {uuid_factura_lower → [CFDI tipo P que la pagan]}.
    Permite saber, para cualquier factura PPD, si ya tiene complemento de pago.
    """
    m: dict[str, list[CFDI]] = {}
    for cfdi in todos_recibidos + todos_emitidos:
        if cfdi.tipo == "P":
            for uid in cfdi.cfdi_facturas_pagadas:
                m.setdefault(uid.lower(), []).append(cfdi)
    return m


def _hoja_complementos_pago(
    wb,
    tipo_p_cfdis:   list[CFDI],
    todos_por_uuid: dict[str, CFDI],
    usados_r:       set[str],
    usados_e:       set[str],
    lookup_banco:   dict[str, dict],
    val_map:        Optional[dict] = None,
    titulo:         str = "Complementos de Pago",
    color_hdr:      str = C_HDR_P,
) -> None:
    """
    Hoja dedicada a los CFDI tipo P (Complementos de Pago).
    Muestra cada complemento con las facturas PPD que liquida, estado en cartera
    y validación fiscal R006 (régimen emisor/receptor vs CSF de la empresa).
    """
    if not tipo_p_cfdis:
        return
    ws  = wb.create_sheet(titulo)
    N   = len(HDRS_COMP_PAGO)
    vm  = val_map or {}

    ws.merge_cells(f"A1:{get_column_letter(N)}1")
    c = ws["A1"]
    c.value     = f"{titulo} — CFDI tipo P"
    c.font      = Font(bold=True, size=13, color="FFFFFF")
    c.fill      = PatternFill("solid", fgColor=color_hdr)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    _header_row(ws, 2, HDRS_COMP_PAGO, color_hdr, WIDTHS_COMP_PAGO)

    usados_all = usados_r | usados_e

    for idx, cfdi in enumerate(tipo_p_cfdis):
        row = idx + 3
        alt = C_ALT if idx % 2 else C_WHITE

        # El complemento tipo P liquida facturas tipo I — esas son las que
        # aparecen en lookup_banco, no el UUID del propio complemento.
        # Buscamos banco info en las facturas referenciadas primero.
        info: dict = {}
        for uid_ref in cfdi.cfdi_facturas_pagadas:
            info = lookup_banco.get(uid_ref.lower(), {})
            if info:
                break
        if not info:
            info = lookup_banco.get(cfdi.uuid.lower(), {})

        # Conciliado si alguna de las facturas que paga está en banco
        conc = (any(uid.lower() in usados_all for uid in cfdi.cfdi_facturas_pagadas)
                or cfdi.uuid.lower() in usados_all)
        bg   = C_CONC if conc else alt

        banco_nom = info.get("banco", "")

        # Facturas pagadas: cuántas están en la cartera cargada
        facturas_en_cartera = [uid for uid in cfdi.cfdi_facturas_pagadas
                               if uid.lower() in todos_por_uuid]
        n_total   = len(cfdi.cfdi_facturas_pagadas)
        n_cartera = len(facturas_en_cartera)

        if n_total == 0:
            cartera_txt = "Sin facturas referenciadas"
            cartera_fg  = "B71C1C"
        elif n_cartera == n_total:
            cartera_txt = f"Todas ({n_cartera}/{n_total})"
            cartera_fg  = "1B5E20"
        elif n_cartera > 0:
            cartera_txt = f"Parcial ({n_cartera}/{n_total})"
            cartera_fg  = "E65100"
        else:
            cartera_txt = f"Ninguna (0/{n_total})"
            cartera_fg  = "B71C1C"

        uuids_txt = "\n".join(cfdi.cfdi_facturas_pagadas)
        if n_total > 1:
            ws.row_dimensions[row].height = max(16, n_total * 13)

        monto_xml = cfdi.total or 0.0
        monto_bco = info.get("monto_banco", 0.0) or 0.0

        _cell(ws, row,  1, cfdi.uuid,                                   bg=bg)
        _cell(ws, row,  2, cfdi.fecha_pago or cfdi.fecha_emision[:10],  bg=bg, align="center")
        _cell(ws, row,  3, monto_xml or None,          fmt=MONEY_FMT,   bg=bg, align="right")

        # Col 4: Monto Banco / Col 5: Diferencia XML vs Banco
        _cell(ws, row,  4, monto_bco or None,          fmt=MONEY_FMT,   bg=bg, align="right")
        if monto_bco:
            diferencia = round(monto_xml - monto_bco, 2)
            dif_bg = C_VAL_VERDE if diferencia == 0.0 else C_VAL_ROJO
            dif_fg = "1B5E20"    if diferencia == 0.0 else "B71C1C"
            _cell(ws, row,  5, diferencia, fmt=MONEY_FMT, bg=dif_bg,
                  align="right", bold=(diferencia != 0.0), fg=dif_fg)
        else:
            _cell(ws, row,  5, None, bg=alt, align="center")

        _cell(ws, row,  6, cfdi.rfc_emisor,                             bg=bg, align="center")
        _cell(ws, row,  7, cfdi.nombre_emisor,                          bg=bg)
        _cell(ws, row,  8, cfdi.rfc_receptor,                           bg=bg, align="center")
        _cell(ws, row,  9, cfdi.nombre_receptor,                        bg=bg)
        _cell(ws, row, 10, n_total,                                     bg=bg, align="center")
        _cell(ws, row, 11, uuids_txt,                                   bg=bg, wrap=True)
        _cell(ws, row, 12, cartera_txt,  bold=True, fg=cartera_fg,      bg=bg, align="center")
        _cell(ws, row, 13, banco_nom or None,                           bg=bg, align="center")
        _cell(ws, row, 14, "Conciliado" if conc else "Sin banco",
              bg=bg, align="center", bold=conc,
              fg="1B5E20" if conc else "B71C1C")

        # ── Columnas 15-16: validación fiscal R006 (régimen vs CSF empresa) ──
        val = vm.get(cfdi.uuid.lower()) or validar_cfdi(cfdi)
        if val.estatus == "VERDE":      val_bg, val_fg = C_VAL_VERDE,    "1B5E20"
        elif val.estatus == "AMARILLO": val_bg, val_fg = C_VAL_AMARILLO, "827717"
        elif val.estatus == "ROJO":     val_bg, val_fg = C_VAL_ROJO,     "B71C1C"
        else:                           val_bg, val_fg = C_VAL_NEGRO,    "FFFFFF"
        obs_txt = "; ".join(val.errores + val.advertencias) or None
        _cell(ws, row, 15, val.estatus, bg=val_bg, align="center", bold=True, fg=val_fg)
        _cell(ws, row, 16, obs_txt, bg=val_bg, wrap=True)

    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:{get_column_letter(N)}{len(tipo_p_cfdis) + 2}"


# ── Hoja: Pólizas COI ────────────────────────────────────────────────────────

# Columna de encabezado de referencia (filas 1-2 informativas, no se importan)
_HDRS_COI_REF = [
    "Tipo Pól", "Núm. Póliza", "Concepto póliza / No. Cuenta / Marcador",
    "Día / Concepto mov.", "T.C. / Serie", "Debe / Folio",
    "Haber / RFC Emisor", "C.Costo / RFC Receptor", "Proyecto / Monto",
    "Tipo CFDI",
]
_WIDTHS_COI = [8, 16, 55, 30, 10, 16, 16, 20, 40, 8]

# Colores de la hoja pólizas
_C_COI_HDR   = "37474F"   # gris oscuro  — encabezado de póliza
_C_COI_PAR   = "E8F5E9"   # verde muy claro — partida normal
_C_COI_PAR2  = "F1F8E9"   # verde pálido alternado
_C_COI_CFDI  = "E3F2FD"   # azul claro  — fila de datos CFDI
_C_COI_MRK   = "FFF8E1"   # amarillo pálido — marcadores INICIO/FIN
_C_COI_ERR   = "FFCDD2"   # rojo claro  — póliza descuadrada


def _hoja_polizas_coi(wb, polizas: list) -> None:
    """
    Genera la hoja 'Pólizas COI' en el formato de importación de Aspel COI.
    Las filas 1-2 son encabezados de referencia; los datos empiezan en la fila 3,
    igual que el formato estándar del archivo de importación COI.
    """
    if not polizas:
        return

    from polizas_coi import polizas_a_filas

    ws = wb.create_sheet("Pólizas COI")
    N  = len(_HDRS_COI_REF)

    # ── Fila 1: título ────────────────────────────────────────────────────────
    ws.merge_cells(f"A1:{get_column_letter(N)}1")
    c1 = ws["A1"]
    c1.value     = "Pólizas Generadas — Formato Aspel COI (importar desde fila 3)"
    c1.font      = Font(bold=True, size=12, color="FFFFFF")
    c1.fill      = PatternFill("solid", fgColor=_C_COI_HDR)
    c1.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 22

    # ── Fila 2: encabezados de columnas (referencia) ──────────────────────────
    for col, lbl in enumerate(_HDRS_COI_REF, 1):
        c = ws.cell(row=2, column=col, value=lbl)
        c.font      = Font(bold=True, size=8, color="FFFFFF")
        c.fill      = PatternFill("solid", fgColor="546E7A")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border    = THIN
    ws.row_dimensions[2].height = 16
    for col, w in enumerate(_WIDTHS_COI, 1):
        ws.column_dimensions[get_column_letter(col)].width = w

    # ── Filas de datos a partir de la 3 ──────────────────────────────────────
    filas = polizas_a_filas(polizas)
    row = 3
    pol_idx = 0

    # Para colorear filas alternas de partidas por póliza
    _par_alts = [_C_COI_PAR, _C_COI_PAR2]

    for fila_data in filas:
        a_val = fila_data[0]   # col A
        b_val = fila_data[1]   # col B
        c_val = fila_data[2]   # col C

        # Determinar tipo de fila y color
        if a_val in ("Ig", "Eg", "Dr"):                  # encabezado de póliza
            # Buscar si la póliza cuadra para colorear en rojo si no
            pol = polizas[pol_idx] if pol_idx < len(polizas) else None
            bg  = _C_COI_ERR if (pol and not pol.cuadrada) else _C_COI_HDR
            fg  = "FFFFFF"
            bold = True
            pol_idx += 1
            par_alt_idx = 0
        elif b_val == "FIN_PARTIDAS":                     # cierre de póliza
            bg, fg, bold = "CFD8DC", "37474F", True
        elif c_val in ("INICIO_CFDI", "FIN_CFDI"):        # marcadores CFDI
            bg, fg, bold = _C_COI_MRK, "5D4037", False
        elif c_val is not None and b_val is None and a_val is None:
            # Fila de datos CFDI (entre marcadores)
            bg, fg, bold = _C_COI_CFDI, "0D47A1", False
        else:                                              # partida normal
            bg   = _par_alts[par_alt_idx % 2] if 'par_alt_idx' in dir() else _C_COI_PAR
            par_alt_idx = (par_alt_idx + 1) if 'par_alt_idx' in dir() else 1
            fg, bold = "212121", False

        ws.row_dimensions[row].height = 14
        for col, val in enumerate(fila_data, 1):
            c = ws.cell(row=row, column=col, value=val)
            c.font      = Font(bold=bold, size=9, color=fg)
            c.fill      = PatternFill("solid", fgColor=bg)
            c.border    = THIN
            c.alignment = Alignment(horizontal="center" if col in (1, 2, 5, 6, 7, 10)
                                    else "left", vertical="center")
            # Formato numérico para Debe/Haber/Monto
            if col in (6, 7, 8) and isinstance(val, (int, float)):
                c.number_format = MONEY_FMT
            # Fecha CFDI como fecha Excel
            if col == 3 and hasattr(val, "strftime"):
                c.number_format = DATE_FMT

        row += 1

    ws.freeze_panes = "A3"

    # Leyenda al pie
    pie = row + 1
    ws.merge_cells(f"A{pie}:{get_column_letter(N)}{pie}")
    lp = ws.cell(row=pie, column=1,
                 value="LEYENDA: Encabezado póliza (oscuro) | Partida (verde) | "
                        "Datos CFDI (azul) | Marcadores (amarillo) | "
                        "Póliza descuadrada → encabezado rojo")
    lp.font      = Font(italic=True, size=8, color="546E7A")
    lp.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[pie].height = 14


# ── Hoja: Cálculo Previo de IVA ──────────────────────────────────────────────

_C_IVA_HDR    = "1B5FAF"   # encabezado azul oscuro
_C_IVA_SEC    = "E8EAF6"   # sección — lila muy claro
_C_IVA_TOTAL  = "E3F2FD"   # total parcial — azul claro
_C_IVA_RESULT = "FFF3E0"   # resultado final — ámbar claro


_MESES_IVA = {1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
              5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
              9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"}


def _hojas_iva_previo(wb, todos_recibidos: list[CFDI],
                      todos_emitidos: list[CFDI],
                      meses_iva: list[int] | None = None) -> None:
    """Genera una hoja 'Calculo Previo IVA' por cada mes en *meses_iva*.
    Si la lista está vacía o es None, genera una sola hoja con todos los meses."""
    if not meses_iva:
        _hoja_iva_previo(wb, todos_recibidos, todos_emitidos, mes_iva=None)
    else:
        for m in sorted(meses_iva):
            _hoja_iva_previo(wb, todos_recibidos, todos_emitidos, mes_iva=m)


def _hoja_iva_previo(wb, todos_recibidos: list[CFDI],
                     todos_emitidos: list[CFDI],
                     mes_iva: int | None = None) -> None:
    """Genera UNA hoja de cálculo previo IVA.
    Si *mes_iva* es 1–12 filtra solo CFDIs de ese mes; None = todos."""

    if mes_iva:
        todos_recibidos = [c for c in todos_recibidos if c.mes == mes_iva]
        todos_emitidos  = [c for c in todos_emitidos  if c.mes == mes_iva]

    titulo_mes = _MESES_IVA.get(mes_iva, "") if mes_iva else "Todos los meses"
    nombre_hoja = f"IVA {_MESES_IVA[mes_iva]}" if mes_iva else "Calculo Previo IVA"

    ws = wb.create_sheet(nombre_hoja)
    ws.column_dimensions["A"].width = 58
    ws.column_dimensions["B"].width = 20

    # ── Acumular por tipo y tasa ─────────────────────────────────────────────
    # Recibidos tipo I = Ingresos (IVA Causado)
    # Recibidos tipo E que nos emiten a nosotros = Devoluciones (reducen ingreso)
    # Emitidos tipo I/E = Gastos/NC (IVA Acreditable)
    ing_16 = ing_8 = ing_0 = ing_ex = 0.0
    iva_trasl_16 = iva_trasl_8 = 0.0
    iva_ret_pm = 0.0   # IVA retenido por persona moral (nos retienen)

    gas_16 = gas_8 = gas_0 = gas_ex = 0.0
    iva_acred_16 = iva_acred_8 = 0.0
    iva_ret_ent = 0.0  # IVA retenido enterado (nosotros retenemos)

    for c in todos_recibidos:
        if c.tipo in ("N", "P", "T"):
            continue
        if c.tipo == "E":
            # Egreso recibido = nos devuelven dinero → reduce IVA causado
            ing_16 -= getattr(c, "iva_base_16", 0.0) or 0.0
            ing_8  -= getattr(c, "iva_base_8",  0.0) or 0.0
            ing_0  -= getattr(c, "iva_base_0",  0.0) or 0.0
            ing_ex -= getattr(c, "iva_exento_base", 0.0) or 0.0
            iva_trasl_16 -= c.iva16 - (getattr(c, "iva_importe_8", 0.0) or 0.0)
            iva_trasl_8  -= getattr(c, "iva_importe_8", 0.0) or 0.0
        else:
            # Ingreso recibido = nos pagan → IVA causado
            ing_16 += getattr(c, "iva_base_16", 0.0) or 0.0
            ing_8  += getattr(c, "iva_base_8",  0.0) or 0.0
            ing_0  += getattr(c, "iva_base_0",  0.0) or 0.0
            ing_ex += getattr(c, "iva_exento_base", 0.0) or 0.0
            iva_trasl_16 += c.iva16 - (getattr(c, "iva_importe_8", 0.0) or 0.0)
            iva_trasl_8  += getattr(c, "iva_importe_8", 0.0) or 0.0
            iva_ret_pm   += c.iva_retenido

    for c in todos_emitidos:
        if c.tipo in ("N", "P", "T"):
            continue
        if c.tipo == "E":
            # Egreso que emitimos = nota de crédito a proveedor → reduce acreditable
            gas_16 -= getattr(c, "iva_base_16", 0.0) or 0.0
            gas_8  -= getattr(c, "iva_base_8",  0.0) or 0.0
            gas_0  -= getattr(c, "iva_base_0",  0.0) or 0.0
            gas_ex -= getattr(c, "iva_exento_base", 0.0) or 0.0
            iva_acred_16 -= c.iva16 - (getattr(c, "iva_importe_8", 0.0) or 0.0)
            iva_acred_8  -= getattr(c, "iva_importe_8", 0.0) or 0.0
        else:
            # Ingreso que emitimos = gasto nuestro → IVA acreditable
            gas_16 += getattr(c, "iva_base_16", 0.0) or 0.0
            gas_8  += getattr(c, "iva_base_8",  0.0) or 0.0
            gas_0  += getattr(c, "iva_base_0",  0.0) or 0.0
            gas_ex += getattr(c, "iva_exento_base", 0.0) or 0.0
            iva_acred_16 += c.iva16 - (getattr(c, "iva_importe_8", 0.0) or 0.0)
            iva_acred_8  += getattr(c, "iva_importe_8", 0.0) or 0.0
            iva_ret_ent  += c.iva_retenido

    total_actos_cobrados = ing_16 + ing_8 + ing_0 + ing_ex
    total_iva_trasl = iva_trasl_16 + iva_trasl_8
    total_actos_pagados = gas_16 + gas_8 + gas_0 + gas_ex
    iva_acred_total = iva_acred_16 + iva_acred_8
    saldo = total_iva_trasl - iva_ret_pm - iva_acred_total - iva_ret_ent

    # ── Helpers de escritura ─────────────────────────────────────────────────
    def fila(row, label, val=None, *, bold=False, bg=C_WHITE, fg=C_INK,
             italic=False, indent=False):
        c_a = ws.cell(row=row, column=1, value=label)
        c_a.font = Font(name=FUENTE, bold=bold, italic=italic, size=9, color=fg)
        c_a.fill = PatternFill("solid", fgColor=bg)
        c_a.alignment = Alignment(horizontal="left", vertical="center",
                                  indent=2 if indent else 0)
        c_a.border = THIN
        c_b = ws.cell(row=row, column=2, value=val)
        c_b.font = Font(name=FUENTE, bold=bold, size=9, color=fg)
        c_b.fill = PatternFill("solid", fgColor=bg)
        c_b.alignment = Alignment(horizontal="right", vertical="center")
        c_b.number_format = MONEY_FMT
        c_b.border = THIN
        ws.row_dimensions[row].height = 18

    # ── Encabezado ───────────────────────────────────────────────────────────
    ws.merge_cells("A1:B1")
    c_h = ws["A1"]
    c_h.value = f"Cálculo Previo de IVA — {titulo_mes}"
    c_h.font = Font(bold=True, size=14, color="FFFFFF")
    c_h.fill = PatternFill("solid", fgColor=_C_IVA_HDR)
    c_h.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    r = 3
    # ── IVA CAUSADO ──────────────────────────────────────────────────────────
    fila(r, "IVA CAUSADO", bold=True, bg=_C_IVA_SEC); r += 1
    fila(r, "Actos a tasa general", ing_16 or None, italic=True, indent=True); r += 1
    fila(r, "Actos a tasa fronteriza", ing_8 or None, italic=True, indent=True); r += 1
    fila(r, "Actos a tasa 0%", ing_0 or None, italic=True, indent=True); r += 1
    fila(r, "Actos exentos", ing_ex or None, italic=True, indent=True); r += 1
    fila(r, "Total de actos efectivamente cobrados al periodo",
         total_actos_cobrados, bold=True, bg=_C_IVA_TOTAL); r += 2
    fila(r, "IVA trasladado a tasa general", iva_trasl_16 or None,
         italic=True, indent=True); r += 1
    fila(r, "IVA trasladado a tasa fronteriza", iva_trasl_8 or None,
         italic=True, indent=True); r += 1
    fila(r, "Total de IVA trasladado del periodo al que corresponde el pago",
         total_iva_trasl, bold=True, bg=_C_IVA_TOTAL); r += 2
    fila(r, "IVA retenido por persona moral", iva_ret_pm or None,
         italic=True, indent=True); r += 2

    # ── IVA ACREDITABLE ──────────────────────────────────────────────────────
    fila(r, "IVA ACREDITABLE", bold=True, bg=_C_IVA_SEC); r += 1
    fila(r, "Actos a tasa general en compras nacionales", gas_16 or None,
         italic=True, indent=True); r += 1
    fila(r, "Actos a tasa general en importaciones", None,
         italic=True, indent=True); r += 1
    fila(r, "Actos a tasa general", gas_16 or None, bold=True, indent=True); r += 1
    fila(r, "Actos a tasa fronteriza en compras nacionales", gas_8 or None,
         italic=True, indent=True); r += 1
    fila(r, "Actos a tasa fronteriza en importaciones", None,
         italic=True, indent=True); r += 1
    fila(r, "Actos a tasa fronteriza", gas_8 or None, bold=True, indent=True); r += 1
    fila(r, "Actos a tasa 0%", gas_0 or None, italic=True, indent=True); r += 1
    fila(r, "Actos exentos en compras nacionales", gas_ex or None,
         italic=True, indent=True); r += 1
    fila(r, "Actos exentos en importaciones", None,
         italic=True, indent=True); r += 1
    fila(r, "Actos exentos", gas_ex or None, bold=True, indent=True); r += 1
    fila(r, "Total de actos efectivamente pagados al periodo",
         total_actos_pagados, bold=True, bg=_C_IVA_TOTAL); r += 2
    fila(r, "IVA acreditable a tasa general en compras nacionales",
         iva_acred_16 or None, italic=True, indent=True); r += 1
    fila(r, "IVA acreditable a tasa general en importaciones", None,
         italic=True, indent=True); r += 1
    fila(r, "IVA acreditable a tasa general", iva_acred_16 or None,
         bold=True, indent=True); r += 1
    fila(r, "IVA acreditable a tasa fronteriza en compras nacionales",
         iva_acred_8 or None, italic=True, indent=True); r += 1
    fila(r, "IVA acreditable a tasa fronteriza en importaciones", None,
         italic=True, indent=True); r += 1
    fila(r, "IVA acreditable a tasa fronteriza", iva_acred_8 or None,
         bold=True, indent=True); r += 1
    fila(r, "Total de IVA acreditable del periodo al que corresponde el pago",
         iva_acred_total, bold=True, bg=_C_IVA_TOTAL); r += 2
    fila(r, "IVA Retenido Enterado", iva_ret_ent or None,
         bold=True, indent=True); r += 2

    # ── Resultado ────────────────────────────────────────────────────────────
    saldo_fg = "B71C1C" if saldo < 0 else "1B5E20"
    fila(r, "Saldo a cargo o a favor del IVA del periodo al que corresponde el pago",
         saldo, bold=True, bg=_C_IVA_RESULT, fg=saldo_fg); r += 2
    fila(r, "Saldo a favor pendiente de acreditar", 0.0, italic=True, indent=True)

    ws.sheet_properties.tabColor = _C_IVA_HDR


# ── Hoja: Resumen ─────────────────────────────────────────────────────────────

def _hoja_resumen(wb,
                  conciliados:    list[MovConciliado],
                  sin_cfdi:       list[MovBanco],
                  cfdi_sin_bco_r: list[CFDI],
                  cfdi_sin_bco_e: list[CFDI],
                  todos_recibidos: Optional[list[CFDI]] = None,
                  todos_emitidos:  Optional[list[CFDI]] = None,
                  val_map:         Optional[dict] = None) -> None:
    ws = wb.create_sheet("Resumen")

    def fila(row, label, val, bg=C_WHITE, fmt=None):
        _cell(ws, row, 1, label, bold=True, bg=bg, size=10)
        _cell(ws, row, 2, val, bg=bg, align="right", fmt=fmt, size=10)
        ws.row_dimensions[row].height = 18

    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 18

    ws.merge_cells("A1:B1")
    c = ws["A1"]
    c.value     = "Resumen de Conciliación"
    c.font      = Font(bold=True, size=14, color="FFFFFF")
    c.fill      = PatternFill("solid", fgColor=C_HDR_S)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    total_conc   = len(conciliados)
    total_revisar= sum(1 for mc in conciliados if mc.revisar)
    total_pend   = len(sin_cfdi)
    total_mov    = total_conc + total_pend
    total_r_sc   = len(cfdi_sin_bco_r)
    total_e_sc   = len(cfdi_sin_bco_e)

    monto_conc    = sum(mc.banco.monto for mc in conciliados)
    monto_revisar = sum(mc.banco.monto for mc in conciliados if mc.revisar)
    monto_pend    = sum(mb.monto       for mb in sin_cfdi)
    total_cfdis_r = sum(c.total for c in cfdi_sin_bco_r) if cfdi_sin_bco_r else 0.0

    fila(2,  "Movimientos bancarios totales",            total_mov)
    fila(3,  "Conciliados con CFDI",                     total_conc,    C_CONC)
    fila(4,  "⚠ Conciliados que requieren revisión",     total_revisar, C_REVISAR)
    fila(5,  "Sin CFDI encontrado",                      total_pend,    C_SIN)
    fila(6,  "Monto conciliado ($)",                     monto_conc,    C_CONC,    MONEY_FMT)
    fila(7,  "⚠ Monto a revisar ($)",                    monto_revisar, C_REVISAR, MONEY_FMT)
    fila(8,  "Monto pendiente ($)",                      monto_pend,    C_SIN,     MONEY_FMT)
    fila(9,  "CFDIs Recibidos sin banco",                total_r_sc,    C_PEND)
    fila(10, "CFDIs Emitidos sin banco",                 total_e_sc,    C_PEND)
    fila(11, "Monto CFDIs Recibidos sin banco ($)",      total_cfdis_r, C_PEND,    MONEY_FMT)
    fila(12, "% Conciliado",
         round(total_conc / total_mov * 100, 1) if total_mov else 0,
         C_CONC, "0.0%")

    # ── Complementos de Pago (PPD / tipo P) ──────────────────────────────────
    todos_i   = (todos_recibidos or []) + (todos_emitidos or [])
    todos_ppd = [c for c in todos_i if c.tipo == "I" and c.metodo_pago == "PPD"]
    todos_cp  = [c for c in todos_i if c.tipo == "P"]
    fp_set    = {uid.lower() for cp in todos_cp for uid in cp.cfdi_facturas_pagadas}
    n_ppd       = len(todos_ppd)
    n_ppd_liq   = sum(1 for c in todos_ppd if c.uuid.lower() in fp_set)
    n_ppd_pend  = n_ppd - n_ppd_liq
    n_comp_pago = len(todos_cp)

    next_row = 14
    if n_ppd or n_comp_pago:
        ws.merge_cells(f"A{next_row}:B{next_row}")
        c_ph = ws.cell(row=next_row, column=1, value="COMPLEMENTOS DE PAGO (PPD / TIPO P)")
        c_ph.font      = Font(bold=True, size=9, color="FFFFFF")
        c_ph.fill      = PatternFill("solid", fgColor=C_HDR_P)
        c_ph.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[next_row].height = 16
        next_row += 1
        fila(next_row, "Facturas PPD (pago diferido)",            n_ppd,       C_PEND);  next_row += 1
        fila(next_row, "  PPD Liquidadas (con complemento P)",    n_ppd_liq,   C_CONC);  next_row += 1
        fila(next_row, "  PPD Pendientes (sin complemento P)",    n_ppd_pend,  C_SIN);   next_row += 1
        fila(next_row, "Complementos de Pago (tipo P)",           n_comp_pago, C_SAL);   next_row += 1
        next_row += 1  # espaciado

    # Nota de advertencia si hay registros a revisar
    if total_revisar:
        nota_row = next_row
        ws.merge_cells(f"A{nota_row}:B{nota_row + 2}")
        c_nota = ws.cell(row=nota_row, column=1,
            value=(
                f"⚠ ATENCIÓN: {total_revisar} movimiento(s) fueron conciliados "
                "automáticamente, pero existen CFDIs con el mismo importe y mes en la "
                "cartera. Estos datos conciliados deben revisarse minuciosamente para "
                "confirmar que el CFDI asignado es el correcto."
            ))
        c_nota.font      = Font(bold=True, size=9, color="7B3F00")
        c_nota.fill      = PatternFill("solid", fgColor=C_REVISAR)
        c_nota.alignment = Alignment(horizontal="left", vertical="top",
                                     wrap_text=True)
        ws.row_dimensions[nota_row].height     = 14
        ws.row_dimensions[nota_row + 1].height = 14
        ws.row_dimensions[nota_row + 2].height = 14
        next_row = nota_row + 4

    # Sección de validación fiscal — incluye todos los tipos (I, E, P, N).
    # Usa val_map ya computado con rfc_empresa para que R006 compare contra
    # el catálogo de la empresa; sin val_map recalcula sin contexto de empresa.
    all_cfdi = [c for c in ((todos_recibidos or []) + (todos_emitidos or [])) if c.uuid]
    if all_cfdi:
        v_lote = val_map if val_map is not None else validar_lote(all_cfdi)
        n_verde    = sum(1 for v in v_lote.values() if v.estatus == "VERDE")
        n_amarillo = sum(1 for v in v_lote.values() if v.estatus == "AMARILLO")
        n_rojo     = sum(1 for v in v_lote.values() if v.estatus == "ROJO")
        n_negro    = sum(1 for v in v_lote.values() if v.estatus == "NEGRO")

        ws.merge_cells(f"A{next_row}:B{next_row}")
        c_vh = ws.cell(row=next_row, column=1, value="VALIDACION FISCAL DE CFDIs")
        c_vh.font      = Font(bold=True, size=9, color="FFFFFF")
        c_vh.fill      = PatternFill("solid", fgColor=C_HDR_S)
        c_vh.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[next_row].height = 16
        next_row += 1

        fila(next_row,     "Deducibles sin errores (VERDE)",      n_verde,    C_VAL_VERDE)
        fila(next_row + 1, "Requieren revision (AMARILLO)",       n_amarillo, C_VAL_AMARILLO)
        fila(next_row + 2, "No deducibles — errores (ROJO)",      n_rojo,     C_VAL_ROJO)
        fila(next_row + 3, "Sin efecto fiscal — critico (NEGRO)", n_negro,    C_VAL_NEGRO)


# ── Función principal ─────────────────────────────────────────────────────────

def exportar_conciliacion(
    conciliados:    list[MovConciliado],
    sin_cfdi:       list[MovBanco],
    cfdi_sin_bco_r: list[CFDI],
    cfdi_sin_bco_e: list[CFDI],
    todos_recibidos: list[CFDI],
    todos_emitidos:  list[CFDI],
    output_path:     str | Path,
    nombre_banco:    str = "Banco",
    color_banco:     str = C_HDR,
    rfc_empresa:     str = "",
    polizas:         Optional[list] = None,
    numero_cuenta:   str = "",
    meses_iva:       list[int] | None = None,
) -> str:
    """
    Genera el Excel completo de conciliación.

    Retorna la ruta del archivo generado.
    """
    wb = _nuevo_workbook()

    # UUIDs conciliados por tipo (incluye cfdi_suma del nivel 4)
    usados_r: set[str] = set()
    usados_e: set[str] = set()
    for mc in conciliados:
        if not mc.cfdi:
            continue
        if mc.tipo == "recibido":
            usados_r.add(mc.cfdi.uuid.lower())
            for c in mc.cfdi_suma:
                usados_r.add(c.uuid.lower())
        elif mc.tipo == "emitido":
            usados_e.add(mc.cfdi.uuid.lower())
            for c in mc.cfdi_suma:
                usados_e.add(c.uuid.lower())

    # Lookup uuid → datos bancarios para trazar en las hojas CFDI
    lookup_banco: dict[str, dict] = {}
    for mc in conciliados:
        if not mc.cfdi:
            continue
        info_base = {
            "banco":          nombre_banco,
            "folio":          mc.banco.folio,
            "fecha_banco":    mc.banco.fecha,
            "fecha_pago_nom": mc.banco.fecha_pago_cfdi,
            "monto_banco":    mc.banco.monto,
            "es_suma":        False,
        }
        lookup_banco[mc.cfdi.uuid.lower()] = info_base
        for c_extra in mc.cfdi_suma:
            lookup_banco[c_extra.uuid.lower()] = {**info_base, "es_suma": True}

    # Separar nómina y complementos de pago del display principal
    nomina_cfdis      = [c for c in todos_recibidos if c.tipo == "N"]
    recibidos_display = [c for c in todos_recibidos if c.tipo not in ("N", "P")]
    emitidos_display  = [c for c in todos_emitidos  if c.tipo != "P"]

    # Complementos de pago: mapa uuid_factura → [tipo P] y listas separadas
    complementos_map  = _build_complementos_map(todos_recibidos, todos_emitidos)
    tipo_p_recibidos  = [c for c in todos_recibidos if c.tipo == "P"]
    tipo_p_emitidos   = [c for c in todos_emitidos  if c.tipo == "P"]
    todos_por_uuid    = {c.uuid.lower(): c for c in todos_recibidos + todos_emitidos if c.uuid}

    # Validación fiscal por lote — incluye verificación de régimen contra CSF de empresa
    val_map = validar_lote(todos_recibidos + todos_emitidos, rfc_empresa=rfc_empresa)

    # 1. Hoja banco
    _n_banco = _hoja_banco(wb, nombre_banco, conciliados, sin_cfdi, color_banco,
                           numero_cuenta=numero_cuenta)
    _hojas_banco = [(nombre_banco[:31], _n_banco)]

    # 2. Hoja Recibidos (excluye nómina y tipo P)
    if recibidos_display:
        _hoja_cfdis(wb, "Recibidos", recibidos_display,
                    usados_r, C_HDR_R, es_emitido=False,
                    lookup=lookup_banco, val_map=val_map,
                    complementos_map=complementos_map)

    # 3. Hoja Emitidos (excluye complementos de pago tipo P)
    if emitidos_display:
        _hoja_cfdis(wb, "Emitidos", emitidos_display,
                    usados_e, C_HDR_E, es_emitido=True,
                    lookup=lookup_banco, val_map=val_map,
                    complementos_map=complementos_map)

    # 4. Hoja Nómina (solo tipo N)
    _hoja_nomina(wb, nomina_cfdis, usados_r, lookup_banco, val_map=val_map)

    # 5. Hojas Complementos de Pago — separadas por dirección
    _hoja_complementos_pago(wb, tipo_p_emitidos, todos_por_uuid,
                             usados_r, usados_e, lookup_banco, val_map=val_map,
                             titulo="Comp. Pago Emitidos", color_hdr=C_HDR_E)
    _hoja_complementos_pago(wb, tipo_p_recibidos, todos_por_uuid,
                             usados_r, usados_e, lookup_banco, val_map=val_map,
                             titulo="Comp. Pago Recibidos", color_hdr=C_HDR_R)

    # 6. Resumen
    _hoja_resumen(wb, conciliados, sin_cfdi, cfdi_sin_bco_r, cfdi_sin_bco_e,
                  todos_recibidos=todos_recibidos, todos_emitidos=todos_emitidos,
                  val_map=val_map)

    # 7. Cálculo Previo de IVA
    _hojas_iva_previo(wb, todos_recibidos, todos_emitidos, meses_iva=meses_iva)

    # 8. Pólizas COI (opcional — solo si se proveen pólizas generadas)
    if polizas:
        _hoja_polizas_coi(wb, polizas)

    # Al final: necesita que las hojas de banco ya existan para apuntarles.
    _hoja_tablero(wb, _hojas_banco)

    return _guardar_workbook(wb, output_path)


# ── Exportación multi-banco ────────────────────────────────────────────────────

def exportar_conciliacion_multi(
    resultados:      list[ResultadoBanco],
    todos_recibidos: list[CFDI],
    todos_emitidos:  list[CFDI],
    output_path:     str | Path,
    colores_banco:   Optional[dict[str, str]] = None,
    rfc_empresa:     str = "",
    polizas:         Optional[list] = None,
    cuentas_banco:   Optional[dict[str, str]] = None,
    meses_iva:       list[int] | None = None,
) -> str:
    """
    Genera el Excel de conciliación para múltiples estados de cuenta.

    Estructura del Excel:
      • Una hoja por banco   — movimientos de ese estado de cuenta
      • Recibidos            — todos los CFDIs recibidos con columna BANCO
      • Emitidos             — todos los CFDIs emitidos  con columna BANCO
      • Nomina               — todos los CFDIs de nómina con columna BANCO
      • Resumen              — totales globales y por banco

    Un CFDI no puede aparecer conciliado en más de un banco.

    colores_banco: dict nombre_banco → hex color (opcional)
    """
    wb = _nuevo_workbook()

    _COLORES_DEFAULT = ["1565C0", "2E7D32", "6A1B9A", "BF360C",
                        "00695C", "4527A0", "558B2F", "AD1457"]
    colores = colores_banco or {}

    _MESES_ABREV = {1:"ENE",2:"FEB",3:"MAR",4:"ABR",5:"MAY",6:"JUN",
                    7:"JUL",8:"AGO",9:"SEP",10:"OCT",11:"NOV",12:"DIC"}

    def _mes_de_fecha(s: str) -> int:
        """Extrae el mes de 'DD/MM/YYYY', 'YYYY-MM-DD' o 'DD-MM-YYYY'."""
        s = (s or "").strip()
        try:
            if len(s) >= 10 and s[2] in "/":   # DD/MM/YYYY
                return int(s[3:5])
            if len(s) >= 10 and s[4] == "-":   # YYYY-MM-DD
                return int(s[5:7])
            if len(s) >= 10 and s[2] == "-":   # DD-MM-YYYY
                return int(s[3:5])
        except (ValueError, IndexError):
            pass
        return 0

    def _mes_res(res) -> int:
        """Mes predominante del resultado (usa mes o parsea fecha como fallback)."""
        meses = []
        for mb in ([mc.banco for mc in res.conciliados] + res.sin_cfdi):
            m = mb.mes or _mes_de_fecha(mb.fecha)
            if m:
                meses.append(m)
        return Counter(meses).most_common(1)[0][0] if meses else 0

    # ── Nombres de hoja únicos: BANCO MES (p.ej. "SANTANDER ENE", "BBVA FEB") ──
    nombres_hoja: list[str] = []
    _conteo: dict[str, int] = {}
    for res in resultados:
        mes_num = _mes_res(res)
        mes_sfx = f" {_MESES_ABREV[mes_num]}" if mes_num in _MESES_ABREV else ""
        base    = (res.nombre[:23] + mes_sfx)[:27]
        _conteo[base] = _conteo.get(base, 0) + 1
        nombres_hoja.append(base if _conteo[base] == 1
                            else f"{base} ({_conteo[base]})")

    # ── Lookup global uuid → traza bancaria ───────────────────────────────────
    lookup_banco: dict[str, dict] = {}
    usados_r: set[str] = set()
    usados_e: set[str] = set()

    for idx, (res, nombre_hoja) in enumerate(zip(resultados, nombres_hoja)):
        color = colores.get(res.nombre, _COLORES_DEFAULT[idx % len(_COLORES_DEFAULT)])
        for mc in res.conciliados:
            if not mc.cfdi:
                continue
            info_base = {
                "banco":          nombre_hoja,
                "folio":          mc.banco.folio,
                "fecha_banco":    mc.banco.fecha,
                "fecha_pago_nom": mc.banco.fecha_pago_cfdi,
                "monto_banco":    mc.banco.monto,
                "es_suma":        False,
            }
            uuid_k = mc.cfdi.uuid.lower()
            lookup_banco[uuid_k] = info_base

            # CFDIs adicionales de suma nivel 4
            for c_extra in mc.cfdi_suma:
                lookup_banco[c_extra.uuid.lower()] = {**info_base, "es_suma": True}

            if mc.tipo == "recibido":
                usados_r.add(uuid_k)
                for c in mc.cfdi_suma:
                    usados_r.add(c.uuid.lower())
            elif mc.tipo == "emitido":
                usados_e.add(uuid_k)
                for c in mc.cfdi_suma:
                    usados_e.add(c.uuid.lower())

    # ── 1. Una hoja por banco (nombre de hoja garantizado único) ─────────────
    _cuentas = cuentas_banco or {}
    _hojas_banco: list[tuple[str, int]] = []
    for idx, (res, nombre_hoja) in enumerate(zip(resultados, nombres_hoja)):
        color = colores.get(res.nombre, _COLORES_DEFAULT[idx % len(_COLORES_DEFAULT)])
        _n = _hoja_banco(wb, nombre_hoja, res.conciliados, res.sin_cfdi, color,
                         numero_cuenta=_cuentas.get(res.nombre, ""))
        _hojas_banco.append((nombre_hoja[:31], _n))

    # ── 2-4. Hojas de CFDIs (compartidas, con columna BANCO) ──────────────────
    nomina_cfdis      = [c for c in todos_recibidos if c.tipo == "N"]
    recibidos_display = [c for c in todos_recibidos if c.tipo not in ("N", "P")]
    emitidos_display  = [c for c in todos_emitidos  if c.tipo != "P"]

    complementos_map = _build_complementos_map(todos_recibidos, todos_emitidos)
    tipo_p_recibidos = [c for c in todos_recibidos if c.tipo == "P"]
    tipo_p_emitidos  = [c for c in todos_emitidos  if c.tipo == "P"]
    todos_por_uuid   = {c.uuid.lower(): c for c in todos_recibidos + todos_emitidos if c.uuid}

    val_map = validar_lote(todos_recibidos + todos_emitidos, rfc_empresa=rfc_empresa)

    if recibidos_display:
        _hoja_cfdis(wb, "Recibidos", recibidos_display,
                    usados_r, C_HDR_R, es_emitido=False,
                    lookup=lookup_banco, val_map=val_map,
                    complementos_map=complementos_map)
    if emitidos_display:
        _hoja_cfdis(wb, "Emitidos", emitidos_display,
                    usados_e, C_HDR_E, es_emitido=True,
                    lookup=lookup_banco, val_map=val_map,
                    complementos_map=complementos_map)
    _hoja_nomina(wb, nomina_cfdis, usados_r, lookup_banco, val_map=val_map)

    # ── 5. Hojas Complementos de Pago — separadas por dirección ─────────────
    _hoja_complementos_pago(wb, tipo_p_emitidos, todos_por_uuid,
                             usados_r, usados_e, lookup_banco, val_map=val_map,
                             titulo="Comp. Pago Emitidos", color_hdr=C_HDR_E)
    _hoja_complementos_pago(wb, tipo_p_recibidos, todos_por_uuid,
                             usados_r, usados_e, lookup_banco, val_map=val_map,
                             titulo="Comp. Pago Recibidos", color_hdr=C_HDR_R)

    # ── 6. Resumen global + por banco ─────────────────────────────────────────
    _hoja_resumen_multi(wb, resultados, nombres_hoja,
                        todos_recibidos, todos_emitidos, val_map=val_map)

    # ── 7. Cálculo Previo de IVA ────────────────────────────────────────────
    _hojas_iva_previo(wb, todos_recibidos, todos_emitidos, meses_iva=meses_iva)

    # ── 8. Pólizas COI (opcional) ─────────────────────────────────────────────
    if polizas:
        _hoja_polizas_coi(wb, polizas)

    # Al final: necesita que las hojas de banco ya existan para apuntarles.
    _hoja_tablero(wb, _hojas_banco)

    return _guardar_workbook(wb, output_path)


def _hoja_resumen_multi(wb,
                        resultados:      list[ResultadoBanco],
                        nombres_hoja:    list[str],
                        todos_recibidos: list[CFDI],
                        todos_emitidos:  list[CFDI],
                        val_map:         Optional[dict] = None) -> None:
    ws = wb.create_sheet("Resumen")
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 18

    def hdr(row, txt, bg=C_HDR_S):
        ws.merge_cells(f"A{row}:B{row}")
        c = ws.cell(row=row, column=1, value=txt)
        c.font      = Font(bold=True, size=10, color="FFFFFF")
        c.fill      = PatternFill("solid", fgColor=bg)
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[row].height = 18

    def fila(row, label, val, bg=C_WHITE, fmt=None):
        _cell(ws, row, 1, label, bold=True, bg=bg, size=9)
        _cell(ws, row, 2, val,   bg=bg, align="right", fmt=fmt, size=9)
        ws.row_dimensions[row].height = 16

    row = 1
    ws.merge_cells(f"A{row}:B{row}")
    c = ws.cell(row=row, column=1, value="Resumen de Conciliacion Multi-Banco")
    c.font = Font(bold=True, size=13, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor=C_HDR_S)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[row].height = 26
    row += 1

    # Totales globales
    total_conc   = sum(len(r.conciliados) for r in resultados)
    total_rev    = sum(sum(1 for mc in r.conciliados if mc.revisar) for r in resultados)
    total_pend   = sum(len(r.sin_cfdi) for r in resultados)
    total_mov    = total_conc + total_pend
    monto_conc   = sum(mc.banco.monto for r in resultados for mc in r.conciliados)
    monto_pend   = sum(mb.monto for r in resultados for mb in r.sin_cfdi)

    hdr(row, "GLOBAL"); row += 1
    fila(row, "Movimientos totales",          total_mov);  row += 1
    fila(row, "Conciliados",                  total_conc,  C_CONC);    row += 1
    fila(row, "Requieren revision",           total_rev,   C_REVISAR); row += 1
    fila(row, "Sin CFDI",                     total_pend,  C_SIN);     row += 1
    fila(row, "Monto conciliado ($)",         monto_conc,  C_CONC,    MONEY_FMT); row += 1
    fila(row, "Monto pendiente ($)",          monto_pend,  C_SIN,     MONEY_FMT); row += 1
    fila(row, "CFDIs recibidos sin banco",
         sum(1 for c in todos_recibidos if c.tipo != "N"), C_PEND); row += 1
    fila(row, "CFDIs emitidos sin banco",
         len(todos_emitidos), C_PEND); row += 1
    row += 1

    # Complementos de Pago (PPD / tipo P)
    todos_i_all = todos_recibidos + todos_emitidos
    todos_ppd_m = [c for c in todos_i_all if c.tipo == "I" and c.metodo_pago == "PPD"]
    todos_cp_m  = [c for c in todos_i_all if c.tipo == "P"]
    fp_set_m    = {uid.lower() for cp in todos_cp_m for uid in cp.cfdi_facturas_pagadas}
    n_ppd_m      = len(todos_ppd_m)
    n_ppd_liq_m  = sum(1 for c in todos_ppd_m if c.uuid.lower() in fp_set_m)
    n_ppd_pend_m = n_ppd_m - n_ppd_liq_m
    n_cp_m       = len(todos_cp_m)
    if n_ppd_m or n_cp_m:
        hdr(row, "COMPLEMENTOS DE PAGO (PPD / TIPO P)", bg=C_HDR_P); row += 1
        fila(row, "Facturas PPD (pago diferido)",           n_ppd_m,       C_PEND); row += 1
        fila(row, "  PPD Liquidadas (con complemento P)",   n_ppd_liq_m,   C_CONC); row += 1
        fila(row, "  PPD Pendientes (sin complemento P)",   n_ppd_pend_m,  C_SIN);  row += 1
        fila(row, "Complementos de Pago (tipo P)",          n_cp_m,        C_SAL);  row += 1
        row += 1

    # Validación fiscal — incluye todos los tipos (I, E, P, N).
    # Usa val_map ya computado con rfc_empresa para que R006 compare contra
    # el catálogo de la empresa.
    all_cfdi = [c for c in todos_recibidos + todos_emitidos if c.uuid]
    if all_cfdi:
        v_lote = val_map if val_map is not None else validar_lote(all_cfdi)
        n_verde    = sum(1 for v in v_lote.values() if v.estatus == "VERDE")
        n_amarillo = sum(1 for v in v_lote.values() if v.estatus == "AMARILLO")
        n_rojo     = sum(1 for v in v_lote.values() if v.estatus == "ROJO")
        n_negro    = sum(1 for v in v_lote.values() if v.estatus == "NEGRO")
        hdr(row, "VALIDACION FISCAL"); row += 1
        fila(row, "Deducibles sin errores (VERDE)",      n_verde,    C_VAL_VERDE);    row += 1
        fila(row, "Requieren revision (AMARILLO)",       n_amarillo, C_VAL_AMARILLO); row += 1
        fila(row, "No deducibles — errores (ROJO)",      n_rojo,     C_VAL_ROJO);     row += 1
        fila(row, "Sin efecto fiscal — critico (NEGRO)", n_negro,    C_VAL_NEGRO);    row += 1
        row += 1

    # Por banco
    hdr(row, "POR BANCO"); row += 1
    for res, nombre_hoja in zip(resultados, nombres_hoja):
        conc_b  = len(res.conciliados)
        pend_b  = len(res.sin_cfdi)
        monto_b = sum(mc.banco.monto for mc in res.conciliados)
        hdr(row, nombre_hoja, bg=C_HDR); row += 1
        fila(row, "  Movimientos",    conc_b + pend_b); row += 1
        fila(row, "  Conciliados",    conc_b, C_CONC); row += 1
        fila(row, "  Sin CFDI",       pend_b, C_SIN);  row += 1
        fila(row, "  Monto conciliado ($)", monto_b, C_CONC, MONEY_FMT); row += 1
        row += 1
