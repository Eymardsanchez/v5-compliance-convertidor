"""
Extractor para estados de cuenta Banamex / Citibanamex.
Formato: tabla con columnas RETIROS | DEPÓSITOS | SALDO.
Las transacciones pueden ser multi-línea; las clasificación DEP/RET se determina
directamente por la columna del monto (x0 en el PDF) sin necesidad de balance tracking.

Columnas (página 612 pt):
  RETIROS    x0 ∈ [340, 430)
  DEPÓSITOS  x0 ∈ [430, 515)
  SALDO      x0 ≥ 515
"""

import re
from typing import Optional, Callable

import pdfplumber

from bancos.base import BancoExtractor, Movimiento, InfoCuenta


MESES = {
    "ENE": "01", "FEB": "02", "MAR": "03", "ABR": "04",
    "MAY": "05", "JUN": "06", "JUL": "07", "AGO": "08",
    "SEP": "09", "OCT": "10", "NOV": "11", "DIC": "12",
}

# Fecha en línea de transacción: 1-2 dígitos al extremo izquierdo
_DATE_NUM_RE  = re.compile(r'^\d{1,2}$')
_DATE_MON_RE  = re.compile(r'^[A-ZÁÉÍÓÚ]{3,4}$')
_AMOUNT_RE    = re.compile(r'^[\d,]+\.\d{2}$')
_PERIOD_YEAR  = re.compile(r'\bdel\s+(\d{4})\b', re.IGNORECASE)

# Umbrales de columna (pt) – página de 612 pt
# Día:  x0 ∈ [22, 40]  (ej. "13" → x0=25)
# Mes:  x0 ∈ (40, 80]  (ej. "MAR" → x0=42)
_X_DATE_MAX  = 40   # columna de día: x0 ≤ 40
_X_MON_MAX   = 80   # columna de mes: x0 ∈ (40, 80]
_X_DESC_MAX  = 340  # descripción: x0 < 340
_X_RET_MAX   = 430  # Retiros: x0 ∈ [340, 430)
_X_DEP_MAX   = 515  # Depósitos: x0 ∈ [430, 515)
#                     Saldo: x0 ≥ 515

# Palabras de encabezado/pie que se ignoran
_SKIP_WORDS = {
    "FECHA", "CONCEPTO", "RETIROS", "DEPÓSITOS", "SALDO",
    "DEPOSITOS", "PAGINA", "PÁGINA",
}

# Líneas que marcan fin de sección de operaciones
_END_MARKERS = [
    "TOTAL DE OPERACIONES",
    "TOTALES DEL PERIODO",
    "SALDO AL CORTE",
    "SALDO FINAL",
]


def _group_rows(words: list[dict]) -> dict[int, list[dict]]:
    rows: dict[int, list[dict]] = {}
    for w in words:
        y = round(w['top'] / 2) * 2
        rows.setdefault(y, []).append(w)
    return rows


def _parse_amount(text: str) -> Optional[float]:
    clean = text.replace(',', '').strip()
    try:
        return float(clean)
    except ValueError:
        return None


def _is_amount(text: str) -> bool:
    return bool(_AMOUNT_RE.match(text.replace(',', '').replace('.', '', 1).replace(',', ''))) or \
           bool(_AMOUNT_RE.match(text))


class BanamexExtractor(BancoExtractor):
    """Extractor para estados de cuenta Banamex / Citibanamex."""

    NOMBRE    = "Banamex"
    COLOR_HEX = "C8001A"   # Rojo corporativo Banamex

    @classmethod
    def detectar(cls, texto: str) -> bool:
        # Buscar SÓLO en el encabezado del documento (primeros 2 000 chars).
        # "BANAMEX" aparece frecuentemente en descripciones de SPEI de otros
        # bancos (p. ej. "SPEI ENVIADO BANAMEX"), por lo que buscar en el texto
        # completo produce falsos positivos en PDFs de BBVA, Santander, etc.
        # 2 000 chars cubre ~1 página completa incluso cuando la página 1 es
        # una imagen en blanco (BBVA) y el encabezado real está en la pág. 2.
        cabecera = texto[:2000].upper()
        return (
            "BANAMEX" in cabecera
            or "CITIBANAMEX" in cabecera
            or "MICUENTA" in cabecera
            or "MICUENTA" in cabecera.replace(" ", "")
            or ("CITI" in cabecera and "BANCO NACIONAL DE MEXICO" in cabecera)
            or "BANCO NACIONAL DE MEXICO" in cabecera
        )

    # ── extracción principal ──────────────────────────────────────────────────

    def extraer(
        self,
        pdf_path: str,
        progress_callback: Optional[Callable[[int, int], None]] = None,
    ) -> tuple[list[Movimiento], InfoCuenta]:

        info = InfoCuenta(banco="Banamex")
        movimientos: list[Movimiento] = []
        anio = ""

        with pdfplumber.open(pdf_path) as pdf:
            total = len(pdf.pages)

            # ── Paso 1: metadatos de páginas 1 y 2 ───────────────────────────
            # Banamex coloca datos de cuenta en pág. 1 y el resumen (Saldo anterior,
            # Depósitos, Retiros, Saldo al corte) normalmente en pág. 2.
            text0 = pdf.pages[0].extract_text(x_tolerance=2, y_tolerance=3) or ""
            text1 = (pdf.pages[1].extract_text(x_tolerance=2, y_tolerance=3) or "") \
                    if len(pdf.pages) > 1 else ""
            # Parsear primero pág. 1 y luego pág. 2 para que los datos de pág. 2
            # complementen (no sobreescriban) los de pág. 1.
            self._parse_header(text0, info)
            self._parse_header(text1, info)

            # Año del período (buscar en ambas páginas)
            combined_text = text0 + "\n" + text1
            m_year = _PERIOD_YEAR.search(combined_text)
            if m_year:
                anio = m_year.group(1)
            else:
                # Buscar año de 4 dígitos 20XX
                m2 = re.search(r'\b(20\d{2})\b', combined_text)
                anio = m2.group(1) if m2 else "2026"

            # ── Paso 2: movimientos página a página ───────────────────────────
            current: Optional[Movimiento] = None

            for pnum, page in enumerate(pdf.pages):
                if progress_callback:
                    progress_callback(pnum + 1, total)

                words = page.extract_words(x_tolerance=2, y_tolerance=3)
                rows  = _group_rows(words)

                for y in sorted(rows.keys()):
                    ws = sorted(rows[y], key=lambda w: w['x0'])
                    texts_all = [w['text'].strip() for w in ws if w['text'].strip()]

                    # Detectar fin de sección
                    row_text = ' '.join(texts_all).upper()

                    # Saltar encabezados de columna y sección de resumen
                    if texts_all and texts_all[0].upper() in _SKIP_WORDS:
                        continue
                    # Saltar líneas de resumen que no son transacciones
                    if any(row_text.startswith(m) for m in [
                        "SALDO AL CORTE", "SALDO FINAL", "TOTAL DE OPERACIONES",
                        "TOTALES DEL PERIODO",
                        "(+)", "(-)", "INTERESES", "IMPUESTO SOBRE",
                    ]):
                        continue

                    # Detectar fila "SALDO ANTERIOR"
                    if ('SALDO' in row_text and 'ANTERIOR' in row_text):
                        sal_words = [w for w in ws if w['x0'] >= _X_DEP_MAX]
                        if sal_words:
                            v = _parse_amount(sal_words[0]['text'])
                            if v is not None and info.saldo_inicial is None:
                                info.saldo_inicial = v
                        continue

                    # ── Analizar columnas de la fila ─────────────────────────
                    date_ws  = [w for w in ws if w['x0'] <= _X_DATE_MAX]
                    month_ws = [w for w in ws if _X_DATE_MAX < w['x0'] <= _X_MON_MAX]
                    desc_ws  = [w for w in ws if _X_MON_MAX < w['x0'] < _X_DESC_MAX]
                    ret_ws   = [w for w in ws if _X_DESC_MAX <= w['x0'] < _X_RET_MAX
                                and re.match(r'^[\d,]+\.\d{2}$', w['text'])]
                    dep_ws   = [w for w in ws if _X_RET_MAX <= w['x0'] < _X_DEP_MAX
                                and re.match(r'^[\d,]+\.\d{2}$', w['text'])]
                    sal_ws   = [w for w in ws if w['x0'] >= _X_DEP_MAX
                                and re.match(r'^[\d,]+\.\d{2}$', w['text'])]

                    # ¿Nueva fecha?
                    if date_ws and _DATE_NUM_RE.match(date_ws[0]['text']):
                        # Guardar transacción previa
                        if current and current.saldo is not None:
                            movimientos.append(current)

                        day = date_ws[0]['text'].zfill(2)
                        mes = "01"
                        if month_ws and _DATE_MON_RE.match(month_ws[0]['text'].upper()):
                            mes = MESES.get(month_ws[0]['text'].upper()[:3], "01")

                        fecha = f"{day}/{mes}/{anio}"
                        desc  = ' '.join(w['text'] for w in desc_ws).strip()
                        current = Movimiento(fecha=fecha, descripcion=desc)

                    elif current and desc_ws:
                        # Continuación de descripción
                        extra = ' '.join(w['text'] for w in desc_ws).strip()
                        if extra:
                            current.descripcion = (current.descripcion + ' ' + extra).strip()

                    # Asignar montos al movimiento actual (o al que acaba de crearse)
                    if current:
                        if ret_ws and current.retiro is None:
                            current.retiro = _parse_amount(ret_ws[0]['text'])
                        if dep_ws and current.deposito is None:
                            current.deposito = _parse_amount(dep_ws[0]['text'])
                        if sal_ws:
                            # Actualizar saldo con el último valor encontrado
                            v = _parse_amount(sal_ws[0]['text'])
                            if v is not None:
                                current.saldo = v

            # Guardar último movimiento si quedó pendiente
            if current and current.saldo is not None:
                movimientos.append(current)

        # Limpiar descripciones
        for mov in movimientos:
            mov.descripcion = re.sub(r'\s+', ' ', mov.descripcion).strip()

        if movimientos:
            info.total_depositos = sum(m.deposito or 0 for m in movimientos)
            info.total_retiros   = sum(m.retiro   or 0 for m in movimientos)
            info.saldo_final     = movimientos[-1].saldo

        return movimientos, info

    # ── parsing de encabezado ─────────────────────────────────────────────────

    def _parse_header(self, text: str, info: InfoCuenta) -> None:
        lines = text.splitlines()
        for line in lines:
            line = line.strip()
            if not line:
                continue

            # Período  — buscar "Período del NN de mes al NN de mes del YYYY"
            if (re.search(r'[Pp]er[íi]odo\s+del\s+\d', line) or
                    re.search(r'[Rr]esumen\s+del\s+\d', line)):
                if not info.periodo:
                    info.periodo = line.strip()

            # RFC
            m = re.search(r'\bRFC\s+([A-Z0-9]{10,13})\b', line, re.IGNORECASE)
            if m:
                info.rfc = m.group(1)

            # CLABE
            m = re.search(r'\b(\d{18})\b', line)
            if m:
                info.clabe = m.group(1)

            # Número de cuenta de cheques
            m = re.search(r'cuenta\s+de\s+cheques\s+(\d{6,12})', line, re.IGNORECASE)
            if m:
                info.numero_cuenta = m.group(1)

            # Número de contrato como fallback
            if not info.numero_cuenta:
                m = re.search(r'contrato\s+(\d{8,12})', line, re.IGNORECASE)
                if m:
                    info.numero_cuenta = m.group(1)

            # Titular
            if re.search(r'titular|nombre', line, re.IGNORECASE):
                parts = line.split(':', 1)
                if len(parts) == 2:
                    info.titular = parts[1].strip()

            # Saldo inicial / anterior
            if 'saldo anterior' in line.lower():
                m = re.search(r'([\d,]+\.\d{2})', line)
                if m and info.saldo_inicial is None:
                    try:
                        info.saldo_inicial = float(m.group(1).replace(',', ''))
                    except ValueError:
                        pass
