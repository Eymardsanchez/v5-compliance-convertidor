"""
Extractor para estados de cuenta Banregio (Banco Regional S.A.).
Formato: DIA  TIPO  CONCEPTO  CARGOS  ABONOS  SALDO

Columnas (página 612 pt):
  DIA         x0 ∈  [34, 52)
  TIPO/REF    x0 ∈  [52, 70)
  DESCRIPCIÓN x0 ∈  [70, 370)
  CARGOS      x0 ∈ [370, 455)
  ABONOS      x0 ∈ [455, 525)
  SALDO       x0 ≥ 525
"""

import re
from typing import Optional, Callable

import pdfplumber

from bancos.base import BancoExtractor, Movimiento, InfoCuenta


MESES_ESP = {
    "ENERO": "01", "FEBRERO": "02", "MARZO": "03", "ABRIL": "04",
    "MAYO": "05", "JUNIO": "06", "JULIO": "07", "AGOSTO": "08",
    "SEPTIEMBRE": "09", "OCTUBRE": "10", "NOVIEMBRE": "11", "DICIEMBRE": "12",
}

AMOUNT_RE  = re.compile(r"^[\d,]+\.\d{2}$")
DAY_RE     = re.compile(r"^\d{1,2}$")
PERIOD_RE  = re.compile(
    r"del\s+\d+\s+al\s+\d+\s+de\s+([A-ZÁÉÍÓÚ]+)\s+(\d{4})", re.IGNORECASE
)
SECTION_RE = re.compile(r"DIA\s+CONCEPTO", re.IGNORECASE)

# Umbrales de columna (pt)
_X_DAY_MAX   =  52   # día:        x0 < 52
_X_TYPE_MAX  =  70   # tipo/ref:   x0 ∈ [52, 70)
_X_DESC_MAX  = 370   # descripción: x0 < 370
_X_CARGO_MAX = 455   # cargos:      x0 ∈ [370, 455)
_X_ABONO_MAX = 525   # abonos:      x0 ∈ [455, 525)
#                      saldo:        x0 ≥ 525

SKIP_WORDS = {
    "BANCO", "REGIONAL", "ESTADO", "CUENTA", "PARQUE", "CORPORATIVO",
    "SAN", "PEDRO", "GARZA", "R.F.C.", "CENTRO", "PAGE", "SALDO",
    "MINIMO", "GRAFICO", "TRANSACCIONAL", "MANDATO", "MENSAJES",
    "IMPORTANTES", "ABREVIATURAS", "INFORMACION", "RELEVANTE",
    "DIA", "CONCEPTO", "CARGOS", "ABONOS", "SALDO",
}

END_SECTION_PATTERNS = [
    "MENSAJES IMPORTANTES",
    "ABREVIATURAS",
    "INFORMACION RELEVANTE",
    "GRAFICO TRANSACCIONAL",
    "SALDO MINIMO REQUERIDO",
    "SALDO MÍNIMO REQUERIDO",
    "MANDATO ESPECIAL",
]


def _parse_period(text: str) -> tuple[str, str]:
    """Extrae mes y año del encabezado. Retorna (mes_num, anio)."""
    m = PERIOD_RE.search(text)
    if m:
        mes_str = m.group(1).upper()
        anio    = m.group(2)
        for key, val in MESES_ESP.items():
            if key in mes_str:
                return val, anio
    return "01", "2026"


def _parse_amount(text: str) -> Optional[float]:
    clean = text.replace(",", "").strip()
    try:
        return float(clean)
    except ValueError:
        return None


def _is_amount(text: str) -> bool:
    return bool(AMOUNT_RE.match(text))


def _group_rows(words: list[dict]) -> dict[int, list[dict]]:
    rows: dict[int, list[dict]] = {}
    for w in words:
        y = round(w["top"] / 2) * 2
        rows.setdefault(y, []).append(w)
    return rows


class BanregioExtractor(BancoExtractor):
    NOMBRE    = "Banregio"
    COLOR_HEX = "005A8E"

    @classmethod
    def detectar(cls, texto: str) -> bool:
        return "BANREGIO" in texto or "BANCO REGIONAL" in texto

    def extraer(
        self,
        pdf_path: str,
        progress_callback: Optional[Callable[[int, int], None]] = None,
    ) -> tuple[list[Movimiento], InfoCuenta]:
        movimientos: list[Movimiento] = []
        current: Optional[Movimiento] = None
        info    = InfoCuenta(banco="Banregio")
        in_section = False
        mes_num    = "01"
        anio       = "2026"

        with pdfplumber.open(pdf_path) as pdf:
            total = len(pdf.pages)

            for pnum, page in enumerate(pdf.pages):
                if progress_callback:
                    progress_callback(pnum + 1, total)

                # ── Encabezado (solo página 1 con extract_text) ───────────────
                if pnum == 0:
                    text0 = page.extract_text(x_tolerance=2, y_tolerance=3) or ""
                    mes_num, anio = _parse_period(text0)
                    self._parse_header(text0, info)
                    lines0 = text0.splitlines()
                    info.periodo = lines0[1].strip() if len(lines0) > 1 else ""

                # ── Movimientos con extract_words (posición exacta) ───────────
                words = page.extract_words(x_tolerance=2, y_tolerance=3)
                rows  = _group_rows(words)

                for y in sorted(rows.keys()):
                    ws = sorted(rows[y], key=lambda w: w["x0"])

                    # Detectar encabezado de sección "DIA CONCEPTO ..."
                    row_text = " ".join(w["text"] for w in ws).upper()
                    if "DIA" in row_text and "CONCEPTO" in row_text:
                        in_section = True
                        continue

                    # Detectar marcadores de fin de sección
                    if in_section and any(
                        p.upper() in row_text for p in END_SECTION_PATTERNS
                    ):
                        if current:
                            movimientos.append(current)
                            current = None
                        in_section = False
                        continue

                    # ── Saldo inicial (puede aparecer ANTES del encabezado DIA) ──
                    if not in_section:
                        if "INICIAL" in row_text and info.saldo_inicial is None:
                            saldo_ws_pre = [w for w in ws if w["x0"] >= _X_ABONO_MAX
                                            and _is_amount(w["text"])]
                            if saldo_ws_pre:
                                v = _parse_amount(saldo_ws_pre[0]["text"])
                                if v is not None:
                                    info.saldo_inicial = v
                        continue

                    # ── Clasificar palabras de la fila por columna ────────────
                    day_ws   = [w for w in ws if w["x0"] <  _X_DAY_MAX  and DAY_RE.match(w["text"])]
                    desc_ws  = [w for w in ws if _X_TYPE_MAX <= w["x0"] <  _X_DESC_MAX
                                and not _is_amount(w["text"])]
                    cargo_ws = [w for w in ws if _X_DESC_MAX <= w["x0"] < _X_CARGO_MAX
                                and _is_amount(w["text"])]
                    abono_ws = [w for w in ws if _X_CARGO_MAX <= w["x0"] < _X_ABONO_MAX
                                and _is_amount(w["text"])]
                    saldo_ws = [w for w in ws if w["x0"] >= _X_ABONO_MAX
                                and _is_amount(w["text"])]

                    # ── Fila de totales: cargo Y abono simultáneos → fin de sección ──
                    if in_section and not day_ws and cargo_ws and abono_ws:
                        if current:
                            movimientos.append(current)
                            current = None
                        in_section = False
                        continue

                    # ── Nueva transacción ─────────────────────────────────────
                    if day_ws:
                        if current:
                            movimientos.append(current)

                        dia   = day_ws[0]["text"].zfill(2)
                        fecha = f"{dia}/{mes_num}/{anio}"
                        desc  = " ".join(w["text"] for w in desc_ws).strip()
                        desc  = re.sub(r"\s+", " ", desc).strip()

                        retiro   = _parse_amount(cargo_ws[0]["text"]) if cargo_ws else None
                        deposito = _parse_amount(abono_ws[0]["text"]) if abono_ws else None
                        saldo    = _parse_amount(saldo_ws[0]["text"]) if saldo_ws else None

                        current = Movimiento(
                            fecha=fecha,
                            descripcion=desc,
                            retiro=retiro,
                            deposito=deposito,
                            saldo=saldo,
                        )

                    elif current and desc_ws:
                        # Línea de continuación de descripción
                        extra = " ".join(w["text"] for w in desc_ws).strip()
                        if extra:
                            current.descripcion = (current.descripcion + " " + extra).strip()
                        # Completar montos si llegaron en fila separada
                        if not current.retiro and cargo_ws:
                            current.retiro = _parse_amount(cargo_ws[0]["text"])
                        if not current.deposito and abono_ws:
                            current.deposito = _parse_amount(abono_ws[0]["text"])
                        if saldo_ws:
                            current.saldo = _parse_amount(saldo_ws[0]["text"])

            if current:
                movimientos.append(current)

        if movimientos:
            info.total_depositos = sum(m.deposito or 0 for m in movimientos)
            info.total_retiros   = sum(m.retiro   or 0 for m in movimientos)
            info.saldo_final     = movimientos[-1].saldo

        return movimientos, info

    def _parse_header(self, text: str, info: InfoCuenta):
        for line in text.splitlines():
            line = line.strip()
            if "RFC:" in line.upper():
                m = re.search(r"RFC:\s*([A-Z0-9]{10,13})", line.upper())
                if m:
                    info.rfc = m.group(1)
            if "CLIENTE :" in line.upper() or "CLIENTE:" in line.upper():
                m = re.search(r"CLIENTE\s*:\s*(\S+)", line.upper())
                if m:
                    info.numero_cuenta = m.group(1)
            if re.search(r"\d{3}-\d{5}-\d{3}-\d", line):
                m = re.search(r"(\d{3}-\d{5}-\d{3}-\d)", line)
                if m:
                    info.numero_cuenta = m.group(1)
            m = re.search(r"\b(\d{18})\b", line)
            if m:
                info.clabe = m.group(1)
