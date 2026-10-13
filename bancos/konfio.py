"""
Extractor para estados de cuenta Konfío (Tarjeta de Crédito).
Emisor: RED AMIGO DAL, S.A.P.I. DE C.V. SOFOM, E.N.R.

Formato de tabla (carta, ≈612 pt de ancho):
  Fecha           x0 <  115        "DD mmm YYYY"  (día/mes/año en fila separada de la descripción)
  Descripción     x0 ∈  [115, 310)
  Tarjeta         x0 ∈  [310, 365) últimos 4 dígitos + FÍSICA/DIGITAL
  Cargo M.N.      x0 ∈  [365, 530) → retiro  (la empresa gastó)
  Abono M.N.      x0 ≥  530        → deposito (devolución / bonificación)

NOTA: cada movimiento ocupa 2–3 filas:
  Fila A (sin fecha): descripción del comercio + número de tarjeta
  Fila B (con fecha): fecha + monto cargo/abono
  Fila C (opcional):  RFC del comercio + fecha referencia + tipo tarjeta

La fuente estilizada del PDF (encabezados, títulos) duplica CADA carácter:
  "SSTTEERRIIMMEEXX" = "STERIMEX"
La función _dedup_chars() detecta texto completamente duplicado y lo corrige.
El texto de transacciones (descripciones, RFCs, montos) viene en fuente normal.
"""

import re
from typing import Optional, Callable

import pdfplumber

from bancos.base import BancoExtractor, Movimiento, InfoCuenta


# ── Constantes ────────────────────────────────────────────────────────────────

MESES_ABR: dict[str, str] = {
    "ENE": "01", "FEB": "02", "MAR": "03", "ABR": "04",
    "MAY": "05", "JUN": "06", "JUL": "07", "AGO": "08",
    "SEP": "09", "OCT": "10", "NOV": "11", "DIC": "12",
}

# "21 ABR 2026"  →  grupo 1=día  2=mes  3=año
_FECHA_RE = re.compile(
    r"^(\d{1,2})\s+(ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC)\s+(\d{4})$",
    re.IGNORECASE,
)

# Monto con o sin símbolo $: "921.66" | "$4,264.00"
_AMOUNT_RE = re.compile(r"^\$?([\d,]+\.\d{2})$")

# RFC partido por espacio: "ESP 060105LJ5" → "ESP060105LJ5"
_RFC_ESPACIO_RE = re.compile(
    r"\b([A-ZÑ&]{3,4})\s+([0-9]{6}[A-Z0-9]{3})\b", re.IGNORECASE
)

# ── Umbrales de columna (pt) ──────────────────────────────────────────────────
_X_FECHA_MAX  = 115    # Fecha:        x0 < 115  (día ~54, mes ~67, año ~84)
_X_DESC_MAX   = 310    # Descripción:  115 ≤ x0 < 310
_X_CARD_MAX   = 350    # Tarjeta:      310 ≤ x0 < 350  (4 dígitos ~311–338, FÍSICA ~314–349)
_X_CARGO_MAX  = 510    # Cargo M.N.:   365 ≤ x0 < 510  (montos ~384–392)
_X_ABONO_MIN  = 510    # Abono M.N.:   x0 ≥ 510        (montos ~526+)

# Palabras que identifican filas de encabezado de columna (con y sin tilde)
_COL_HEADER_WORDS = {
    "FECHA", "DESCRIPCION", "DESCRIPCIÓN", "TARJETA", "CARGO",
    "ABONO", "DIVISA", "M.N.", "M.N", "MN", "HISTORIAL", "MOVIMIENTOS", "TITULAR",
}

# Palabras de encabezado/pie que nunca son movimientos (para _is_skip_row)
_SKIP_WORDS = {
    "FECHA", "DESCRIPCION", "DESCRIPCIÓN", "TARJETA", "CARGO",
    "ABONO", "DIVISA", "SUBTOTAL", "HISTORIAL", "MOVIMIENTOS",
    "DEL", "TITULAR", "FISICA", "DIGITAL", "M.N.", "MN",
    "ADICIONALES", "MESES", "SALDO", "KONFIO", "KONFÍO",
    "STERIMEX", "RED", "AMIGO",
}


# ── Helpers ───────────────────────────────────────────────────────────────────

def _dedup_chars(text: str) -> str:
    """
    Corrige duplicación de fuente Konfío.
    Solo actúa cuando TODOS los caracteres vienen en pares idénticos consecutivos:
      'SSTTEERRIIMMEEXX' → 'STERIMEX'   (totalmente duplicado → toma posiciones pares)
      'OFFICE'          → 'OFFICE'      (primer par O≠F → no toca)
      'JEFF'            → 'JEFF'        (primer par J≠E → no toca)
      '2669'            → '2669'        (primer par 2≠6 → no toca)
    """
    n = len(text)
    if n < 2 or n % 2 != 0:
        return text
    for i in range(0, n, 2):
        if text[i] != text[i + 1]:
            return text
    return text[::2]


def _parse_date(words_fecha: list[str]) -> Optional[str]:
    """Convierte lista de palabras ["21", "ABR", "2026"] → "21/04/2026"."""
    texto = " ".join(words_fecha).strip().upper()
    m = _FECHA_RE.match(texto)
    if not m:
        return None
    dia  = m.group(1).zfill(2)
    mes  = MESES_ABR.get(m.group(2).upper(), "01")
    anio = m.group(3)
    return f"{dia}/{mes}/{anio}"


def _parse_amount(text: str) -> Optional[float]:
    m = _AMOUNT_RE.match(text.strip())
    if not m:
        return None
    try:
        return float(m.group(1).replace(",", ""))
    except ValueError:
        return None


def _normalizar_rfc(texto: str) -> str:
    """Junta el RFC partido: 'ESP 060105LJ5' → 'ESP060105LJ5'."""
    return _RFC_ESPACIO_RE.sub(lambda m: m.group(1) + m.group(2), texto)


def _group_rows(words: list[dict], tol: int = 3) -> dict[int, list[dict]]:
    """Agrupa palabras por fila (tolerancia ±tol pt en y)."""
    rows: dict[int, list[dict]] = {}
    for w in words:
        y = round(w["top"] / tol) * tol
        rows.setdefault(y, []).append(w)
    return rows


def _is_col_header_row(ws: list[dict]) -> bool:
    """
    True si la fila parece ser un encabezado de columna.
    Requiere al menos 2 palabras de encabezado para evitar falsos positivos
    con descripciones de comercios que contengan una sola palabra clave.
    """
    texts = {_dedup_chars(w["text"]).upper().strip(".,;:") for w in ws}
    return len(texts & _COL_HEADER_WORDS) >= 2


def _is_skip_row(row_words: list[dict]) -> bool:
    texts = {w["text"].upper().strip(".,;:") for w in row_words}
    return bool(texts & _SKIP_WORDS) and not any(
        _FECHA_RE.match(w["text"]) for w in row_words
    )


# ── Extractor ─────────────────────────────────────────────────────────────────

class KonfioExtractor(BancoExtractor):
    """
    Extractor para estados de cuenta Konfío Tarjeta de Crédito.
    Detecta el banco por las palabras "KONFIO" o "RED AMIGO DAL" en el PDF.
    """

    NOMBRE    = "Konfio"
    COLOR_HEX = "6B2D8B"   # morado corporativo Konfío

    @classmethod
    def detectar(cls, texto: str) -> bool:
        t = texto.upper()
        return "KONFIO" in t or "RED AMIGO DAL" in t

    def extraer(
        self,
        pdf_path: str,
        progress_callback: Optional[Callable[[int, int], None]] = None,
    ) -> tuple[list[Movimiento], InfoCuenta]:

        movimientos: list[Movimiento] = []
        info = InfoCuenta(banco="Konfio")
        current: Optional[Movimiento] = None
        in_section  = False
        pending_desc = ""   # descripción acumulada de fila A (sin fecha)
        pending_card = ""   # tarjeta de fila A

        with pdfplumber.open(pdf_path) as pdf:
            total_pages = len(pdf.pages)

            for pnum, page in enumerate(pdf.pages):
                if progress_callback:
                    progress_callback(pnum + 1, total_pages)

                if pnum == 0:
                    # El encabezado usa fuente estilizada → cada carácter duplicado.
                    # Deduplicar token por token preservando saltos de línea.
                    text0_raw = page.extract_text(x_tolerance=3, y_tolerance=3) or ""
                    text0 = "\n".join(
                        " ".join(_dedup_chars(tok) for tok in line.split())
                        for line in text0_raw.splitlines()
                    )
                    self._parse_header(text0, info)

                words = page.extract_words(x_tolerance=3, y_tolerance=3)
                rows  = _group_rows(words, tol=4)

                for y in sorted(rows.keys()):
                    ws = sorted(rows[y], key=lambda w: w["x0"])

                    # Texto normalizado (sin duplicados) sólo para detección de sección
                    row_text_norm = " ".join(_dedup_chars(w["text"]) for w in ws).upper()

                    # ── Inicio de sección de movimientos ─────────────────────
                    if "HISTORIAL DE MOVIMIENTOS" in row_text_norm:
                        in_section   = True
                        pending_desc = ""
                        pending_card = ""
                        continue

                    # ── Fin de sección (Subtotal) ─────────────────────────────
                    if in_section and row_text_norm.startswith("SUBTOTAL"):
                        if current:
                            movimientos.append(current)
                            current = None
                        in_section   = False
                        pending_desc = ""
                        pending_card = ""
                        # Los montos del subtotal pueden venir en fuente duplicada
                        amounts: list[float] = []
                        for w in ws:
                            a = _parse_amount(_dedup_chars(w["text"]))
                            if a is not None:
                                amounts.append(a)
                        if len(amounts) >= 2:
                            info.total_retiros   = amounts[0]
                            info.total_depositos = amounts[1]
                        elif len(amounts) == 1:
                            info.total_retiros = amounts[0]
                        continue

                    if not in_section:
                        continue

                    # Pie de página (parte baja de carta ≈792 pt)
                    if y > 700:
                        continue

                    # ── Saltar encabezados de columna ─────────────────────────
                    if _is_col_header_row(ws):
                        continue

                    # ── Clasificar palabras por columna ───────────────────────
                    fecha_ws    = [w for w in ws if w["x0"] < _X_FECHA_MAX]
                    desc_ws     = [w for w in ws if _X_FECHA_MAX <= w["x0"] < _X_DESC_MAX]
                    card_col_ws = [w for w in ws if _X_DESC_MAX   <= w["x0"] < _X_CARD_MAX]
                    cargo_ws    = [
                        w for w in ws
                        if _X_CARD_MAX <= w["x0"] < _X_CARGO_MAX
                        and _parse_amount(w["text"]) is not None
                    ]
                    abono_ws    = [
                        w for w in ws
                        if w["x0"] >= _X_ABONO_MIN
                        and _parse_amount(w["text"]) is not None
                    ]

                    # Número de tarjeta (4 dígitos en columna tarjeta)
                    card_num_ws = [
                        w for w in card_col_ws
                        if w["text"].isdigit() and len(w["text"]) == 4
                    ]

                    desc_texto  = " ".join(w["text"] for w in desc_ws).strip()
                    fecha_texto = " ".join(w["text"] for w in fecha_ws).strip()
                    fecha = _parse_date(fecha_texto.split()) if fecha_texto else None

                    if fecha:
                        # ── Fila B: fecha + monto → crear movimiento ──────────
                        if current:
                            movimientos.append(current)

                        retiro   = _parse_amount(cargo_ws[0]["text"]) if cargo_ws else None
                        deposito = _parse_amount(abono_ws[0]["text"]) if abono_ws else None

                        # Combinar descripción de fila A con cualquier texto en fila B
                        full_desc = pending_desc
                        if desc_texto:
                            full_desc = (full_desc + " " + desc_texto).strip()

                        # Si fila A y fila B cayeron en el mismo grupo de y,
                        # pending_card estará vacío → usar card_num_ws de esta fila.
                        folio = pending_card or (card_num_ws[0]["text"] if card_num_ws else "")

                        current = Movimiento(
                            fecha       = fecha,
                            folio       = folio,
                            descripcion = full_desc,
                            retiro      = retiro,
                            deposito    = deposito,
                            saldo       = None,
                        )
                        pending_desc = ""
                        pending_card = ""

                    elif current:
                        if card_num_ws:
                            # ── Fila A del SIGUIENTE movimiento ──────────────
                            # Un número de tarjeta de 4 dígitos indica que esta
                            # fila pertenece al próximo movimiento, no al actual.
                            pending_desc = desc_texto
                            pending_card = card_num_ws[0]["text"]
                        else:
                            # ── Fila C: RFC / tipo tarjeta (continuación) ────
                            if desc_texto:
                                skip_terms = {"FÍSICA", "FISICA", "DIGITAL"}
                                if desc_texto.upper() not in skip_terms:
                                    desc_norm = _normalizar_rfc(desc_texto)
                                    current.descripcion = (
                                        current.descripcion + " | " + desc_norm
                                    ).strip(" | ")

                            if not current.retiro and cargo_ws:
                                current.retiro = _parse_amount(cargo_ws[0]["text"])
                            if not current.deposito and abono_ws:
                                current.deposito = _parse_amount(abono_ws[0]["text"])

                    else:
                        # ── Fila A: descripción + tarjeta (aún sin fecha) ─────
                        if desc_texto:
                            pending_desc = (pending_desc + " " + desc_texto).strip()
                        if card_num_ws:
                            pending_card = card_num_ws[0]["text"]

        # Agregar último movimiento pendiente
        if current:
            movimientos.append(current)

        # Calcular totales si el Subtotal no los dio
        if info.total_retiros is None:
            info.total_retiros = sum(m.retiro or 0 for m in movimientos)
        if info.total_depositos is None:
            info.total_depositos = sum(m.deposito or 0 for m in movimientos)

        return movimientos, info

    # ── Parseo de encabezado ──────────────────────────────────────────────────

    def _parse_header(self, text: str, info: InfoCuenta) -> None:
        """Extrae RFC, No. de tarjeta, período y titular del encabezado.
        El texto ya llega deduplicado (token por token)."""
        lines = text.splitlines()
        for i, line in enumerate(lines):
            lu = line.upper().strip()

            # Titular: primera línea con razón social (truncar antes de "Periodo")
            if i < 5 and ("SA DE CV" in lu or "S.A. DE C.V" in lu or "S DE RL" in lu):
                if not info.titular:
                    # Quitar info de período que puede aparecer en la misma línea
                    info.titular = re.split(r"\s+[Pp]eriodo\b", line)[0].strip()

            # RFC
            if "RFC" in lu:
                m = re.search(r"\bRFC\s+([A-Z0-9]{10,13})\b", lu)
                if m:
                    info.rfc = m.group(1)

            # No. de tarjeta (16 dígitos en grupos de 4)
            m = re.search(r"(\d{4}\s+\d{4}\s+\d{4}\s+\d{4})", line)
            if m:
                info.numero_cuenta = m.group(1).replace(" ", "")

            # Período
            m = re.search(
                r"[Pp]eriodo\s+del\s+(.+?al\s+\d+\s+\w+\s+\d{4})", line, re.IGNORECASE
            )
            if m:
                info.periodo = m.group(1).strip()

        if not info.titular:
            for line in lines[:4]:
                if re.search(r"[A-ZÁÉÍÓÚ]{4,}", line) and "KONFIO" not in line.upper():
                    info.titular = line.strip()
                    break
