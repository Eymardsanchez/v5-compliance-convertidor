"""
bancos/estandar.py
──────────────────
Clase base unificada para extractores con formato tabular estándar.

PATRÓN COMÚN (Santander · BBVA · Banorte · Inbursa):
  [FECHA]  [REF/FOLIO?]  [DESCRIPCION ...]  [MONTO?]  [SALDO]

Reglas generales:
  • La línea de transacción SIEMPRE empieza con una fecha (formato varía por banco).
  • Los últimos 1-2 números de esa línea son el importe y el saldo.
  • Las líneas siguientes que NO empiezan con fecha son continuación del concepto.
  • Balance tracking: prev_saldo ± monto ≈ saldo_nuevo  →  define CARGO o ABONO.

Las subclases solo necesitan declarar los atributos de clase y sobreescribir
los métodos que difieren del comportamiento por defecto.

Banregio usa coordenadas X (extract_words) y tiene su propio extractor.
"""

import re
from collections import defaultdict
from typing import Optional, Callable

import pdfplumber

from bancos.base import BancoExtractor, Movimiento, InfoCuenta

AMOUNT_RE = re.compile(r"([\d,]+\.\d{2})")

# ── Clasificación por COLUMNA (CARGOS/ABONOS) — regla anti-traspaso ───────────
# La señal contable REAL de si un importe es cargo o abono es la COLUMNA (la
# posición X) donde el banco lo imprime. El balance-tracking (prev ± monto ≈
# saldo) es la señal primaria, pero cuando el banco NO imprime el saldo corrido
# en esa línea, no puede decidir y el código caía a keyword→"retiro" por
# defecto. Eso leía un TRASPASO entre cuentas propias como retiro en las DOS
# cuentas. Aquí se lee la columna y se usa SOLO cuando el balance no pudo
# decidir → cero regresión (el balance sigue mandando donde funciona).
_MESES_KEY = {
    "ENE": "01", "FEB": "02", "MAR": "03", "ABR": "04", "MAY": "05", "JUN": "06",
    "JUL": "07", "AGO": "08", "SEP": "09", "OCT": "10", "NOV": "11", "DIC": "12",
}
CARGO_HEADERS = frozenset({"CARGOS", "CARGO", "RETIRO", "RETIROS", "DEBE"})
ABONO_HEADERS = frozenset({"ABONOS", "ABONO", "DEPOSITO", "DEPOSITOS",
                           "DEPÓSITO", "DEPÓSITOS", "HABER"})
SALDO_HEADERS = frozenset({"SALDO", "SALDOS", "OPERACION", "OPERACIÓN",
                           "LIQUIDACION", "LIQUIDACIÓN"})
_MONTO_RE_FULL = re.compile(r"\d[\d,]*\.\d{2}")


def _norm_hdr(t: str) -> str:
    return re.sub(r"[^A-ZÑÁÉÍÓÚ]", "", (t or "").upper())


def fecha_key(fecha: str) -> str:
    """Normaliza una fecha ('DD/MM', 'DD/MMM', 'DD-MES-AA', 'D MES',
    'DD/MM/YYYY') a una clave estable 'DDMM' (día+mes en dígitos). '' si no."""
    if not fecha:
        return ""
    m = re.match(r"\s*(\d{1,2})[/\-.\s]+([A-Za-zÑ0-9]{2,})", str(fecha))
    if not m:
        return ""
    dia = m.group(1).zfill(2)
    mes = m.group(2).upper()[:3]
    mes = _MESES_KEY.get(mes, mes)
    mes = re.sub(r"\D", "", mes)
    if not mes:
        return ""
    return dia + mes.zfill(2)


def _detectar_centros_columnas(words):
    """(cx_cargo, cx_abono, [cx_saldo...]) de los encabezados en la página, o
    None si no se detectan cargo Y abono Y al menos un saldo (exigir el saldo
    evita confundir la columna de saldo con la de abono)."""
    cx_cargo = cx_abono = None
    saldos = []
    for w in words:
        h = _norm_hdr(w.get("text", ""))
        if not h:
            continue
        cx = (w["x0"] + w["x1"]) / 2.0
        if h in CARGO_HEADERS and cx_cargo is None:
            cx_cargo = cx
        elif h in ABONO_HEADERS and cx_abono is None:
            cx_abono = cx
        elif h in SALDO_HEADERS:
            saldos.append(cx)
    if cx_cargo is None or cx_abono is None or not saldos:
        return None
    return cx_cargo, cx_abono, saldos


def _agrupar_renglones(words, tol: float = 3.0):
    """Agrupa words por coordenada 'top' (mismo renglón visual)."""
    filas: list[tuple[float, list]] = []
    for w in sorted(words, key=lambda w: (round(w["top"] / tol), w["x0"])):
        if filas and abs(w["top"] - filas[-1][0]) <= tol:
            filas[-1][1].append(w)
        else:
            filas.append((w["top"], [w]))
    return [f[1] for f in filas]


def construir_mapa_columnas(pdf_path: str) -> dict:
    """Mapa {(fecha_key, centavos) -> 'retiro'|'deposito'} leyendo la posición X
    del importe respecto a los encabezados CARGOS/ABONOS/SALDO. Descarta claves
    ambiguas (mismo importe+fecha en ambas columnas) → sin falso positivo.
    Reutilizable por cualquier banco. Nunca lanza: ante cualquier problema
    devuelve {} (comportamiento histórico = sin cambio)."""
    acc: dict = defaultdict(set)
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                words = page.extract_words(x_tolerance=1.5, y_tolerance=3,
                                           use_text_flow=False)
                if not words:
                    continue
                centros = _detectar_centros_columnas(words)
                if not centros:
                    continue
                cx_cargo, cx_abono, saldos = centros
                candidatos = ([("retiro", cx_cargo), ("deposito", cx_abono)]
                              + [("saldo", s) for s in saldos])
                for row in _agrupar_renglones(words):
                    fk = fecha_key(" ".join(w.get("text", "") for w in row[:3]))
                    if not fk:
                        continue
                    for w in row:
                        t = (w.get("text") or "").replace(",", "")
                        if not _MONTO_RE_FULL.fullmatch(t):
                            continue
                        try:
                            cents = int(round(float(t) * 100))
                        except ValueError:
                            continue
                        cx = (w["x0"] + w["x1"]) / 2.0
                        lado = min(candidatos, key=lambda c: abs(cx - c[1]))[0]
                        if lado in ("retiro", "deposito"):
                            acc[(fk, cents)].add(lado)
    except Exception:
        return {}
    return {k: next(iter(v)) for k, v in acc.items() if len(v) == 1}


class BancoEstandar(BancoExtractor):
    """
    Extractor base reutilizable para bancos con formato tabular estándar.

    Atributos que las subclases DEBEN definir
    ──────────────────────────────────────────
    NOMBRE      : str            nombre del banco (para InfoCuenta y Excel)
    COLOR_HEX   : str            color del encabezado en Excel  (ej. "CC0000")
    DATE_RE     : re.Pattern     regex que detecta el inicio de una línea de movimiento
    SECTION_RE  : re.Pattern|None regex que activa la sección de movimientos (None si se
                                 usa _es_saldo_anterior para activarla)

    Atributos opcionales con valor por defecto
    ──────────────────────────────────────────
    SKIP_PATTERNS : list[str]   fragmentos que, si aparecen en una línea, la descartan
    END_PATTERNS  : list[str]   fragmentos que marcan el FIN de la sección
    """

    NOMBRE:        str        = ""
    COLOR_HEX:     str        = "333333"
    DATE_RE:       re.Pattern = None   # type: ignore[assignment]
    SECTION_RE:    re.Pattern | None = None
    SKIP_PATTERNS: list[str]  = []
    END_PATTERNS:  list[str]  = []
    HEADER_PAGE:   int        = 0      # página donde está el encabezado (BBVA usa 1)
    COLUMNAS_AUTO: bool       = True   # auto-clasificar por columna cuando el balance no decide
    SECCION_LIBRE: bool       = False  # si True, no exige activar sección: toda línea
                                       # con fecha es movimiento (lo usa el genérico)

    # ─────────────────────────────────────────────────────────────────────────
    # Hooks de texto — sobreescribir si el banco tiene codificación especial
    # ─────────────────────────────────────────────────────────────────────────

    def _preprocesar(self, linea: str) -> str:
        """Limpieza / decodificación antes de analizar (ej: undouble Santander)."""
        return linea

    # ─────────────────────────────────────────────────────────────────────────
    # Hooks de parseo — sobreescribir para formatos de fecha / importe especiales
    # ─────────────────────────────────────────────────────────────────────────

    def _parse_fecha(self, match: re.Match) -> str:
        """Convierte el match de DATE_RE a un string de fecha normalizado."""
        return match.group(0).strip()

    def _extraer_importes(self, texto: str
                          ) -> tuple[Optional[float], Optional[float], str]:
        """
        Extrae (monto, saldo, descripcion_cruda) del texto que sigue a la fecha.

        Convención:
          · Si hay ≥2 números: penúltimo = monto, último = saldo.
          · Si hay 1 número: es el saldo (el monto estaba en otra columna o no existe).
          · Si no hay números: todo es descripción, saldo = None.
        """
        amounts = AMOUNT_RE.findall(texto)
        if len(amounts) >= 2:
            monto = float(amounts[-2].replace(",", ""))
            saldo = float(amounts[-1].replace(",", ""))
            pos   = texto.rfind(amounts[-2])
            desc  = texto[:pos].strip()
        elif len(amounts) == 1:
            monto = None
            saldo = float(amounts[-1].replace(",", ""))
            pos   = texto.rfind(amounts[-1])
            desc  = texto[:pos].strip()
        else:
            monto = None
            saldo = None
            desc  = texto.strip()
        return monto, saldo, desc

    def _parse_linea(self, match: re.Match, resto: str,
                     prev_saldo: Optional[float]) -> Optional[Movimiento]:
        """
        Construye un Movimiento a partir de la línea de transacción.
        Sobreescribir para formatos muy diferentes (ej: BBVA con 3 importes).
        """
        fecha  = self._parse_fecha(match)
        monto, saldo, desc_raw = self._extraer_importes(resto)
        desc_raw = re.sub(r"\s+", " ", desc_raw).strip()

        # Primer token = folio/referencia; el resto = descripción
        partes = desc_raw.split(None, 1)
        folio  = partes[0] if partes else ""
        desc   = partes[1].strip() if len(partes) > 1 else ""

        mov = Movimiento(fecha=fecha, folio=folio, descripcion=desc, saldo=saldo)

        if monto is not None:
            tipo = self._clasificar(monto, prev_saldo, saldo, desc_raw, fecha=fecha)
            if tipo == "deposito":
                mov.deposito = monto
            else:
                mov.retiro = monto

        return mov

    # ─────────────────────────────────────────────────────────────────────────
    # Clasificación — sobreescribir _clasificar_keywords si el banco
    # tiene palabras clave propias; _clasificar hace balance tracking primero.
    # ─────────────────────────────────────────────────────────────────────────

    def _clasificar(self, monto: float,
                    prev:  Optional[float],
                    saldo: Optional[float],
                    desc:  str = "",
                    fecha: str = "") -> str:
        """Balance → COLUMNA (si el balance no pudo) → keywords → 'retiro'.

        La columna es la corrección anti-traspaso: solo actúa cuando el balance
        no puede decidir (el banco no imprimió el saldo en esa línea). Nunca
        pisa lo que el balance ya resolvió → cero regresión."""
        b = self._clasificar_balance(monto, prev, saldo)
        if b is not None:
            return b
        col = self._clasificar_columna(fecha, monto)
        if col is not None:
            return col
        kw = self._clasificar_keywords(desc)
        if kw != "retiro":
            return kw
        from conciliacion import cargar_kw_usuario
        upper = desc.upper().replace(" ", "")
        for d in cargar_kw_usuario("depositos"):
            if d in upper:
                return "deposito"
        for r in cargar_kw_usuario("retiros"):
            if r in upper:
                return "retiro"
        return kw

    def _clasificar_balance(self, monto: float,
                            prev: Optional[float],
                            saldo: Optional[float]) -> Optional[str]:
        """'deposito'/'retiro' por balance tracking, o None si no puede decidir."""
        if prev is not None and saldo is not None and monto:
            if abs((prev + monto) - saldo) < 1.0:
                return "deposito"
            if abs((prev - monto) - saldo) < 1.0:
                return "retiro"
        return None

    def _clasificar_columna(self, fecha: str,
                            monto: Optional[float]) -> Optional[str]:
        """Consulta el mapa de columnas por (fecha, importe). None si no aplica."""
        mapa = getattr(self, "_col_map", None)
        if not mapa or monto is None or not fecha:
            return None
        return mapa.get((fecha_key(fecha), int(round(monto * 100))))

    def _clasificar_keywords(self, desc: str) -> str:
        """Clasificación por palabras clave. Subclases sobreescriben esto."""
        return "retiro"

    # ─────────────────────────────────────────────────────────────────────────
    # Hooks de sección
    # ─────────────────────────────────────────────────────────────────────────

    def _es_saldo_anterior(self, linea: str) -> Optional[float]:
        """
        Si la línea contiene el saldo inicial del período, retorna su valor
        y activa la sección de movimientos.
        Subclases sobreescriben esto.
        """
        return None

    def _es_fin_seccion(self, linea: str) -> bool:
        return any(p.upper() in linea.upper() for p in self.END_PATTERNS)

    def _es_skip(self, linea: str) -> bool:
        """True si la línea debe descartarse (encabezados, pies de página)."""
        u = linea.upper()
        return (any(p in u for p in self.SKIP_PATTERNS)
                or _es_encabezado_columnas(linea))

    def _post_movimiento(self, mov: Movimiento,
                         prev_saldo: Optional[float]) -> None:
        """
        Hook ejecutado justo antes de guardar un movimiento.
        Útil para reajustar la clasificación por balance (Santander, Banorte).
        """
        pass

    # ─────────────────────────────────────────────────────────────────────────
    # Encabezado — sobreescribir para extraer campos adicionales
    # ─────────────────────────────────────────────────────────────────────────

    def _parse_encabezado(self, texto: str, info: InfoCuenta) -> None:
        """Extrae RFC, CLABE, período del encabezado de la primera página."""
        for ln in texto.splitlines():
            ln = ln.strip()
            if not info.rfc:
                m = re.search(
                    r"R\.?F\.?C\.?\s*[:\-]?\s*([A-Z&]{3,4}\d{6}[A-Z0-9]{3})",
                    ln.upper())
                if m:
                    info.rfc = m.group(1)
            if not info.periodo and "PERIODO" in ln.upper():
                info.periodo = ln
            if not info.clabe:
                m = re.search(r"\b(\d{18})\b", ln)
                if m:
                    info.clabe = m.group(1)

    # ─────────────────────────────────────────────────────────────────────────
    # Extracción principal — común a todos los bancos estándar
    # ─────────────────────────────────────────────────────────────────────────

    def extraer(self, pdf_path: str,
                progress_callback: Optional[Callable[[int, int], None]] = None
                ) -> tuple[list[Movimiento], InfoCuenta]:

        movimientos: list[Movimiento]  = []
        current:     Optional[Movimiento] = None
        info         = InfoCuenta(banco=self.NOMBRE)
        in_section   = False
        prev_saldo:  Optional[float] = None
        done         = False

        # Mapa de columnas (CARGOS/ABONOS por posición X) — para clasificar los
        # movimientos donde el balance no puede decidir. try/except: si falla,
        # queda vacío y todo sigue como antes.
        self._col_map = {}
        if self.COLUMNAS_AUTO:
            try:
                self._col_map = construir_mapa_columnas(pdf_path)
            except Exception:
                self._col_map = {}

        with pdfplumber.open(pdf_path) as pdf:
            total = len(pdf.pages)

            for pnum, page in enumerate(pdf.pages):
                if done:
                    break
                if progress_callback:
                    progress_callback(pnum + 1, total)

                raw_text = page.extract_text(x_tolerance=2, y_tolerance=3) or ""
                if not raw_text:
                    continue

                if pnum == self.HEADER_PAGE:
                    self._parse_encabezado(raw_text, info)

                for raw in raw_text.splitlines():
                    linea = self._preprocesar(raw.strip())
                    if not linea:
                        continue

                    # ── Fin de sección ────────────────────────────────────────
                    if in_section and self._es_fin_seccion(linea):
                        if current:
                            self._post_movimiento(current, prev_saldo)
                            movimientos.append(current)
                            current = None
                        done = True
                        break

                    # ── Saldo inicial / activación de sección ─────────────────
                    si = self._es_saldo_anterior(linea)
                    if si is not None and info.saldo_inicial is None:
                        info.saldo_inicial = si
                        prev_saldo = si
                        in_section = True
                        continue

                    # ── Encabezado de sección (SECTION_RE) ───────────────────
                    if (not in_section and self.SECTION_RE
                            and self.SECTION_RE.search(linea)):
                        in_section = True
                        continue

                    # ── Filtros ───────────────────────────────────────────────
                    if self._es_skip(linea):
                        continue

                    if not in_section and not self.SECCION_LIBRE:
                        continue

                    # ── Línea de transacción ──────────────────────────────────
                    m = self.DATE_RE.match(linea)
                    if m:
                        if current:
                            self._post_movimiento(current, prev_saldo)
                            prev_saldo = current.saldo
                            movimientos.append(current)
                        resto   = linea[m.end():]
                        current = self._parse_linea(m, resto, prev_saldo)

                    # ── Línea de continuación (detalle SPEI, etc.) ────────────
                    elif current:
                        current.descripcion += " | " + linea

            if current:
                self._post_movimiento(current, prev_saldo)
                movimientos.append(current)

        if movimientos:
            info.total_depositos = sum(m.deposito or 0 for m in movimientos)
            info.total_retiros   = sum(m.retiro   or 0 for m in movimientos)
            info.saldo_final     = movimientos[-1].saldo

        return movimientos, info


# ─────────────────────────────────────────────────────────────────────────────
# Utilitario compartido
# ─────────────────────────────────────────────────────────────────────────────

def _es_encabezado_columnas(linea: str) -> bool:
    """
    Detecta la línea de encabezado de columnas de la tabla de movimientos.
    Se filtra si tiene simultáneamente palabras de FECHA/FOLIO Y DEPOSITO/CARGO.
    Esto evita filtrar líneas de detalle SPEI que solo contienen una de esas palabras.
    """
    u = linea.upper()
    tiene_fecha = "FECHA" in u or "FOLIO" in u
    tiene_importe = "DEPOSITO" in u or "RETIRO" in u or "ABONO" in u or "CARGO" in u
    return tiene_fecha and tiene_importe
