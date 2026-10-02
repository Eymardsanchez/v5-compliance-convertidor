"""
Extractor para estados de cuenta Banorte (Enlace Negocios Basica y similares).
Formato: DD-MES-AA  DESCRIPCION  [DEPOSITO | RETIRO]  SALDO
Año de 2 dígitos (26 → 2026). Sin columna de folio separada.
Hereda de BancoEstandar (bancos/estandar.py).

CORRECCIONES v2 (2026-06-27)
─────────────────────────────
BUG 1 — Saldo con punto como separador de miles (OCR Banorte)
  El PDF Banorte en algunas líneas cortas (ej. "FERRE PINTMTZ") el OCR extrae el
  saldo como "296.076.78" en lugar de "296,076.78".  El AMOUNT_RE original capturaba
  "296.07" y "6.78" como dos montos distintos, dejando RETIRO/DEPOSITO y SALDO
  incorrectos.
  FIX: _corregir_saldo_punto() en _preprocesar() normaliza NNN.NNN.NN → NNN,NNN.NN
  antes de cualquier otro parseo.

BUG 2 — Montos atrapados en líneas de continuación
  Algunos movimientos SPEI multilínea tienen el monto y el saldo en la última línea
  de continuación (la primera línea -la de fecha- no contiene números).  _parse_linea
  devuelve un Movimiento con retiro=None, deposito=None, saldo=None; la línea de
  continuación se concatena como texto a descripcion pero jamás se re-parsea.
  FIX: _post_movimiento() detecta mov con saldo=None y retiro=None+deposito=None,
  rescata los dos últimos amounts de la descripción acumulada y los aplica.

BUG 3 — Falsa detección de continuación cuando la línea YA tiene montos
  (consecuencia del BUG 1 corregido en BUG 1) — con el preprocesado correcto, las
  líneas cortas como FERRE ahora producen exactamente 2 amounts y se clasifican
  correctamente.
"""

import re
import unicodedata
from typing import Optional

from bancos.estandar import BancoEstandar
from bancos.base import Movimiento, InfoCuenta


def _sin_acentos(s: str) -> str:
    """Elimina diacríticos para comparación de SKIP_PATTERNS."""
    return "".join(
        c for c in unicodedata.normalize("NFD", s)
        if unicodedata.category(c) != "Mn"
    )


AMOUNT_RE = re.compile(r"([\d,]+\.\d{2})")

# ── BUG 1 FIX ────────────────────────────────────────────────────────────────
# OCR Banorte a veces usa punto como separador de miles en saldos altos.
# Patrón afectado: NNN.NNN.NN  (p. ej. 296.076.78 → debe ser 296,076.78)
# Se reemplaza SOLO cuando el patrón es exactamente d{1,3}.d{3}.d{2} al final
# o seguido de espacio/fin, para no afectar horas tipo "11:36:21" ni RFCs.
_SALDO_PUNTO_MILES = re.compile(r"(\d{1,3})\.(\d{3})\.(\d{2})(?=\s|$)")


def _corregir_saldo_punto(linea: str) -> str:
    """Normaliza saldos con punto como separador de miles → coma."""
    return _SALDO_PUNTO_MILES.sub(
        lambda m: f"{m.group(1)},{m.group(2)}.{m.group(3)}", linea
    )


MESES = {
    "ENE": "01", "FEB": "02", "MAR": "03", "ABR": "04", "MAY": "05",
    "JUN": "06", "JUL": "07", "AGO": "08", "SEP": "09", "OCT": "10",
    "NOV": "11", "DIC": "12",
}

# ── OCR FIX (escaneos) ─────────────────────────────────────────────────────────
# El OCR confunde 0↔O y 1↔l/I en la fecha inicial (ej. "O3-AGO-26", "03-AG0-26"),
# lo que hacía que DATE_RE no reconociera la línea y el movimiento se perdiera.
# El día/año pueden traer un espacio OCR interno ("1 l-AGO-26" = 11-AGO-26).
_FECHA_OCR = re.compile(r"^([0-9OoIl]\s?[0-9OoIl])-([A-Za-z0-9]{3})-([0-9OoIl]\s?[0-9OoIl])")
_A_DIGITO  = str.maketrans({"O": "0", "o": "0", "I": "1", "l": "1"})
_A_LETRA   = str.maketrans({"0": "O", "1": "I"})


def _corregir_fecha_ocr(linea: str) -> str:
    """Normaliza la fecha DD-MES-AA inicial cuando el OCR metió letras por dígitos
    (o viceversa en el mes) o un espacio dentro del día/año. Solo actúa si el mes
    resultante es válido, para no tocar líneas que no son de movimiento."""
    m = _FECHA_OCR.match(linea)
    if not m:
        return linea
    dia  = m.group(1).replace(" ", "").translate(_A_DIGITO)
    mes  = m.group(2).upper().translate(_A_LETRA)
    anio = m.group(3).replace(" ", "").translate(_A_DIGITO)
    if mes not in MESES or not dia.isdigit() or not anio.isdigit():
        return linea
    return f"{dia}-{mes}-{anio}{linea[m.end():]}"

SKIP_PATTERNS = [
    "ESTADO DE CUENTA", "FECHA DESCRIPCION",
    "MONTO DEL DEPOSITO", "SIN MOVIMIENTOS", "COMPROBANTE FISCAL",
    "FOLIO FISCAL", "CADENA ORIGINAL", "SELLO DIGITAL",
    "RESUMEN INTEGRAL", "RESUMEN DEL PERIODO", "RESUMEN DE COMISIONES",
    "GANANCIA ANUAL", "OTROS CARGOS",
    # Pies de página (aparecen en TODAS las páginas — solo filtrar, no detener)
    "BANCO MERCANTIL DEL NORTE S.A.",
    "CONSULTAS, RECLAMACIONES",
    "LINEA DIRECTA",
    "CIUDAD DE MEXICO",
    "NUEVO LEON",
    # Encabezados de página repetidos
    "DETALLE DE MOVIMIENTOS",
    "ENLACE NEGOCIOS BASICA",
]

END_PATTERNS = [
    "MENSAJES IMPORTANTES",
    "AVISO DE PRIVACIDAD",
    "COMPROBANTE FISCAL DIGITAL",
    "INVERSION ENLACE NEGOCIOS",
]


def _expand_date(date_str: str) -> str:
    """Convierte DD-MES-YY a DD/MM/YYYY."""
    parts = date_str.split("-")
    if len(parts) == 3:
        day, mes, yr = parts
        mes_num = MESES.get(mes, mes)
        year = "20" + yr if len(yr) == 2 else yr
        return f"{day}/{mes_num}/{year}"
    return date_str


def _verify_by_balance(mov: Movimiento, prev_saldo: Optional[float]) -> None:
    """Corrige DEP/RET comparando el saldo anterior con el saldo actual."""
    if prev_saldo is None or mov.saldo is None:
        return
    amount = (mov.deposito or 0.0) + (mov.retiro or 0.0)
    if amount == 0:
        return
    if abs((prev_saldo + amount) - mov.saldo) < 1.0:
        mov.deposito = amount
        mov.retiro = None
    elif abs((prev_saldo - amount) - mov.saldo) < 1.0:
        mov.retiro = amount
        mov.deposito = None


class BanorteExtractor(BancoEstandar):
    NOMBRE        = "Banorte"
    COLOR_HEX     = "006341"
    DATE_RE       = re.compile(r"^(\d{2}-[A-Z]{3}-\d{2})")
    SECTION_RE    = re.compile(r"DETALLE DE MOVIMIENTOS")
    SKIP_PATTERNS = SKIP_PATTERNS
    END_PATTERNS  = END_PATTERNS

    @classmethod
    def detectar(cls, texto: str) -> bool:
        t = texto.upper()
        return (
            "BANCO MERCANTIL DEL NORTE" in t
            or "GRUPO FINANCIERO BANORTE" in t
            or "ENLACE NEGOCIOS" in t
            or "ENLACE GLOBAL" in t
        )

    # ── Hooks ────────────────────────────────────────────────────────────────

    def _preprocesar(self, linea: str) -> str:
        """
        BUG 1 FIX: normaliza saldos OCR con punto como separador de miles
        antes de cualquier parseo de importes.
        OCR FIX: corrige la fecha inicial (O↔0, l/I↔1) para que DATE_RE la reconozca.
        """
        return _corregir_saldo_punto(_corregir_fecha_ocr(linea))

    def _es_skip(self, linea: str) -> bool:
        """Normaliza acentos antes de comparar."""
        return super()._es_skip(_sin_acentos(linea))

    def _parse_fecha(self, match: re.Match) -> str:
        return _expand_date(match.group(1))

    def _es_saldo_anterior(self, linea: str) -> Optional[float]:
        """La línea de saldo inicial en Banorte empieza con fecha."""
        m = re.match(r"^(\d{2}-[A-Z]{3}-\d{2})", linea)
        if not m:
            return None
        if "SALDO ANTERIOR" not in linea.upper():
            return None
        amts = AMOUNT_RE.findall(linea)
        return float(amts[-1].replace(",", "")) if amts else 0.0

    def _parse_linea(self, match: re.Match, resto: str,
                     prev_saldo: Optional[float]) -> Optional[Movimiento]:
        """Banorte no tiene folio; la descripción ocupa todo el campo de texto."""
        fecha   = self._parse_fecha(match)
        amounts = AMOUNT_RE.findall(resto)

        saldo: Optional[float]      = None
        amount_val: Optional[float] = None
        desc = resto

        if len(amounts) >= 2:
            amount_val = float(amounts[-2].replace(",", ""))
            saldo      = float(amounts[-1].replace(",", ""))
            last_pos   = resto.rfind(amounts[-2])
            desc       = resto[:last_pos].strip()
        elif len(amounts) == 1:
            saldo    = float(amounts[0].replace(",", ""))
            last_pos = resto.rfind(amounts[0])
            desc     = resto[:last_pos].strip()

        desc = re.sub(r"\s+", " ", desc).strip()
        mov  = Movimiento(fecha=fecha, descripcion=desc, saldo=saldo)
        if amount_val is not None:
            # Asignamos como retiro; _post_movimiento lo corregirá si es depósito
            mov.retiro = amount_val
        return mov

    def _post_movimiento(self, mov: Movimiento,
                         prev_saldo: Optional[float]) -> None:
        """
        BUG 2 FIX: cuando _parse_linea no encontró montos en la línea de fecha
        (ocurre en SPEI multilínea donde el monto viene al final de la descripción
        acumulada), rescata los dos últimos amounts de la descripción completa.

        Condición de rescate: saldo=None AND retiro=None AND deposito=None
        y la descripción acumulada contiene al menos 2 importes.
        """
        if mov.saldo is None and mov.retiro is None and mov.deposito is None:
            # Intentar rescatar montos de la descripción acumulada (incluye continuaciones)
            desc_normalizada = _corregir_saldo_punto(mov.descripcion)
            amts = AMOUNT_RE.findall(desc_normalizada)
            if len(amts) >= 2:
                try:
                    amount_val = float(amts[-2].replace(",", ""))
                    saldo_val  = float(amts[-1].replace(",", ""))
                    mov.saldo  = saldo_val
                    # Clasificar por balance tracking
                    if prev_saldo is not None:
                        if abs((prev_saldo + amount_val) - saldo_val) < 1.0:
                            mov.deposito = amount_val
                        elif abs((prev_saldo - amount_val) - saldo_val) < 1.0:
                            mov.retiro = amount_val
                        else:
                            # No coincide con balance — asignar como retiro por defecto
                            mov.retiro = amount_val
                    else:
                        mov.retiro = amount_val
                    return
                except (ValueError, IndexError):
                    pass

        # Flujo normal: corregir DEP/RET por balance tracking
        _verify_by_balance(mov, prev_saldo)

        # Regla anti-traspaso: si el balance no pudo decidir (saldo ausente),
        # usar la COLUMNA (CARGOS/ABONOS por posición X). Solo actúa en ese hueco.
        if mov.saldo is None or prev_saldo is None:
            monto = (mov.retiro or 0.0) + (mov.deposito or 0.0)
            col = self._clasificar_columna(mov.fecha, monto) if monto else None
            if col == "deposito":
                mov.deposito, mov.retiro = monto, None
            elif col == "retiro":
                mov.retiro, mov.deposito = monto, None

    def _parse_encabezado(self, texto: str, info: InfoCuenta) -> None:
        for line in texto.splitlines():
            line = line.strip()
            if not info.rfc and "RFC" in line.upper():
                up = line.upper()
                # Evitar el RFC del propio banco (pie de página) y tomar el del titular.
                # El OCR escribe "RFC. PIIG7105099W2" (punto en vez de dos puntos).
                if "BANCO MERCANTIL" not in up and "BMN930209927" not in up:
                    m = re.search(r"RFC[.:]?\s*([A-ZÑ&]{3,4}\d{6}[A-Z0-9]{3})", up)
                    if m:
                        info.rfc = m.group(1)
            if not info.periodo and (
                    "PERIODO DEL" in line.upper() or "PERÍODO DEL" in line.upper()):
                m = re.search(r"[Pp]er[íi]odo\s+[Dd]el\b.+", line)
                info.periodo = m.group(0).strip() if m else line
            if not info.clabe:
                m = re.search(r"\b(\d{18})\b", line)
                if m:
                    info.clabe = m.group(1)
            if not info.numero_cuenta and "NO. DE CUENTA" in line.upper():
                m = re.search(r"NO\.\s*DE\s*CUENTA\s+(\d+)", line.upper())
                if m:
                    info.numero_cuenta = m.group(1)
