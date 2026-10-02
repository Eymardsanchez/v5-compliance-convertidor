"""
Extractor para estados de cuenta Banco Santander México.
Maneja la doble capa de texto (fuente visual duplicada) característica de Santander.
Hereda de BancoEstandar (bancos/estandar.py).
"""

import re
from typing import Optional

from bancos.estandar import BancoEstandar
from bancos.base import Movimiento, InfoCuenta

AMOUNT_RE = re.compile(r"([\d,]+\.\d{2})")

DEPOSIT_KW = {"ABONO", "DEPOSITO", "INTERES", "DEVOLUCION", "NOMINA"}

SKIP_PATTERNS = [
    "BANCO SANTANDER", "ESTADO DE CUENTA", "CORPORATIVO EN ADMINISTRACION",
    "PERIODO DEL", "PAGINA", "PAG.", "SALDO FINAL DEL", "RESUMEN", "SUCURSAL",
]


def _is_doubled_token(token: str) -> bool:
    if len(token) < 2 or len(token) % 2 != 0:
        return False
    return all(token[i] == token[i + 1] for i in range(0, len(token) - 1, 2))


def _undouble_line(line: str) -> str:
    """Decodifica líneas con fuente visual duplicada: 'PPAAGGOO' → 'PAGO'."""
    return " ".join(
        tok[::2] if _is_doubled_token(tok) else tok
        for tok in line.split(" ")
    )


def _verify_by_balance(mov: Movimiento, prev_saldo: Optional[float]) -> None:
    """Corrige la clasificación DEP/RET comparando saldo anterior con saldo actual."""
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


class SantanderExtractor(BancoEstandar):
    NOMBRE        = "Santander"
    COLOR_HEX     = "CC0000"
    DATE_RE       = re.compile(r"^\d{2}-[A-Z]{3}-\d{4}\s")
    SECTION_RE    = None          # sección activada por _es_saldo_anterior
    SKIP_PATTERNS = SKIP_PATTERNS
    # "SALDO FINAL DEL PERIODO" (sin ANTERIOR) marca el fin de la sección
    END_PATTERNS  = ["SALDO FINAL DEL PERIODO"]

    @classmethod
    def detectar(cls, texto: str) -> bool:
        t = texto.upper()
        return "BANCO SANTANDER" in t or "SANTANDER MEXICO" in t

    # ── Hooks ────────────────────────────────────────────────────────────────

    def _es_skip(self, linea: str) -> bool:
        if super()._es_skip(linea):
            return True
        u = linea.upper()
        # Filtrar la fila de totales al final del estado (empieza con TOTAL + importes)
        if u.startswith("TOTAL") and AMOUNT_RE.search(linea):
            return True
        return False

    def _preprocesar(self, linea: str) -> str:
        return _undouble_line(linea)

    def _es_saldo_anterior(self, linea: str) -> Optional[float]:
        if "SALDO FINAL DEL PERIODO ANTERIOR" in linea.upper():
            amts = AMOUNT_RE.findall(linea)
            return float(amts[-1].replace(",", "")) if amts else 0.0
        return None

    def _clasificar_keywords(self, desc: str) -> str:
        upper = desc.upper()
        for kw in DEPOSIT_KW:
            if kw in upper:
                return "deposito"
        return "retiro"

    def _post_movimiento(self, mov: Movimiento,
                         prev_saldo: Optional[float]) -> None:
        _verify_by_balance(mov, prev_saldo)

    def _parse_encabezado(self, texto: str, info: InfoCuenta) -> None:
        for line in texto.splitlines():
            line = line.strip()
            if not info.periodo and "PERIODO DEL" in line.upper():
                info.periodo = line
            if not info.rfc:
                m = re.search(r"R\.?F\.?C\.?\s*([A-Z0-9]{10,13})", line.upper())
                if m:
                    info.rfc = m.group(1)
            if not info.numero_cuenta:
                m = re.search(r"\d{2}-\d{10}-\d", line)
                if m:
                    info.numero_cuenta = m.group()
            if not info.clabe:
                m = re.search(r"\b(\d{18})\b", line)
                if m:
                    info.clabe = m.group(1)
