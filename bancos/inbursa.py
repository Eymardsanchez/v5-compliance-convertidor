"""
Extractor para estados de cuenta Banco Inbursa México.
Formato: MMM. DD  REFERENCIA  CONCEPTO  MONTO  SALDO
Líneas con glifos (cid:N) que se eliminan antes de parsear.
Hereda de BancoEstandar (bancos/estandar.py).
"""

import re
from typing import Optional

from bancos.estandar import BancoEstandar
from bancos.base import Movimiento, InfoCuenta

CID_RE    = re.compile(r"\(cid:\d+\)")
AMOUNT_RE = re.compile(r"([\d,]+\.\d{2})")

MESES = {
    "ENE": "01", "FEB": "02", "MAR": "03", "ABR": "04", "MAY": "05",
    "JUN": "06", "JUL": "07", "AGO": "08", "SEP": "09", "OCT": "10",
    "NOV": "11", "DIC": "12",
}

BALANCE_INICIAL_RE = re.compile(
    r"BALANCE INICIAL\s+([\d,]+\.\d{2})"
)

SKIP_PATTERNS = [
    "ESTADO DE CUENTA", "CLIENTE INBURSA", "PAGINA:", "RESUMEN DE SALDOS",
    "SALDO ACTUAL", "FECHA DE CORTE", "TASA BRUTA", "TASA NETA",
    "RENDIMIENTOS NETOS", "SALDO PROMEDIO",
    "CONSULTAS Y RECLAMACIONES", "PARA CUALQUIER",
]

DEPOSIT_KW = {"DEPOSITO", "ABONO", "RENDIMIENTO", "DEVOLUCION", "INTERES"}


class InbursaExtractor(BancoEstandar):
    NOMBRE        = "Inbursa"
    COLOR_HEX     = "E31837"
    DATE_RE       = re.compile(
        r"^(ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC)\.\s+(\d{1,2})\s+"
    )
    # Activa sección cuando aparece la línea de columnas FECHA REFERENCIA CONCEPTO ...
    SECTION_RE    = re.compile(
        r"FECHA\s+REFERENCIA\s+CONCEPTO\s+CARGOS\s+ABONOS\s+SALDO"
    )
    SKIP_PATTERNS = SKIP_PATTERNS
    END_PATTERNS  = [
        "RESUMEN DE SALDOS",      # sección de resumen al final del estado Inbursa
        "SALDO ACTUAL",           # aparece en el resumen final
    ]

    # _anio se fija en _parse_encabezado (página 0) y se usa en _parse_fecha
    _anio: str = ""

    @classmethod
    def detectar(cls, texto: str) -> bool:
        return "INBURSA" in texto and "CLIENTE INBURSA" in texto

    # ── Hooks ────────────────────────────────────────────────────────────────

    def _preprocesar(self, linea: str) -> str:
        """Elimina glifos (cid:N) que Inbursa incrusta en el PDF."""
        return CID_RE.sub("", linea).strip()

    def _parse_fecha(self, match: re.Match) -> str:
        mes_abr = match.group(1)
        dia     = match.group(2).zfill(2)
        anio    = getattr(self, "_anio", "")
        return f"{dia}/{mes_abr}./{anio}" if anio else f"{dia}/{mes_abr}."

    def _es_saldo_anterior(self, linea: str) -> Optional[float]:
        m = BALANCE_INICIAL_RE.search(linea.upper())
        if m:
            return float(m.group(1).replace(",", ""))
        return None

    def _clasificar_keywords(self, desc: str) -> str:
        upper = desc.upper()
        for kw in DEPOSIT_KW:
            if kw in upper:
                return "deposito"
        return "retiro"

    def _parse_encabezado(self, texto: str, info: InfoCuenta) -> None:
        # Extraer año del primer número de 4 dígitos que aparezca
        m = re.search(r"\b(20\d{2})\b", texto)
        self._anio = m.group(1) if m else ""

        for line in texto.splitlines():
            line = line.strip()
            if not info.periodo and "PERIODO" in line.upper():
                info.periodo = line
            if not info.clabe and "CLABE" in line.upper():
                m2 = re.search(r"\b(\d{18})\b", line)
                if m2:
                    info.clabe = m2.group(1)
            if not info.numero_cuenta and "CUENTA" in line.upper():
                m2 = re.search(r"\bCUENTA\b\s+(\d{8,})", line.upper())
                if m2:
                    info.numero_cuenta = m2.group(1)
            if not info.rfc and "RFC:" in line.upper():
                m2 = re.search(r"RFC:\s*([A-Z0-9]{10,13})", line.upper())
                if m2:
                    info.rfc = m2.group(1)
