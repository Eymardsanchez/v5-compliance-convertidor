"""
Extractor para estados de cuenta BBVA México (Cuenta Maestra PYME y similares).
Formato: DD/MMM  DD/MMM  COD  DESCRIPCION  [CARGO|ABONO]  [SALDO_OPER]  [SALDO_LIQ]
La página 1 es imagen (sin texto). Encabezado en página 2 (HEADER_PAGE = 1).
Hereda de BancoEstandar (bancos/estandar.py).
"""

import re
from typing import Optional

from bancos.estandar import BancoEstandar
from bancos.base import Movimiento, InfoCuenta

AMOUNT_RE = re.compile(r"([\d,]+\.\d{2})")

MESES = {
    "ENE": "01", "FEB": "02", "MAR": "03", "ABR": "04", "MAY": "05",
    "JUN": "06", "JUL": "07", "AGO": "08", "SEP": "09", "OCT": "10",
    "NOV": "11", "DIC": "12",
}

# ABONO = depósito / crédito
ABONO_KW = {
    "VENTAS", "ABONO", "DEPOSITO", "BONIFICACION",
    "VENTA NAL", "VENTAS TDC", "VENTA INTER",
    "RECIBIDO",   # T20: SPEI RECIBIDO = transferencia entrante
    "DEVUELTO",   # T22: SPEI devuelto por banco destino
}
# CARGO = retiro / débito
CARGO_KW = {
    "ENVIADO", "PAGO", "APLI TASA", "IVA COM", "IVA TASA",
    "SERV BANCA", "ANUALIDAD", "COMISION", "CARGO", "RETIRO",
    "COBRANZA", "NETFLIX", "SPOTIFY", "AMAZON",
    "COM VTAS",   # I73: comisión sobre ventas TDC internacional
    "COM. VTA",   # K59/K65: comisión AMEX
}

SKIP_PATTERNS = [
    "ESTADO DE CUENTA", "MAESTRA PYME", "PAGINA", "PERIODO DEL",
    "RENDIMIENTO", "COMPORTAMIENTO", "INFORMACION FINANCIERA",
    "SALDO PROMEDIO", "DIAS DEL PERIODO", "TASA BRUTA",
    "INTERESES A FAVOR", "ISR RETENIDO", "COMISIONES",
    "DEPOSITOS / ABONOS", "RETIROS / CARGOS", "SALDO FINAL",
    "No. de Cuenta", "No. de Cliente", "R.F.C", "SUCURSAL",
    "DOMICILIO FISCAL", "PLAZA:", "TELEFONO", "ESTIMADO CLIENTE",
    "SU ESTADO DE CUENTA", "LA GAT REAL",
    "FECHA SALDO", "OPER LIQ COD",
]

SALDO_INI_RE = re.compile(
    r"Saldo de (?:Liquidaci[oó]n|Operaci[oó]n) Inicial\s+([\d,]+\.\d{2})"
)


class BBVAExtractor(BancoEstandar):
    NOMBRE        = "BBVA"
    COLOR_HEX     = "004481"
    # Sólo captura los tres tokens de encabezado; el resto queda en `resto`
    DATE_RE       = re.compile(
        r"^(\d{2}/[A-Z]{3})\s+(\d{2}/[A-Z]{3})\s+([A-Z0-9]{2,3})\s+"
    )
    SECTION_RE    = re.compile(r"OPER\s+LIQ\s+COD")
    SKIP_PATTERNS = SKIP_PATTERNS
    END_PATTERNS  = [
        "Total de Movimientos",   # encabezado de la sección de totales BBVA
        "TOTAL IMPORTE CARGOS",   # líneas de totales (por si aparecen antes)
    ]
    HEADER_PAGE   = 1   # página 1 tiene el encabezado; página 0 es imagen

    @classmethod
    def detectar(cls, texto: str) -> bool:
        header = texto[:800]
        return (
            "BBVA" in header
            and ("MAESTRA PYME" in header
                 or "GRUPO FINANCIERO BBVA" in header
                 or "ESTADO DE CUENTA" in header)
            and "INBURSA" not in header
            and "SANTANDER" not in header
        )

    # ── Hooks ────────────────────────────────────────────────────────────────

    def _parse_fecha(self, match: re.Match) -> str:
        """Convierte DD/MMM (fecha operación) a DD/MM."""
        parts = match.group(1).split("/")
        if len(parts) == 2:
            mes_num = MESES.get(parts[1], parts[1])
            return f"{parts[0]}/{mes_num}"
        return match.group(1)

    def _es_saldo_anterior(self, linea: str) -> Optional[float]:
        m = SALDO_INI_RE.search(linea)
        if m:
            return float(m.group(1).replace(",", ""))
        return None

    def _parse_linea(self, match: re.Match, resto: str,
                     prev_saldo: Optional[float]) -> Optional[Movimiento]:
        """BBVA puede tener 3 importes: monto + saldo_operación + saldo_liquidación."""
        fecha = self._parse_fecha(match)
        code  = match.group(3)

        amounts = AMOUNT_RE.findall(resto)

        if len(amounts) >= 3:
            amount_val = float(amounts[-3].replace(",", ""))
            saldo_oper = float(amounts[-2].replace(",", ""))
            last_pos   = resto.rfind(amounts[-3])
            desc_raw   = resto[:last_pos].strip()
            saldo_final = saldo_oper
        elif len(amounts) == 2:
            amount_val  = float(amounts[-2].replace(",", ""))
            saldo_final = float(amounts[-1].replace(",", ""))
            last_pos    = resto.rfind(amounts[-2])
            desc_raw    = resto[:last_pos].strip()
        elif len(amounts) == 1:
            amount_val  = float(amounts[0].replace(",", ""))
            saldo_final = None
            last_pos    = resto.rfind(amounts[0])
            desc_raw    = resto[:last_pos].strip()
        else:
            return Movimiento(fecha=fecha, folio=code, descripcion=resto)

        desc_raw = re.sub(r"\s+", " ", desc_raw).strip()
        kind = self._clasificar(amount_val, prev_saldo, saldo_final,
                                code + " " + desc_raw, fecha=fecha)

        mov = Movimiento(fecha=fecha, folio=code, descripcion=desc_raw,
                         saldo=saldo_final)
        if kind == "deposito":
            mov.deposito = amount_val
        else:
            mov.retiro = amount_val
        return mov

    def _clasificar_keywords(self, desc: str) -> str:
        upper = desc.upper()
        for kw in ABONO_KW:
            if kw in upper:
                return "deposito"
        for kw in CARGO_KW:
            if kw in upper:
                return "retiro"
        return "retiro"

    def _parse_encabezado(self, texto: str, info: InfoCuenta) -> None:
        for line in texto.splitlines():
            line = line.strip()
            if not info.periodo and "PERIODO DEL" in line.upper():
                info.periodo = line
            if not info.numero_cuenta and "No. de Cuenta" in line:
                m = re.search(r"No\.\s*de\s*Cuenta\s+(\d+)", line)
                if m:
                    info.numero_cuenta = m.group(1)
            if not info.rfc and "R.F.C" in line.upper():
                m = re.search(r"R\.F\.C\s+([A-Z0-9]{10,13})", line.upper())
                if m:
                    info.rfc = m.group(1)
            if not info.clabe and "CLABE" in line.upper():
                m = re.search(r"\b(\d{18})\b", line)
                if m:
                    info.clabe = m.group(1)
