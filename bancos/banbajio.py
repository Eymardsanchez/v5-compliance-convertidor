"""
Extractor para estados de cuenta Banbajío (Banco del Bajío S.A.).
Formato: D/DD MES [REF] DESCRIPCION  $ DEPOSITO/RETIRO  $ SALDO
Columnas DEPOSITO y RETIRO separadas; clasificación por balance tracking.
Soporta cuentas MN y USD.
"""

import re
from typing import Optional, Callable

import pdfplumber

from bancos.base import BancoExtractor, Movimiento, InfoCuenta
from bancos.estandar import construir_mapa_columnas, fecha_key


AMOUNT_RE = re.compile(r"([\d,]+\.\d{2})")
MESES = {
    "ENE": "01", "FEB": "02", "MAR": "03", "ABR": "04", "MAY": "05",
    "JUN": "06", "JUL": "07", "AGO": "08", "SEP": "09", "OCT": "10",
    "NOV": "11", "DIC": "12",
}

# Línea de transacción: comienza con 1-2 dígitos (día) + espacio + 3 letras (mes)
DATE_RE = re.compile(r"^(\d{1,2})\s+([A-Z]{3})\s*(.*)")

# Año del encabezado: "PERIODO: 1 DE ABRIL AL 30 DE ABRIL DE 2026"
PERIOD_RE = re.compile(r"DE\s+(\d{4})\s*$", re.IGNORECASE)

SKIP_PATTERNS = [
    "ESTADO DE CUENTA", "NUMERO DE CLIENTE", "R.F.C.",
    "NO. REF.", "FECHA DESCRIPCION", "DOCTO",
    "CONTINUA EN LA SIGUIENTE PAGINA", "PAGINA",
    "SALDO PROMEDIO", "DIAS DEL PERIODO", "GAT NOMINAL", "GAT REAL",
    "TASA ANUAL", "RENDIMIENTOS", "COMISIONES EFECTIVAMENTE",
    "IMPORTE TOTAL DE COMISIONES", "RESUMEN DE COMISIONES",
    "RESUMEN DE CARGOS", "NO EXISTEN CARGOS",
    "ESTE DOCUMENTO ES", "NUMERO DE CERTIFICADO", "SELLO DIGITAL",
    "CADENA ORIGINAL", "TIPO DE COMPROBANTE", "FOLIO FISCAL",
    "FORMA DE PAGO", "NOMBRE O RAZON SOCIAL",
    "SIGNIFICADO DE ABREVIATURAS", "ADVERTENCIAS Y RECOMENDACIONES",
    "RIESGOS", "EL IPAB PROTEGE", "GRACIAS POR HACER",
    "BANCO DEL BAJIO", "BANCO DELBAJIO",
    "(DATO NO VERIFICADO", "CLAVE DE RASTREO",
    "PRODUCTOS DE VISTA", "CLABE INTERBANCARIA",
    "SALDO ANTERIOR", "FECHA DE CORTE",
    "DEL PERIODO EN EL AÑO",
]


def _extract_year(text: str) -> str:
    """Extrae el año del encabezado del periodo."""
    m = PERIOD_RE.search(text)
    if m:
        return m.group(1)
    # Buscar cualquier año de 4 dígitos 20XX
    m2 = re.search(r"\b(20\d{2})\b", text)
    return m2.group(1) if m2 else "2026"


def _verify_by_balance(mov: Movimiento, prev_saldo: Optional[float]) -> None:
    """Corrige DEP/RET eligiendo el sentido (suma o resta) más cercano al saldo real."""
    if prev_saldo is None or mov.saldo is None:
        return
    amount = (mov.deposito or 0.0) + (mov.retiro or 0.0)
    if amount == 0:
        return
    diff_dep = abs((prev_saldo + amount) - mov.saldo)
    diff_ret = abs((prev_saldo - amount) - mov.saldo)
    # Solo corregir si al menos uno de los dos casos está dentro de 0.05
    if min(diff_dep, diff_ret) > 0.05:
        return
    if diff_dep <= diff_ret:
        mov.deposito = amount
        mov.retiro = None
    else:
        mov.retiro = amount
        mov.deposito = None


class BanbajioExtractor(BancoExtractor):
    NOMBRE = "Banbajío"
    COLOR_HEX = "004B87"   # Azul corporativo Banco del Bajío

    @classmethod
    def detectar(cls, texto: str) -> bool:
        return "BANBAJIO" in texto or "BANCO DEL BAJIO" in texto

    def extraer(
        self,
        pdf_path: str,
        progress_callback: Optional[Callable[[int, int], None]] = None,
    ) -> tuple[list[Movimiento], InfoCuenta]:
        movimientos: list[Movimiento] = []
        current: Optional[Movimiento] = None
        info = InfoCuenta(banco="Banbajío")
        in_section = False
        prev_saldo: Optional[float] = None
        anio = "2026"
        done = False

        # Mapa de columnas (CARGOS/ABONOS por posición X) — regla anti-traspaso.
        self._col_map = {}
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

                text = page.extract_text(x_tolerance=2, y_tolerance=3) or ""
                if not text:
                    continue

                if pnum == 0:
                    anio = _extract_year(text)
                    self._parse_header(text, info)

                for line in text.splitlines():
                    line = line.strip()
                    if not line:
                        continue

                    # Fin de transacciones
                    if "SALDO TOTAL" in line.upper() or "TOTAL DE MOVIMIENTOS" in line.upper():
                        if current:
                            _verify_by_balance(current, prev_saldo)
                            self._aplicar_columna(current)
                            movimientos.append(current)
                            current = None
                        done = True
                        break

                    # Filtrar líneas de encabezado / pie / CFDI
                    if any(p.upper() in line.upper() for p in SKIP_PATTERNS):
                        continue

                    # Inicio de sección de movimientos
                    if "SALDO INICIAL" in line.upper():
                        in_section = True
                        amts = AMOUNT_RE.findall(line)
                        if amts and info.saldo_inicial is None:
                            # Tomar el primer monto bien formateado
                            for a in amts:
                                val = float(a.replace(",", ""))
                                if val > 0:
                                    info.saldo_inicial = val
                                    prev_saldo = val
                                    break
                        continue

                    if not in_section:
                        continue

                    m = DATE_RE.match(line)
                    if m:
                        dia  = m.group(1).zfill(2)
                        mes  = MESES.get(m.group(2), m.group(2))
                        rest = m.group(3).strip()
                        fecha = f"{dia}/{mes}/{anio}"

                        if current:
                            _verify_by_balance(current, prev_saldo)
                            self._aplicar_columna(current)
                            prev_saldo = current.saldo
                            movimientos.append(current)

                        current = self._parse_line(fecha, rest)
                    elif current and line:
                        current.descripcion += " | " + line

            if current:
                _verify_by_balance(current, prev_saldo)
                self._aplicar_columna(current)
                movimientos.append(current)

        if movimientos:
            info.total_depositos = sum(m.deposito or 0 for m in movimientos)
            info.total_retiros   = sum(m.retiro   or 0 for m in movimientos)
            info.saldo_final     = movimientos[-1].saldo

        return movimientos, info

    def _aplicar_columna(self, mov: Optional[Movimiento]) -> None:
        """Regla anti-traspaso: si el balance no pudo decidir (saldo ausente),
        clasifica por la COLUMNA (CARGOS/ABONOS por posición X)."""
        if mov is None or mov.saldo is not None:
            return
        monto = (mov.retiro or 0.0) + (mov.deposito or 0.0)
        mapa = getattr(self, "_col_map", None)
        if not monto or not mapa:
            return
        lado = mapa.get((fecha_key(mov.fecha), int(round(monto * 100))))
        if lado == "deposito":
            mov.deposito, mov.retiro = monto, None
        elif lado == "retiro":
            mov.retiro, mov.deposito = monto, None

    def _parse_line(self, fecha: str, rest: str) -> Optional[Movimiento]:
        """Parsea el resto de la línea de transacción: [REF] DESCRIPCION $ MONTO $ SALDO."""
        amounts = AMOUNT_RE.findall(rest)

        saldo:      Optional[float] = None
        amount_val: Optional[float] = None
        desc = rest

        if len(amounts) >= 2:
            amount_val = float(amounts[-2].replace(",", ""))
            saldo      = float(amounts[-1].replace(",", ""))
            last_pos   = rest.rfind(amounts[-2])
            desc       = rest[:last_pos].strip()
        elif len(amounts) == 1:
            saldo    = float(amounts[0].replace(",", ""))
            last_pos = rest.rfind(amounts[0])
            desc     = rest[:last_pos].strip()

        desc = re.sub(r"\s+", " ", desc).strip()

        # Extraer folio si el primer token es puramente numérico
        folio = ""
        tokens = desc.split(None, 1)
        if tokens and tokens[0].isdigit():
            folio = tokens[0]
            desc  = tokens[1].strip() if len(tokens) > 1 else ""

        mov = Movimiento(fecha=fecha, folio=folio, descripcion=desc, saldo=saldo)
        if amount_val is not None:
            mov.retiro = amount_val   # _verify_by_balance corregirá si es depósito
        return mov

    def _parse_header(self, text: str, info: InfoCuenta):
        lines = text.splitlines()
        for i, line in enumerate(lines):
            line = line.strip()
            if "PERIODO:" in line.upper():
                info.periodo = line
            if "R.F.C." in line.upper():
                m = re.search(r"R\.F\.C\.\s+([A-Z0-9]{10,13})", line.upper())
                if m:
                    info.rfc = m.group(1)
            if "CLABE INTERBANCARIA" in line.upper():
                m = re.search(r"\b(\d{18})\b", line)
                if m:
                    info.clabe = m.group(1)
            if "CUENTA CONECTA BANBAJIO" in line.upper():
                m = re.search(r"(\d{13,16})", line)
                if m:
                    info.numero_cuenta = m.group(1)
