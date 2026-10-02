"""
conciliacion.py — Motor de conciliación banco ↔ CFDIs
Puerto fiel del módulo VBA ModConciliacion, adaptado a Python/pandas.

Fases de cruce (se aplican en orden a cada movimiento; el primer match gana):
  Fase 1: RFC/nombre + IMPORTE + MISMO MES (1:1) — solo bancarizadas
  Fase 2: Varios movimientos = UN mismo CFDI (N:1) — subset-sum sobre
          movimientos, mismo mes, exige RFC/nombre
  Fase 3: Un movimiento = varios CFDIs (1:N) — subset-sum sobre CFDIs
  Fase 4: PPD con complemento — verifica el REP (tipo P) por MontoPagado +
          RFC/nombre y |FechaPago − fecha banco| ≤ 1 día; vincula sus facturas
  Fase 5: Agrupado por RFC — grupos configurados + automático (suma por
          RFC+mes de CFDIs vs movimientos)
  Fase 6: PPD SIN complemento (anomalía) — exige RFC/nombre en la descripción
          bancaria (nunca por importe solo)
  Fase 7: Solo importe + MISMO MES (opcional, apagada por defecto)

Niveles internos de _buscar (Fases 1/4/6):
  Nivel 1: MONTO | MES | RFC  (exacto)
  Nivel 2: MONTO | MES + coincidencia parcial de nombre
  Dirección: mes_actual → 12, luego mes_actual-1 → 1

Salida:
  - mov_conciliados  : lista de MovConciliado  (banco + CFDI cruzado)
  - sin_cfdi         : movimientos sin CFDI encontrado
  - cfdi_sin_banco_r : CFDIs Recibidos sin movimiento bancario
  - cfdi_sin_banco_e : CFDIs Emitidos sin movimiento bancario
"""

from __future__ import annotations

import itertools
import re
import unicodedata
from collections import Counter
from dataclasses import dataclass, field, replace as _dc_replace
from datetime import datetime, date
from typing import Optional

from cfdi_parser import CFDI

# ── Regex helpers ─────────────────────────────────────────────────────────────
# Patrón bare de RFC (puede tener falsos positivos en cadenas SPEI como
# "036INBU30032026264373973" que contiene "INBU300320264")
_RFC_RE = re.compile(r"[A-ZÑ&]{3,4}[0-9]{6}[A-Z0-9]{3}", re.IGNORECASE)

# Patrón etiquetado: "RFC XXXXXX" — sin falsos positivos.
# En descripciones SPEI el RFC siempre viene precedido de la palabra "RFC".
_RFC_LABELED_RE = re.compile(
    r"\bRFC[:\s]+([A-ZÑ&]{3,4}[0-9]{6}[A-Z0-9]{3})\b",
    re.IGNORECASE,
)

_UUID_RE = re.compile(
    r"[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}",
    re.IGNORECASE
)

# Tolerancia de importe para match exacto (Fase 1): ±$1.00 MXN
TOLERANCIA_MXN: float = 1.00

# Tolerancia de centavos SOLO para nómina (redondeo del neto): ±$0.10 = 10 cent.
TOLERANCIA_NOMINA_CENT: int = 10

# Fase 4 (REP): días de desfase permitidos entre la FechaPago del complemento y
# el movimiento bancario cuando NO caen en el mismo mes. Cubre el pago de fin de
# mes que el banco refleja en los primeros días del siguiente.
TOLERANCIA_DIAS_REP: int = 5

# Formas de pago SAT que producen un movimiento bancario real y por lo tanto
# son elegibles para conciliar automáticamente contra un estado de cuenta
# (Fases 1-3). Lista alineada con la skill conciliacion-cfdi-banco:
#   02 Cheque nominativo        03 Transferencia electrónica
#   04 Tarjeta de crédito       05 Monedero electrónico
#   06 Dinero electrónico       28 Tarjeta de débito
#   29 Tarjeta de servicios     30 Aplicación de anticipos
#   31 Intermediario de pagos
# Efectivo (01), condonación, compensación, etc. nunca aparecen en el banco —
# no deben ser candidatos, sin importar que monto/mes/RFC coincidan.
# PPD (típicamente FormaPago="99") se resuelve aparte, en Fase 4, vía su
# complemento de pago — no pasa por este filtro.
FORMAS_PAGO_BANCARIZADAS: frozenset[str] = frozenset(
    {"02", "03", "04", "05", "06", "28", "29", "30", "31"}
)

# Movimientos bancarios EXCLUIDOS de la conciliación por regla de negocio:
# un PRÉSTAMO (recibido u otorgado, y sus pagos) no es una operación facturable
# — no existe CFDI que lo respalde — así que NUNCA debe cruzarse contra un
# CFDI aunque monto/mes/RFC coincidan. Estos movimientos se apartan ANTES de
# todas las fases y quedan como pendientes con estatus propio.
PALABRAS_NO_CONCILIABLES: tuple[str, ...] = ("PRESTAMO", "PRÉSTAMO")


def _es_no_conciliable(mov: "MovBanco") -> bool:
    """True si la descripción del movimiento lo excluye de la conciliación."""
    d = (mov.descripcion or "").upper()
    return any(kw in d for kw in PALABRAS_NO_CONCILIABLES)


# ── Catálogo de fases de conciliación ─────────────────────────────────────────
# (id, etiqueta legible). El id se usa para activar/desactivar cada fase desde
# la UI (casillas). Cualquier fase cuyo id NO esté en el conjunto `fases` que se
# pasa a conciliar()/conciliar_multibanco() simplemente no se ejecuta.
FASES_CONCILIACION: list[tuple[str, str]] = [
    ("exacto",       "Fase 1 — Exacto (RFC/nombre + importe, mismo mes)"),
    ("ppd_con_rfc",  "Fase 2 — PPD con complemento, por RFC"),
    ("ppd_con_nom",  "Fase 3 — PPD con complemento, por Nombre"),
    ("ppd_sin_rfc",  "Fase 4 — PPD sin complemento, por RFC"),
    ("ppd_sin_nom",  "Fase 5 — PPD sin complemento, por Nombre"),
    ("combinatorio", "Fase 6 — Un movimiento = varios CFDIs (RFC/nombre)"),
    ("multi_mov",    "Fase 7 — Varios movimientos = un mismo CFDI (RFC/nombre)"),
    ("grupos",       "Fase 8 — Agrupado por RFC (grupos config. + automático)"),
    ("solo_importe", "Fase 9 — Solo importe, mismo mes (último recurso, opcional)"),
]
FASES_IDS: list[str] = [f[0] for f in FASES_CONCILIACION]
# Por defecto se ejecutan todas MENOS la Fase 7: cruzar solo por importe (sin
# mes ni identidad) es el último recurso y es opcional — se activa a mano.
FASES_DEFAULT: frozenset[str] = frozenset(set(FASES_IDS) - {"solo_importe"})


# ── Data classes ──────────────────────────────────────────────────────────────

@dataclass
class MovBanco:
    """Movimiento bancario enriquecido para conciliación."""
    fecha:       str   = ""
    folio:       str   = ""
    descripcion: str   = ""
    retiro:      float = 0.0
    deposito:    float = 0.0
    saldo:       float = 0.0
    mes:         int   = 0
    id_banco:    str   = ""
    orden:       int   = 0     # posición en el estado de cuenta (para no reordenar)
    # Campos llenados por conciliación
    uuid:           str   = ""
    rfc:            str   = ""
    nombre:         str   = ""
    fecha_cfdi:     str   = ""   # fecha emisión del CFDI
    fecha_pago_cfdi: str  = ""   # fecha pago (solo nómina tipo N)
    serie_cfdi:     str   = ""   # Serie del CFDI conciliado
    folio_cfdi:     str   = ""   # Folio del CFDI conciliado
    uuid_rep:       str   = ""   # UUID del complemento de pago (REP tipo P) vinculado, si existe
    estatus:        str   = ""
    uuid_suma:      str   = ""   # UUIDs adicionales en suma (nivel 4), separados por ","

    @property
    def monto(self) -> float:
        """Monto relevante para conciliación (retiro si existe, si no deposito)."""
        return self.retiro if self.retiro else self.deposito

    @property
    def monto_str(self) -> str:
        return f"{round(self.monto, 2):.2f}"

    def rfc_en_descripcion(self) -> str:
        texto = self.descripcion.upper()
        # Prioridad 1: RFC etiquetado explícitamente ("RFC XXXXXX")
        # Las descripciones SPEI siempre incluyen "RFC <valor>" al final.
        # Evita falsos positivos como "INBU300320264" dentro de
        # "036INBU30032026264373973" (clave de rastreo SPEI).
        m = _RFC_LABELED_RE.search(texto)
        if m:
            return m.group(1).upper()
        # Prioridad 2: patrón bare (más propenso a falsos positivos)
        m = _RFC_RE.search(texto)
        return m.group(0).upper() if m else ""

    def nombre_en_descripcion(self) -> str:
        """
        Extrae el nombre del beneficiario/ordenante de la descripción bancaria.
        Estrategias en orden de prioridad:
          1. Texto después de 'CLIENTE'
          2. Texto después de 'NOMBRE:' o 'NOMBRE '
          3. Segmentos entre '/' que parezcan nombre (SPEI típico)
        """
        desc  = self.descripcion
        upper = desc.upper()

        # 1. Keyword "CLIENTE"
        pos = upper.find("CLIENTE")
        if pos >= 0:
            return desc[pos + 7:].strip()

        # 1.5 Keyword "BENEFICIARIO" / "BENEFICIARIO TEF" (Banbajío y pagos TEF).
        #     El nombre puede venir TRUNCADO y con "(BI-...)" pegado al final; se
        #     corta en el primer separador de ruido. Las variantes largas van
        #     primero para no dejar "TEF:" dentro del nombre.
        for kw in ("BENEFICIARIO TEF:", "BENEFICIARIO TEF",
                   "BENEFICIARIO:", "BENEFICIARIO"):
            pos = upper.find(kw)
            if pos >= 0:
                seg = desc[pos + len(kw):].strip(" :")
                for sep in ("(", "|", "\n", "/", " RFC", " CLABE", " CUENTA"):
                    idx = seg.upper().find(sep.upper())
                    if 0 < idx < len(seg):
                        seg = seg[:idx].strip()
                        break
                if len(seg) >= 4:
                    return seg

        # 2. Keyword "NOMBRE"
        for kw in ("NOMBRE:", "NOMBRE "):
            pos = upper.find(kw)
            if pos >= 0:
                seg = desc[pos + len(kw):].strip()
                for sep in ("/", "|", "\n", " RFC", " CLABE"):
                    idx = seg.upper().find(sep.upper())
                    if 0 < idx < len(seg):
                        seg = seg[:idx].strip()
                        break
                if len(seg) >= 4:
                    return seg

        # 2.5 SPEI Banregio con campos separados por COMA:
        #     "SPEI, <banco ordenante>, <NOMBRE ORDENANTE>, <referencia>, <concepto>"
        #     El nombre del ordenante es el 3er campo. Gated a descripciones que
        #     empiezan con "SPEI" para no tocar otros bancos/formatos (Inbursa usa
        #     '|', el SPEI estándar usa '/'). Habilita el match por nombre (nivel 2)
        #     en depósitos Banregio que sólo traen el nombre, no el RFC.
        if upper.startswith("SPEI") and "," in desc:
            campos = [c.strip() for c in desc.split(",")]
            if len(campos) >= 3:
                cand = campos[2]
                letras = sum(1 for ch in cand if ch.isalpha() or ch == " ")
                if len(cand) >= 5 and letras / max(len(cand), 1) >= 0.7:
                    return cand

        # 3. Segmentos entre '/' — formato SPEI estándar
        #    Ejemplo: "CONCEPTO / JUAN GARCIA LOPEZ / RFC GALJ800101AB1 / CLABE 012..."
        partes = [p.strip() for p in desc.split("/") if p.strip()]
        if len(partes) >= 2:
            candidatos = []
            for p in partes:
                # Descartar CLABEs (≥10 dígitos seguidos), montos, RFCs
                if re.search(r"\d{10,}", p):
                    continue
                if _RFC_LABELED_RE.search(p) or _RFC_RE.search(p.strip()):
                    continue
                letras = sum(1 for ch in p if ch.isalpha() or ch == " ")
                if len(p) >= 5 and letras / max(len(p), 1) >= 0.55:
                    candidatos.append(p)
            if candidatos:
                return max(candidatos, key=len)

        return ""


@dataclass
class MovConciliado:
    banco:      MovBanco
    cfdi:       Optional[CFDI]
    tipo:       str   = ""    # "recibido" | "emitido" | ""
    revisar:    bool  = False
    nivel:      int   = 0     # 1=RFC exacto 2=nombre 3=nombre-en-desc 4=suma multi-CFDI
    cfdi_suma:  list  = field(default_factory=list)   # CFDIs adicionales (solo nivel 4)
    tipo_match: str   = ""    # "Exacto" | "Global Mensual" | "Combinatorio"


# ── Helpers de coincidencia ───────────────────────────────────────────────────

# Partículas y términos genéricos que NO deben usarse para identificar un nombre
_PARTICULAS = {
    "DE", "DEL", "LA", "LAS", "LOS", "EL", "Y", "E", "O",
    "SA", "SAB", "SRL", "SCV", "SPR", "SC", "CV", "SAPI",
    "MEXICO", "MEXICANA", "MEXICANO",
}


from functools import lru_cache as _lru_cache


@_lru_cache(maxsize=100_000)
def _sin_acentos(s: str) -> str:
    """Quita diacríticos para comparar nombres (NIÑO=NINO, PEÑA=PENA, JOSÉ=JOSE).
    Cacheado: en conciliaciones grandes se llama millones de veces sobre las
    MISMAS cadenas (nombres/descripciones repetidos) — era un hot spot de
    unicodedata."""
    return "".join(
        c for c in unicodedata.normalize("NFD", s)
        if unicodedata.category(c) != "Mn"
    )


def _palabras_nombre(texto: str) -> set[str]:
    """Palabras significativas de un nombre: ≥4 chars, no partícula, no dígitos.
    Se comparan SIN acentos para tolerar variantes de tildes/Ñ entre el CFDI y
    la descripción bancaria."""
    return {
        w for w in _sin_acentos(texto.upper()).split()
        if len(w) >= 4 and w not in _PARTICULAS and not w.isdigit()
    }


def _nombre_coincide(nom_a: str, nom_b: str) -> bool:
    """
    Coincidencia de nombres tolerante al orden (nombre primero vs apellido primero).
    Reglas:
      • Una palabra ≥7 chars en común es suficiente  (apellido o razón social distinctive)
      • ≥2 palabras ≥4 chars en común también es válido
    La búsqueda es bidireccional: se usan palabras de ambas cadenas.
    """
    pa = _palabras_nombre(nom_a)
    pb = _palabras_nombre(nom_b)
    if not pa or not pb:
        return False
    comunes = pa & pb
    if not comunes:
        return False
    # Un apellido/razón social largo es evidencia fuerte
    if any(len(p) >= 7 for p in comunes):
        return True
    # Dos palabras cortas (4-6 chars) también son suficientes
    return len(comunes) >= 2


def _nombre_en_desc(desc_upper: str, nom_cfdi: str) -> bool:
    """
    Verifica si palabras significativas del nombre CFDI aparecen en la descripción
    bancaria completa (útil cuando nombre_en_descripcion() no extrajo nada).
    Mismas reglas que _nombre_coincide: 1 palabra ≥7 chars O ≥2 palabras ≥4 chars.
    """
    palabras = _palabras_nombre(nom_cfdi)
    if not palabras:
        return False
    desc_norm = _sin_acentos(desc_upper)
    presentes = {p for p in palabras if p in desc_norm}
    if not presentes:
        return False
    if any(len(p) >= 7 for p in presentes):
        return True
    return len(presentes) >= 2


# Abreviaturas de mes en español (Inbursa: "04/JUN./2026", Banorte, etc.)
_MESES_ES: dict[str, int] = {
    "ENE": 1, "FEB": 2, "MAR": 3, "ABR": 4, "MAY": 5, "JUN": 6,
    "JUL": 7, "AGO": 8, "SEP": 9, "SET": 9, "OCT": 10, "NOV": 11, "DIC": 12,
}


def _mes_fecha(fecha_str: str) -> int:
    """Extrae el mes de una cadena de fecha.
    Soporta: DD/MM/YYYY · YYYY-MM-DD · DD-MM-YYYY · DD/MM/YY · DD/MM (BBVA) ·
             DD/MMM/YYYY con mes en letras español (Inbursa: "04/JUN./2026").
    """
    s = fecha_str.strip()
    for fmt, n in [
        ("%d/%m/%Y", 10), ("%Y-%m-%d", 10),
        ("%d-%m-%Y", 10), ("%d/%m/%y", 8),
        ("%d/%m",    5),   # BBVA: "30/03"
    ]:
        try:
            return datetime.strptime(s[:n], fmt).month
        except ValueError:
            pass
    # Mes en letras (español): "04/JUN./2026", "04-JUN-2026", "04 JUN 2026".
    # Debe ir ANTES del último recurso numérico (que tomaría el DÍA como mes).
    m = re.search(r"[/\-\s]([A-ZÑ]{3})\.?[/\-\s]", s.upper())
    if m and m.group(1) in _MESES_ES:
        return _MESES_ES[m.group(1)]
    # Último recurso: segundo bloque numérico separado por / o -
    m = re.search(r"(?:^|[/\-])(\d{1,2})[/\-]", s)
    return int(m.group(1)) if m else 0


def _dia_fecha(fecha_str: str) -> int:
    """Extrae el DÍA de una cadena de fecha — para casar la FechaPago de la
    nómina con el DÍA del movimiento bancario. Soporta los mismos formatos que
    _mes_fecha, incluido el mes en letras de Inbursa ('15/JUN./2026' -> 15)."""
    s = fecha_str.strip()
    for fmt, n in [
        ("%d/%m/%Y", 10), ("%Y-%m-%d", 10),
        ("%d-%m-%Y", 10), ("%d/%m/%y", 8), ("%d/%m", 5),
    ]:
        try:
            return datetime.strptime(s[:n], fmt).day
        except ValueError:
            pass
    # DD al inicio ("15/JUN./2026", "15 JUN 2026", "15-JUN-2026")
    m = re.match(r"\s*(\d{1,2})[/\-\s]", s)
    return int(m.group(1)) if m else 0


def _fecha_partes(fecha_str: str) -> tuple[int, int, int | None] | None:
    """(día, mes, año|None) de una cadena de fecha; None si no se pudo parsear.
    Soporta los mismos formatos que _mes_fecha/_dia_fecha."""
    s = fecha_str.strip()
    for fmt, n in [("%d/%m/%Y", 10), ("%Y-%m-%d", 10),
                   ("%d-%m-%Y", 10), ("%d/%m/%y", 8)]:
        try:
            d = datetime.strptime(s[:n], fmt)
            return (d.day, d.month, d.year)
        except ValueError:
            pass
    # Mes en letras con año: "04/JUN./2026", "04-JUN-2026", "04 JUN 2026"
    m = re.search(r"(\d{1,2})[/\-\s]([A-ZÑ]{3})\.?[/\-\s](\d{4})", s.upper())
    if m and m.group(2) in _MESES_ES:
        return (int(m.group(1)), _MESES_ES[m.group(2)], int(m.group(3)))
    # DD/MM sin año (BBVA "30/03")
    m = re.match(r"\s*(\d{1,2})[/\-](\d{1,2})(?![/\-]?\d)", s)
    if m:
        return (int(m.group(1)), int(m.group(2)), None)
    return None


def _dias_entre(f1: str, f2: str) -> Optional[int]:
    """Diferencia ABSOLUTA en días entre dos fechas (str). None si no se pueden
    parsear. Si a alguna le falta el año, usa el de la otra (o 2000)."""
    p1, p2 = _fecha_partes(f1), _fecha_partes(f2)
    if not p1 or not p2:
        return None
    y1 = p1[2] or p2[2] or 2000
    y2 = p2[2] or p1[2] or 2000
    try:
        d1 = date(y1, p1[1], p1[0])
        d2 = date(y2, p2[1], p2[0])
    except ValueError:
        return None
    return abs((d1 - d2).days)


def _subset_sum_idx(vals: list[int], target: int, max_terms: int,
                    tol: int = 0) -> list[int]:
    """Indices de una combinacion de 2..max_terms elementos de `vals` (centavos)
    cuya suma llega a `target`, o [] si no existe.

    `tol` = tolerancia en CENTAVOS (0 = suma exacta, comportamiento original).
    Con tolerancia se prefiere SIEMPRE la coincidencia exacta; si no la hay, se
    devuelve la combinacion mas cercana dentro de +/-tol.
    """
    if len(vals) < 2:
        return []
    limite = target + tol
    srt = sorted(vals)
    if srt[0] + srt[1] > limite:
        return []
    dp: dict[int, list[int]] = {0: []}
    mejor: list[int] = []
    mejor_dif: int | None = None
    for idx, v in enumerate(vals):
        for s, comb in list(dp.items()):
            ns = s + v
            if ns > limite:
                continue
            nc = comb + [idx]
            if len(nc) > max_terms:
                continue
            # El objetivo se comprueba ANTES del memo: si esta combinacion llega al
            # target con >=2 terminos es valida aunque esa suma ya se hubiera
            # alcanzado por otro camino. Antes se descartaba aqui y se perdian
            # combinaciones legitimas.
            if len(nc) >= 2:
                dif = abs(ns - target)
                if dif == 0:
                    return nc                     # exacto: inmejorable
                if dif <= tol and (mejor_dif is None or dif < mejor_dif):
                    mejor, mejor_dif = nc, dif
            # Memo: conservar el camino con MENOS terminos (mas extensible).
            prev = dp.get(ns)
            if prev is None or len(nc) < len(prev):
                dp[ns] = nc
    return mejor


def _mov_coincide_identidad(mov: "MovBanco", rfc: str, nom: str) -> bool:
    """True si el RFC o el nombre del tercero aparece/coincide en la descripción
    del movimiento bancario (RFC exacto, nombre tolerante al orden, o nombre-en-
    texto). Sirve para exigir identidad en las fases de suma (2, 5 auto)."""
    if rfc:
        rfc_m = mov.rfc_en_descripcion()
        if rfc_m and rfc_m.upper() == rfc.upper():
            return True
    if nom:
        nom_m = mov.nombre_en_descripcion()
        if nom_m and _nombre_coincide(nom_m, nom):
            return True
        if _nombre_en_desc(mov.descripcion.upper(), nom):
            return True
    return False


# Instituciones financieras: emisor de las facturas de comisión bancaria.
# Sin acentos (se compara con _sin_acentos), ordenadas de más a menos específicas.
_BANCO_EMISOR_KW = (
    "BBVA", "BANORTE", "SANTANDER", "INBURSA", "BANREGIO", "BANBAJIO",
    "BAJIO", "BANCOMER", "SCOTIABANK", "HSBC", "AFIRME", "MIFEL", "MULTIVA",
    "ACTINVER", "INTERCAM", "BANCO",   # "BANCO" al final: genérico, último recurso
)


# 'IVA' como palabra suelta: acepta I.V.A e I V A (OCR), pero NO dentro de
# PRIVADA/ACTIVA/DERIVADO… (lookaround para no exigir límites de palabra clásicos,
# que fallan con la forma espaciada 'I V A').
_IVA_SUELTA_RE = re.compile(r"(?<![A-Z])I\.?\s*V\.?\s*A(?![A-Z])")


def _es_emisor_banco(cfdi: "CFDI") -> bool:
    """True si el CFDI lo emitió una institución financiera. Sirve para cruzar
    las comisiones bancarias (cuyo movimiento no nombra al banco) contra la
    factura del banco por identidad del EMISOR, no de la descripción."""
    ne = _sin_acentos((cfdi.nombre_emisor or "").upper())
    return any(kw in ne for kw in _BANCO_EMISOR_KW)


def _es_mov_comision(mov: "MovBanco") -> bool:
    """True si el movimiento es una comisión bancaria O su IVA. Genérico para
    CUALQUIER banco (Banorte trae la comisión y el IVA como movimientos
    SEPARADOS, y la línea de IVA no dice "COMISION"). Tolera el espaciado del
    OCR ('C O MISION', 'I V A', 'I.V.A')."""
    up = _sin_acentos((mov.descripcion or "").upper())
    ns = up.replace(" ", "")
    if "COMISION" in ns:
        return True
    # Cuota SPEI de Banorte: la comisión de envío NO dice "COMISION"; viene como
    # "TRANSFERENCIA - ENVIO" (el pago real es "ORDEN DE PAGO SPEI"/"TRASPASO").
    if "TRANSFERENCIA-ENVIO" in ns:
        return True
    # Campo informativo 'IVA:0.00' que Banorte mete en CADA traspaso/SPEI — NO es
    # una comisión. Debe excluirse ANTES de la detección de 'IVA' suelta.
    if "IVA:" in ns:
        return False
    # 'IVA' como línea real (I.V.A. ORDEN DE PAGO, IVA COMISION…): palabra suelta,
    # no dentro de PRIVADA/ACTIVA/DERIVADO; acepta I.V.A / I V A del OCR.
    if _IVA_SUELTA_RE.search(up):
        return True
    return False


# ── Notas de crédito ─────────────────────────────────────────────────────────

def _aplicar_notas_credito(
    cfdis: list[CFDI],
) -> tuple[list[CFDI], set[str]]:
    """
    Aplica las notas de crédito tipo "E" a sus facturas relacionadas.

    Para cada CFDI tipo "E" que tenga CfdiRelacionados apuntando a una
    factura del mismo pool, se reduce el total efectivo de esa factura y
    la nota de crédito se marca como 'absorbida' (no aparece como
    CFDI sin movimiento).

    Returns
    -------
    cfdis_ajustados : list[CFDI]
        Lista sin las notas absorbidas y con totales de facturas reducidos.
        Las copias se crean con dataclasses.replace() — los originales
        no se modifican.
    absorbidos : set[str]
        UUIDs (lower) de las notas de crédito que fueron aplicadas.
    """
    por_uuid: dict[str, CFDI] = {c.uuid.lower(): c for c in cfdis}
    netos:    dict[str, float] = {}   # uuid_lower → total ajustado
    absorbidos: set[str] = set()

    for c in cfdis:
        if c.tipo != "E" or not c.cfdi_relacionados:
            continue
        for uuid_rel in c.cfdi_relacionados:
            k = uuid_rel.lower()
            if k not in por_uuid:
                continue
            # Reducir el total de la factura relacionada
            base = netos.get(k, por_uuid[k].total)
            netos[k] = max(0.0, round(base - c.total, 2))
            absorbidos.add(c.uuid.lower())
            break  # cada nota aplica a la primera factura encontrada

    if not netos and not absorbidos:
        return cfdis, absorbidos

    resultado: list[CFDI] = []
    for c in cfdis:
        k = c.uuid.lower()
        if k in absorbidos:
            continue  # nota de crédito absorbida — no incluir en el pool
        if k in netos:
            c = _dc_replace(c, total=netos[k])  # copia con total reducido
        resultado.append(c)
    return resultado, absorbidos


def _uuids_con_complemento(cfdis: list[CFDI]) -> set[str]:
    """UUIDs de facturas PPD que ya están referenciadas en un complemento de pago tipo P."""
    resultado: set[str] = set()
    for c in cfdis:
        if c.tipo == "P":
            for uid in c.cfdi_facturas_pagadas:
                resultado.add(uid.lower())
    return resultado


# ── Motor de conciliación ─────────────────────────────────────────────────────

class Conciliador:
    """
    Concilia una lista de MovBanco con CFDIs Recibidos y Emitidos.

    Parámetros
    ----------
    movimientos   : lista de MovBanco
    recibidos     : CFDIs donde la empresa es RECEPTOR  (retiros del banco)
    emitidos      : CFDIs donde la empresa es EMISOR    (depósitos del banco)
    """

    def __init__(self,
                 movimientos: list[MovBanco],
                 recibidos:   list[CFDI],
                 emitidos:    list[CFDI]):
        self.movimientos = movimientos
        self.recibidos   = recibidos
        self.emitidos    = emitidos

    # ── API pública ───────────────────────────────────────────────────────────

    def conciliar(self, skip_suma: bool = False,
                  grupos: list | None = None,
                  fases: "set[str] | frozenset[str] | None" = None,
                  bitacora: list | None = None) -> tuple[
            list[MovConciliado],  # conciliados
            list[MovBanco],       # sin CFDI
            list[CFDI],           # CFDIs recibidos sin banco
            list[CFDI],           # CFDIs emitidos sin banco
        ]:
        """
        skip_suma=True : solo niveles 1-3 (rápido, usado en Pasada 1 multi-banco).
        grupos         : list[GrupoConciliacion] — Fase 5 multi-RFC (opcional).
        fases          : conjunto de ids de fase a ejecutar (ver FASES_IDS).
                         None → FASES_DEFAULT (todas menos la Fase 7 "Solo
                         importe"). Una fase cuyo id no esté en el conjunto se
                         omite por completo.
        bitacora       : lista donde la Fase 7 anota cada cruce por-importe
                         (para evitar duplicados / auditoría). Opcional.
        """
        if fases is None:
            fases = FASES_DEFAULT
        # Posición original en el estado de cuenta (para que la hoja NO reordene).
        for _i, _m in enumerate(self.movimientos):
            _m.orden = _i
        # Aplicar notas de crédito (tipo "E") antes de construir los índices.
        # Esto reduce el total de las facturas relacionadas y excluye las notas
        # absorbidas del pool de búsqueda.
        recibidos_adj, absorbidos_r = _aplicar_notas_credito(self.recibidos)
        emitidos_adj,  absorbidos_e = _aplicar_notas_credito(self.emitidos)

        # Mapas uuid -> CFDI para lookups O(1) (evita next(...) lineal)
        por_uuid_rec = {c.uuid.lower(): c for c in recibidos_adj}
        por_uuid_emi = {c.uuid.lower(): c for c in emitidos_adj}

        usados_r: set[str] = set(absorbidos_r)
        usados_e: set[str] = set(absorbidos_e)

        # Fases 1-3 (matching directo contra un movimiento bancario) solo
        # consideran CFDIs con forma de pago bancarizada — efectivo,
        # condonación, compensación, etc. nunca producen un movimiento real
        # en el estado de cuenta. PPD (forma_pago típicamente "99") se
        # excluye aquí a propósito: se resuelve en Fase 4 vía su complemento
        # de pago, usando el pool SIN filtrar (recibidos_adj/emitidos_adj).
        # NÓMINA (tipo "N") SIEMPRE es elegible: la dispersión de sueldos es un
        # movimiento bancario real aunque su FormaPago sea "99" u otra no
        # bancarizada — de lo contrario nunca conciliaría.
        # Los tipo P (complementos de pago / REP) NUNCA entran a las fases de
        # factura directa (1-3, 5): los resuelve la Fase 4 vinculando sus
        # facturas pagadas — si no, un REP con forma bancarizada cruzaría por
        # importe en Fase 1 y saltaría el enlace complemento→factura.
        recibidos_banc = [c for c in recibidos_adj
                          if c.tipo != "P"
                          and (c.forma_pago in FORMAS_PAGO_BANCARIZADAS or c.tipo == "N")]
        emitidos_banc  = [c for c in emitidos_adj
                          if c.tipo != "P"
                          and (c.forma_pago in FORMAS_PAGO_BANCARIZADAS or c.tipo == "N")]

        # recibidos: empresa es RECEPTORA → el tercero es el EMISOR del CFDI
        # emitidos:  empresa es EMISORA   → el tercero es el RECEPTOR del CFDI
        dr_monto_mes_rfc, dr_monto_mes, dr_nom = self._build_dicts(recibidos_banc, tercero_es_emisor=True)
        de_monto_mes_rfc, de_monto_mes, de_nom = self._build_dicts(emitidos_banc,  tercero_es_emisor=False)

        # Día de la FechaPago de cada NÓMINA — para exigir que coincida con el
        # día del movimiento bancario (una quincena no toma el pago de la otra).
        dia_nomina = {c.uuid.lower(): _dia_fecha(c.fecha_pago)
                      for c in recibidos_banc + emitidos_banc
                      if c.tipo == "N" and c.fecha_pago}

        conciliados: list[MovConciliado] = []
        sin_cfdi:    list[MovBanco]      = []
        # Préstamos: apartados de TODAS las fases (no son facturables);
        # se reincorporan a sin_cfdi al final, después de la última fase.
        no_conciliables: list[MovBanco]  = []

        # Fase 1 — Exacto: solo se ejecuta si su id está en `fases`.
        ejecutar_f1 = "exacto" in fases

        for mov in self.movimientos:
            if _es_no_conciliable(mov):
                mov.estatus = "Préstamo — no conciliable"
                no_conciliables.append(mov)
                continue

            if mov.monto == 0:
                sin_cfdi.append(mov)
                continue

            es_retiro   = mov.retiro   > 0
            es_deposito = mov.deposito > 0

            # mismo_mes=True: Fase 1 solo concilia CFDIs bancarizados (PUE u
            # otra forma bancarizada) del MISMO mes que el movimiento — un
            # PUE de enero nunca debe conciliar contra un banco de junio.
            # El cruce de meses queda reservado exclusivamente a PPD vía
            # Fase 4 (más abajo), que sí busca su complemento de pago.
            uuid_r, nivel_r = (
                self._buscar(mov, dr_monto_mes_rfc, dr_monto_mes, dr_nom, usados_r,
                             mismo_mes=True, dia_nomina=dia_nomina)
                if ejecutar_f1 and es_retiro else (None, 0)
            )
            uuid_e, nivel_e = (
                self._buscar(mov, de_monto_mes_rfc, de_monto_mes, de_nom, usados_e,
                             mismo_mes=True, dia_nomina=dia_nomina)
                if ejecutar_f1 and es_deposito else (None, 0)
            )

            if uuid_r:
                usados_r.add(uuid_r.lower())
                cfdi = por_uuid_rec.get(uuid_r.lower())
                if cfdi:
                    self._llenar_mov(mov, cfdi, tipo="recibido")
                conciliados.append(MovConciliado(mov, cfdi, "recibido",
                                                 revisar=(nivel_r >= 3),
                                                 nivel=nivel_r,
                                                 tipo_match="Exacto"))
            elif uuid_e:
                usados_e.add(uuid_e.lower())
                cfdi = por_uuid_emi.get(uuid_e.lower())
                if cfdi:
                    self._llenar_mov(mov, cfdi, tipo="emitido")
                conciliados.append(MovConciliado(mov, cfdi, "emitido",
                                                 revisar=(nivel_e >= 3),
                                                 nivel=nivel_e,
                                                 tipo_match="Exacto"))
            else:
                mov.estatus = "Pendiente de CFDI"
                sin_cfdi.append(mov)

        # ── Fase 2 — PPD con complemento, por RFC (REP) ──────────────────────
        # Verifica el complemento de pago (tipo P): casa su MontoPagado contra el
        # banco exigiendo el RFC del tercero en la descripción (identidad fuerte),
        # con la FechaPago del REP en el mismo mes o ±1 día. Al cruzar vincula el
        # REP y TODAS sus facturas. Va antes de las fases de suma para que un REP
        # reclame sus facturas antes de que un subset-sum las tome por importe.
        if not skip_suma and sin_cfdi and "ppd_con_rfc" in fases:
            mc, sin_cfdi = self._conciliar_rep(
                sin_cfdi, recibidos_adj, emitidos_adj, usados_r, usados_e,
                modo="rfc")
            conciliados.extend(mc)

        # ── Fase 3 — PPD con complemento, por Nombre (REP) ───────────────────
        # Igual que la Fase 2 pero cruza por NOMBRE del tercero (sin RFC), sobre
        # lo que la Fase 2 dejó pendiente. El RFC (más fuerte) siempre va primero.
        if not skip_suma and sin_cfdi and "ppd_con_nom" in fases:
            mc, sin_cfdi = self._conciliar_rep(
                sin_cfdi, recibidos_adj, emitidos_adj, usados_r, usados_e,
                modo="nombre")
            conciliados.extend(mc)

        # ── Fase 4 — PPD sin complemento, por RFC ────────────────────────────
        # PPD sin complemento de pago: hay movimiento en banco pero falta emitir
        # el complemento. Se concilia (pendiente de regularizar) exigiendo el RFC
        # del tercero en la descripción bancaria — nunca por importe solo.
        if not skip_suma and sin_cfdi and "ppd_sin_rfc" in fases:
            mc, sin_cfdi = self._conciliar_ppd(
                sin_cfdi, recibidos_adj, emitidos_adj, usados_r, usados_e,
                con_complemento=False, tipo_match="PPD Sin Complemento",
                nivel=7, solo_identidad=True, modo="rfc",
            )
            conciliados.extend(mc)

        # ── Fase 5 — PPD sin complemento, por Nombre ─────────────────────────
        # Igual que la Fase 4 pero por NOMBRE del tercero, sobre lo que quedó.
        if not skip_suma and sin_cfdi and "ppd_sin_nom" in fases:
            mc, sin_cfdi = self._conciliar_ppd(
                sin_cfdi, recibidos_adj, emitidos_adj, usados_r, usados_e,
                con_complemento=False, tipo_match="PPD Sin Complemento",
                nivel=7, solo_identidad=True, modo="nombre",
            )
            conciliados.extend(mc)

        # ── Fase 6 — Un movimiento = varios CFDIs (1:N) ──────────────────────
        # Subset-sum sobre los CFDIs bancarizados del mismo mes, priorizando
        # RFC → nombre extraído → nombre-en-descripción (ver _buscar_suma).
        if not skip_suma and "combinatorio" in fases:
            pendientes_f3: list[MovBanco] = []
            for mov in sin_cfdi:
                es_ret = mov.retiro   > 0
                es_dep = mov.deposito > 0
                suma_f3: list[CFDI] = []
                tipo_f3 = ""
                if es_ret:
                    suma_f3 = self._buscar_suma(mov, recibidos_banc, usados_r)
                    tipo_f3 = "recibido" if suma_f3 else ""
                elif es_dep:
                    suma_f3 = self._buscar_suma(mov, emitidos_banc, usados_e)
                    tipo_f3 = "emitido" if suma_f3 else ""
                if suma_f3:
                    usados_f3 = usados_r if tipo_f3 == "recibido" else usados_e
                    for c in suma_f3:
                        usados_f3.add(c.uuid.lower())
                    self._llenar_mov(mov, suma_f3[0], tipo=tipo_f3)
                    mov.estatus   = "Conciliado SUMA"
                    mov.uuid_suma = ",".join(c.uuid for c in suma_f3[1:])
                    conciliados.append(MovConciliado(
                        mov, suma_f3[0], tipo_f3,
                        revisar=True, nivel=4, cfdi_suma=suma_f3[1:],
                        tipo_match="Combinatorio",
                    ))
                else:
                    mov.estatus = "Pendiente de CFDI"
                    pendientes_f3.append(mov)
            sin_cfdi = pendientes_f3

        # ── Fase 7 — Varios movimientos = UN mismo CFDI (N:1) ────────────────
        # Un CFDI liquidado en varias parcialidades bancarias del mismo RFC/
        # nombre y mismo mes (subset-sum sobre los MOVIMIENTOS). Exige identidad
        # (RFC o nombre en la descripción de cada movimiento).
        if not skip_suma and "multi_mov" in fases:
            mc_mm, sin_cfdi = self._conciliar_multimov(
                sin_cfdi, recibidos_banc, emitidos_banc, usados_r, usados_e)
            conciliados.extend(mc_mm)

            # ── Fase 7b — Varios movimientos = UNA factura PPD 99 (N:1) ──────
            # Las PPD 99 se excluyen del pool bancarizado (recibidos_banc/
            # emitidos_banc), así que la pasada de arriba nunca las ve. Una
            # factura PPD pagada en parcialidades bancarias (parcialidad 1, 2…)
            # solo la cubre esta pasada: subset-sum sobre los MOVIMIENTOS
            # pendientes contra el TOTAL de la PPD, exigiendo identidad
            # (RFC/nombre) en cada movimiento y SIN candado de mes (las
            # parcialidades PPD pueden caer en meses distintos al de emisión).
            ppd_r = [c for c in recibidos_adj
                     if c.metodo_pago == "PPD" and c.forma_pago == "99"
                     and c.uuid.lower() not in usados_r]
            ppd_e = [c for c in emitidos_adj
                     if c.metodo_pago == "PPD" and c.forma_pago == "99"
                     and c.uuid.lower() not in usados_e]
            if sin_cfdi and (ppd_r or ppd_e):
                mc_mmp, sin_cfdi = self._conciliar_multimov(
                    sin_cfdi, ppd_r, ppd_e, usados_r, usados_e,
                    exigir_mismo_mes=False)
                conciliados.extend(mc_mmp)

            # Fase 7c — Comisiones bancarias (comisión + IVA => factura del banco)
            if sin_cfdi:
                mc_com, sin_cfdi = self._conciliar_comisiones_banco(
                    sin_cfdi, recibidos_adj, usados_r)
                conciliados.extend(mc_com)

        # ── Fase 8 — Agrupado por RFC (grupos configurados + automático) ─────
        # (a) Grupos multi-RFC configurados a mano (si los hay).
        # (b) Automático: suma de CFDIs del mismo RFC+mes ≈ suma de movimientos
        #     del mismo RFC (antes "Global Mensual").
        if not skip_suma and sin_cfdi and "grupos" in fases:
            if grupos:
                mc_g5, sin_cfdi = self._conciliar_grupos(
                    sin_cfdi, recibidos_adj, emitidos_adj,
                    usados_r, usados_e, grupos,
                )
                conciliados.extend(mc_g5)
            if sin_cfdi:
                pool_r_d = [c for c in recibidos_banc if c.uuid.lower() not in usados_r]
                pool_e_d = [c for c in emitidos_banc  if c.uuid.lower() not in usados_e]
                mc_gm, sin_cfdi = self._buscar_global_mensual(
                    sin_cfdi, pool_r_d, pool_e_d, usados_r, usados_e
                )
                conciliados.extend(mc_gm)

        # ── Fase 9 — Solo importe, MISMO MES (OPCIONAL, último recurso) ───────
        # Cruza lo que siga pendiente contra cualquier CFDI no usado del MISMO
        # MES cuyo importe coincida, sin validar identidad. Anti-duplicados vía
        # usados_r/usados_e. Todo match => revisar=True.
        if not skip_suma and sin_cfdi and "solo_importe" in fases:
            mc_f7, sin_cfdi = self._conciliar_solo_importe(
                sin_cfdi, recibidos_adj, emitidos_adj, usados_r, usados_e,
                bitacora=bitacora,
            )
            conciliados.extend(mc_f7)

        # Reincorporar los préstamos como pendientes (ya con su estatus propio),
        # DESPUÉS de todas las fases para que ninguna intente cruzarlos.
        sin_cfdi.extend(no_conciliables)

        cfdi_sin_banco_r = [c for c in self.recibidos
                            if c.uuid.lower() not in usados_r]
        cfdi_sin_banco_e = [c for c in self.emitidos
                            if c.uuid.lower() not in usados_e]

        # Marcar duplicados potenciales (mismo importe+mes en pool)
        freq_r = Counter(f"{c.total:.2f}|{c.mes}" for c in self.recibidos)
        freq_e = Counter(f"{c.total:.2f}|{c.mes}" for c in self.emitidos)
        for mc in conciliados:
            if mc.cfdi is None:
                continue
            if mc.tipo == "recibido":
                mc.revisar = mc.revisar or \
                    freq_r[f"{mc.cfdi.total:.2f}|{mc.cfdi.mes}"] > 1
            elif mc.tipo == "emitido":
                mc.revisar = mc.revisar or \
                    freq_e[f"{mc.cfdi.total:.2f}|{mc.cfdi.mes}"] > 1

        return conciliados, sin_cfdi, cfdi_sin_banco_r, cfdi_sin_banco_e

    # ── Internos ──────────────────────────────────────────────────────────────

    def _build_dicts(self, cfdis: list[CFDI], tercero_es_emisor: bool = True):
        """
        Construye tres diccionarios de búsqueda para una lista de CFDIs.

        tercero_es_emisor=True  → recibidos: empresa es receptora, el tercero es el emisor
        tercero_es_emisor=False → emitidos:  empresa es emisora,   el tercero es el receptor
        """
        d_rfc:  dict[str, list[str]] = {}   # "monto|mes|rfc" → [uuid, ...]
        d_mes:  dict[str, list[str]] = {}   # "monto|mes"     → [uuid, ...]
        d_nom:  dict[str, str]       = {}   # uuid_lower → nombre del tercero
        for c in cfdis:
            u = c.uuid
            if not u:
                continue
            key = f"{c.total:.2f}|{c.mes}"
            d_mes.setdefault(key, []).append(u)
            # Nivel 1: indexar ambos RFCs para máxima cobertura
            if c.rfc_emisor:
                d_rfc.setdefault(f"{key}|{c.rfc_emisor}", []).append(u)
            if c.rfc_receptor:
                d_rfc.setdefault(f"{key}|{c.rfc_receptor}", []).append(u)
            # Nombre del TERCERO (quien aparece en la descripción bancaria):
            #   Nómina         → empleado = receptor
            #   Recibidos      → proveedor = emisor
            #   Emitidos       → cliente   = receptor
            if c.tipo == "N":
                nombre_match = c.nombre_receptor or c.nombre_emisor
            elif tercero_es_emisor:
                nombre_match = c.nombre_emisor or c.nombre_receptor
            else:
                nombre_match = c.nombre_receptor or c.nombre_emisor
            d_nom[u.lower()] = nombre_match.upper()
        return d_rfc, d_mes, d_nom

    def _buscar(self,
                mov:            MovBanco,
                d_rfc:          dict[str, list[str]],
                d_mes:          dict[str, list[str]],
                d_nom:          dict[str, str],
                usados:         set[str],
                mismo_mes:      bool = False,
                solo_identidad: bool = False,
                dia_nomina:     dict[str, int] | None = None) -> tuple[Optional[str], int]:
        """
        Búsqueda bidireccional en 3 niveles.
        Retorna (uuid, nivel) o (None, 0).

        Nivel 1 — MONTO | MES | RFC exacto
        Nivel 2 — MONTO | MES + nombre extraído de la descripción (tolerante al orden)
        Nivel 3 — MONTO | MES + palabras del nombre CFDI en descripción completa
                  (cubre casos con nombres invertidos no detectados por nivel 2)
        Nivel 3.5 — MONTO | MES solo (sin identidad) — último recurso.

        mismo_mes=True restringe la búsqueda EXCLUSIVAMENTE al mes del propio
        movimiento — sin explorar meses adyacentes. Se usa en Fase 1 (CFDIs
        bancarizados: PUE nunca debe conciliar fuera de su mes). Cuando es
        False (default), conserva la búsqueda de años completo hacia
        adelante/atrás — el comportamiento que necesita Fase 4 para resolver
        PPD contra su complemento de pago, que puede caer en otro mes.

        solo_identidad=True omite el Nivel 3.5 (monto+mes sin identidad): el
        match SOLO procede si el RFC o el nombre coinciden. Se usa en Fase 6
        (PPD 99 con complemento, cualquier fecha) donde soltar la fecha exige
        una identidad fuerte para no producir falsos positivos.
        """
        monto_str = f"{mov.monto:.2f}"
        mes_base  = mov.mes or _mes_fecha(mov.fecha)
        rfc       = mov.rfc_en_descripcion()
        nom       = mov.nombre_en_descripcion()
        desc_up   = mov.descripcion.upper()

        # Filtro por DÍA solo para NÓMINA: un CFDI de nómina solo cruza si su
        # FechaPago cae en el MISMO día que el movimiento bancario. Evita que
        # una quincena tome el depósito de la otra (mismo monto, mismo mes,
        # distinto empleado). No afecta a proveedores/clientes (no están en
        # dia_nomina) ni a las fases sin dia_nomina (Fase 4/6 PPD).
        dia_mov = _dia_fecha(mov.fecha) if dia_nomina else 0

        def _dia_ok(uid: str) -> bool:
            if not dia_nomina:
                return True
            d = dia_nomina.get(uid.lower())
            return d is None or d == dia_mov

        if mismo_mes:
            meses_fwd = [mes_base] if mes_base else []
            meses_bwd = []
        else:
            meses_fwd = list(range(mes_base, 13))
            meses_bwd = list(range(mes_base - 1, 0, -1))

        # Nivel 1 — MONTO | MES | RFC  (exacto)
        if rfc:
            for orden in (meses_fwd, meses_bwd):
                for mes in orden:
                    key = f"{monto_str}|{mes}|{rfc}"
                    uid = self._primer_libre(d_rfc.get(key, []), usados, _dia_ok)
                    if uid:
                        return uid, 1

        # Nivel 1 tolerancia ±TOLERANCIA_MXN (solo cuando RFC identificado)
        if rfc:
            base_c = int(round(mov.monto * 100))
            tol_c  = int(round(TOLERANCIA_MXN * 100))
            for delta in range(1, tol_c + 1):
                for sign in (-1, 1):
                    adj_str = f"{(base_c + sign * delta) / 100:.2f}"
                    if adj_str == monto_str:
                        continue
                    for orden in (meses_fwd, meses_bwd):
                        for mes in orden:
                            uid = self._primer_libre(
                                d_rfc.get(f"{adj_str}|{mes}|{rfc}", []), usados, _dia_ok)
                            if uid:
                                return uid, 1

        # ── Nivel 1.1 — MONTO | MES + SOLO NOMBRE (sin depender del RFC) ─────
        # Para casos donde el RFC del tercero no está disponible/‌confiable y el
        # banco solo trae el nombre — típico de NÓMINA: el empleado viene con
        # RFC genérico o "NO disponible", así que el Nivel 1 (por RFC) no lo
        # encuentra. Aquí se cruza SOLO por el nombre (con el monto y mes):
        #   (a) nombre extraído de la descripción que coincide con el del CFDI
        #       (bidireccional, tolerante al orden) → match confiable (nivel 2)
        #   (b) el nombre del CFDI aparece textualmente en la descripción
        #       bancaria → match por nombre-en-texto (nivel 3, se marca revisar)
        for orden in (meses_fwd, meses_bwd):
            for mes in orden:
                key = f"{monto_str}|{mes}"
                for uuid in d_mes.get(key, []):
                    if uuid.lower() in usados:
                        continue
                    if not _dia_ok(uuid):
                        continue
                    nom_cfdi = d_nom.get(uuid.lower(), "")
                    if not nom_cfdi:
                        continue
                    if nom and _nombre_coincide(nom, nom_cfdi):
                        return uuid, 2
                    if _nombre_en_desc(desc_up, nom_cfdi):
                        return uuid, 3

        # ── Nivel 1.1b — NÓMINA con tolerancia de centavos (±$0.10) ──────────
        # El neto de nómina a veces difiere del depósito bancario por redondeo
        # (uno o dos centavos). SOLO para nómina (uuid en dia_nomina), mismo
        # día y con el nombre del empleado en la descripción, se acepta una
        # diferencia de hasta TOLERANCIA_NOMINA_CENT centavos. Muy específico
        # (nómina + día + nombre) → sin falsos positivos.
        if dia_nomina:
            base_c = int(round(mov.monto * 100))
            for delta in range(1, TOLERANCIA_NOMINA_CENT + 1):
                for sign in (-1, 1):
                    adj = f"{(base_c + sign * delta) / 100:.2f}"
                    for orden in (meses_fwd, meses_bwd):
                        for mes in orden:
                            for uuid in d_mes.get(f"{adj}|{mes}", []):
                                low = uuid.lower()
                                if low in usados or low not in dia_nomina:
                                    continue
                                if not _dia_ok(uuid):
                                    continue
                                nom_cfdi = d_nom.get(low, "")
                                if not nom_cfdi:
                                    continue
                                if (nom and _nombre_coincide(nom, nom_cfdi)) \
                                        or _nombre_en_desc(desc_up, nom_cfdi):
                                    return uuid, 2

        # Nivel 3.5 — MONTO | MES solo (sin identidad del tercero)
        # Fallback cuando la descripción bancaria no contiene RFC ni nombre.
        # Siempre se marcará "revisar" porque no hay señal de identidad.
        # solo_identidad=True lo omite (Fase 6 exige RFC/nombre).
        if not solo_identidad:
            for orden in (meses_fwd, meses_bwd):
                for mes in orden:
                    key = f"{monto_str}|{mes}"
                    uid = self._primer_libre(d_mes.get(key, []), usados, _dia_ok)
                    if uid:
                        return uid, 3  # nivel 3 → revisar=True en el llamador

        return None, 0

    @staticmethod
    def _primer_libre(uuids: list[str], usados: set[str],
                      ok=None) -> Optional[str]:
        for u in uuids:
            if u.lower() in usados:
                continue
            if ok is not None and not ok(u):
                continue
            return u
        return None

    # ── Nivel 4: búsqueda por suma de N CFDIs ────────────────────────────────

    _MAX_COMBO = 5   # máximo CFDIs en una combinación (limita explosión combinatoria)

    def _buscar_suma(self,
                     mov:    MovBanco,
                     pool:   list[CFDI],
                     usados: set[str]) -> list[CFDI]:
        """
        Nivel 4 optimizado: busca 2..._MAX_COMBO CFDIs que sumen mov.monto usando
        un enfoque de subset-sum dinámico (DP) en centavos, con las mismas
        prioridades de identidad que antes (RFC -> nombre -> nombre en desc).

        Restringido al MISMO mes que el movimiento — una suma de CFDIs
        bancarizados (Fase 3) no debe cruzar meses, igual que Fase 1.
        """
        target = int(round(mov.monto * 100))
        if target <= 0:
            return []

        mes_mov = mov.mes or _mes_fecha(mov.fecha)

        # Candidatos: no usados, estrictamente menores al target, mismo mes
        cands = [c for c in pool
                 if c.uuid and c.uuid.lower() not in usados
                 and int(round(c.total * 100)) < target
                 and c.mes == mes_mov]

        if len(cands) < 2:
            return []

        # Determinar qué RFC/nombre identifica al TERCERO según dirección del movimiento:
        #   retiro   → empresa paga a proveedor → tercero = emisor del CFDI recibido
        #   depósito → empresa cobra a cliente  → tercero = receptor del CFDI emitido
        es_retiro = mov.retiro > 0

        grupos: dict[str, list[CFDI]] = {}
        for c in cands:
            if es_retiro:
                rfc_key = (c.rfc_emisor or "").upper().strip()
                nom_key = (c.nombre_emisor or c.nombre_receptor or "").upper().strip()
            else:
                rfc_key = (c.rfc_receptor or "").upper().strip()
                nom_key = (c.nombre_receptor or c.nombre_emisor or "").upper().strip()
            if c.tipo == "N":
                # Nómina: el tercero es el empleado (receptor), no la empresa.
                nom_key = (c.nombre_receptor or c.nombre_emisor or "").upper().strip()
            # Agrupa por RFC si lo hay; si el RFC del tercero viene vacío o
            # "NO disponible" (típico de nómina o CFDIs sin RFC del receptor),
            # usa el NOMBRE como clave para que igual pueda cruzar por
            # nombre + mes + valores que suman el monto.
            clave = rfc_key or (f"NOM:{nom_key}" if nom_key else "")
            if clave:
                grupos.setdefault(clave, []).append(c)

        def _dp_find_combo(grupo: list[CFDI]) -> list[CFDI]:
            """DP subset-sum que devuelve la primera combinación encontrada."""
            vals = [int(round(c.total * 100)) for c in grupo]

            # Cota inferior: si los 2 menores ya superan el objetivo,
            # ningún par puede sumar exactamente target.
            # Tolerancia +/-TOLERANCIA_MXN, la MISMA que ya usa la Fase 1: una
            # diferencia de centavos (comision, redondeo) no debe apagar la fase.
            _tol = int(round(TOLERANCIA_MXN * 100))
            sorted_asc = sorted(vals)
            if len(sorted_asc) >= 2 and sorted_asc[0] + sorted_asc[1] > target + _tol:
                return []
            # Cota superior: si los _MAX_COMBO mayores juntos no llegan al target,
            # ninguna combinación válida existe.
            sorted_desc = sorted(vals, reverse=True)
            if sum(sorted_desc[:min(len(sorted_desc), self._MAX_COMBO)]) < target - _tol:
                return []

            # Reutiliza el buscador COMPARTIDO ya corregido en vez de una copia
            # local del DP: la copia descartaba caminos validos cuando la suma ya
            # se habia alcanzado por otra via (mismo bug que tenia la Fase 7).
            idxs = _subset_sum_idx(vals, target, self._MAX_COMBO, _tol)
            if len(idxs) >= 2:
                return [grupo[i] for i in idxs]
            return []

        def _nombre_grupo(grupo: list[CFDI]) -> str:
            c = grupo[0]
            if c.tipo == "N":
                return (c.nombre_receptor or c.nombre_emisor or "").upper()
            if es_retiro:
                return (c.nombre_emisor or c.nombre_receptor or "").upper()
            return (c.nombre_receptor or c.nombre_emisor or "").upper()

        rfc_desc = mov.rfc_en_descripcion()
        nom_desc = mov.nombre_en_descripcion()
        desc_up  = mov.descripcion.upper()

        # Prioridad 1: RFC exacto
        if rfc_desc and rfc_desc in grupos:
            res = _dp_find_combo(grupos[rfc_desc])
            if res:
                return res

        # Prioridad 2: nombre extraído coincide con nombre del grupo
        if nom_desc:
            for rfc_key, grupo in grupos.items():
                if _nombre_coincide(nom_desc, _nombre_grupo(grupo)):
                    res = _dp_find_combo(grupo)
                    if res:
                        return res

        # Prioridad 3: nombre aparece en la descripción completa
        for rfc_key, grupo in grupos.items():
            if _nombre_en_desc(desc_up, _nombre_grupo(grupo)):
                res = _dp_find_combo(grupo)
                if res:
                    return res

        return []

    @staticmethod
    def _conciliar_grupos(
        sin_cfdi: list[MovBanco],
        pool_r:   list["CFDI"],   # type: ignore[name-defined]
        pool_e:   list["CFDI"],   # type: ignore[name-defined]
        usados_r: set[str],
        usados_e: set[str],
        grupos:   list,           # list[GrupoConciliacion]
    ) -> tuple[list["MovConciliado"], list[MovBanco]]:
        """
        Fase 5 — Conciliación Multi-RFC (Grupos).

        Para cada movimiento sin conciliar prueba si la suma de todos los
        CFDIs disponibles de los RFCs del grupo coincide con el monto del
        movimiento (dentro de la tolerancia del grupo).

        Depósito → busca en emitidos (empresa cobró a esos clientes).
        Retiro   → busca en recibidos (empresa pagó a esos proveedores).
        """
        nuevos:     list[MovConciliado] = []
        pendientes: list[MovBanco]      = []

        for mov in sin_cfdi:
            matched = False
            for grupo in grupos:
                if not grupo.activo:
                    continue
                tol = getattr(grupo, "tolerancia", TOLERANCIA_MXN)

                rfcs_prov = {r.strip().upper()
                             for r in getattr(grupo, "rfcs_proveedores", grupo.rfcs)}
                rfcs_cli  = {r.strip().upper()
                             for r in getattr(grupo, "rfcs_clientes",    [])}

                if mov.deposito > 0 and rfcs_cli:
                    cands = [c for c in pool_e
                             if (c.rfc_receptor or "").upper() in rfcs_cli
                             and c.uuid.lower() not in usados_e]
                    usados_set = usados_e
                    tipo       = "emitido"
                elif mov.retiro > 0 and rfcs_prov:
                    cands = [c for c in pool_r
                             if (c.rfc_emisor or "").upper() in rfcs_prov
                             and c.uuid.lower() not in usados_r]
                    usados_set = usados_r
                    tipo       = "recibido"
                else:
                    continue

                if not cands:
                    continue

                # Elegir el SUBCONJUNTO de facturas del grupo que cuadra con el
                # movimiento (antes exigía que TODAS sumaran el monto: "todo o
                # nada"; si sobraba una factura, no cruzaba). Se prueba en orden:
                #   1) todas las facturas (comportamiento previo, compatibilidad)
                #   2) una sola factura dentro de la tolerancia
                #   3) un subconjunto (≥2) que sume el movimiento (exacto o ±tol)
                # Las facturas NO elegidas quedan libres para otros movimientos.
                target = int(round(mov.monto * 100))
                tol_c  = int(round(tol * 100))
                vals   = [int(round(c.total * 100)) for c in cands]

                elegidos_idx: list[int] | None = None
                if abs(sum(vals) - target) <= tol_c:
                    elegidos_idx = list(range(len(cands)))
                if elegidos_idx is None:
                    for _i, _v in enumerate(vals):
                        if abs(_v - target) <= tol_c:
                            elegidos_idx = [_i]
                            break
                if elegidos_idx is None:
                    for _d in range(0, tol_c + 1):
                        _targets = {target} if _d == 0 else {target - _d, target + _d}
                        for _t in _targets:
                            if _t <= 0:
                                continue
                            _sub = _subset_sum_idx(vals, _t, 12)
                            if _sub:
                                elegidos_idx = _sub
                                break
                        if elegidos_idx is not None:
                            break
                if elegidos_idx is None:
                    continue

                elegidos = [cands[i] for i in elegidos_idx]

                # ¡Match! — registrar solo los CFDIs elegidos
                for c in elegidos:
                    usados_set.add(c.uuid.lower())

                Conciliador._llenar_mov(mov, elegidos[0], tipo=tipo)
                mov.estatus   = f"Grupo RFC: {grupo.nombre}"
                mov.uuid_suma = ",".join(c.uuid for c in elegidos[1:])

                nuevos.append(MovConciliado(
                    mov, elegidos[0], tipo,
                    revisar   = True,
                    nivel     = 6,
                    cfdi_suma = elegidos[1:],
                    tipo_match= f"Grupo: {grupo.nombre}",
                ))
                matched = True
                break   # un movimiento → primer grupo que cuadre

            if not matched:
                pendientes.append(mov)

        return nuevos, pendientes

    def _conciliar_ppd(
        self,
        sin_cfdi:      list[MovBanco],
        recibidos_adj: list[CFDI],
        emitidos_adj:  list[CFDI],
        usados_r:      set[str],
        usados_e:      set[str],
        *,
        con_complemento: bool,
        tipo_match:      str,
        nivel:           int,
        solo_identidad:  bool,
        modo:            str = "ambos",
    ) -> tuple[list["MovConciliado"], list[MovBanco]]:
        """
        Cruce de facturas PPD (forma_pago 99) contra el banco, por MONTO +
        RFC/nombre, sin restricción de mes (fecha pasada o futura). Usado en
        dos fases con distinta prioridad y semántica:

        Fase 4 (con_complemento=True, solo_identidad=True) — el pago legítimo:
          PPD 99 que SÍ tiene complemento de pago tipo P. Exige identidad
          (RFC/nombre) porque suelta la fecha; es el caso contable correcto y
          va primero.
        Fase 6 (con_complemento=False, solo_identidad=True) — la anomalía:
          PPD 99 SIN complemento (hay movimiento en banco pero falta emitir el
          complemento). Se marca conciliado pero pendiente de regularizar.
          Cruza mes pero exige RFC o nombre del tercero (nunca por importe solo).

        Ambas marcan revisar=True.
        """
        con_comp = _uuids_con_complemento(self.recibidos + self.emitidos)

        def _elegible(c: CFDI, usados: set[str]) -> bool:
            if c.metodo_pago != "PPD" or c.forma_pago != "99":
                return False
            if c.uuid.lower() in usados:
                return False
            return (c.uuid.lower() in con_comp) == con_complemento

        ppd_r = [c for c in recibidos_adj if _elegible(c, usados_r)]
        ppd_e = [c for c in emitidos_adj  if _elegible(c, usados_e)]

        if not ppd_r and not ppd_e:
            return [], sin_cfdi

        dr, dm_r, dn_r = self._build_dicts(ppd_r, tercero_es_emisor=True)
        de, dm_e, dn_e = self._build_dicts(ppd_e, tercero_es_emisor=False)
        por_r = {c.uuid.lower(): c for c in ppd_r}
        por_e = {c.uuid.lower(): c for c in ppd_e}

        # modo restringe la identidad exigida sin tocar _buscar: pasando un dict
        # vacío se "apaga" ese nivel de match. "rfc" → solo Nivel 1 (RFC en la
        # descripción); "nombre" → solo Niveles 2/3 (nombre); "ambos" → los dos.
        use_rfc = modo in ("ambos", "rfc")
        use_nom = modo in ("ambos", "nombre")
        _dr = dr if use_rfc else {}
        _de = de if use_rfc else {}
        _dnr = dn_r if use_nom else {}
        _dne = dn_e if use_nom else {}

        nuevos:     list[MovConciliado] = []
        pendientes: list[MovBanco]      = []

        for mov in sin_cfdi:
            matched = False
            if mov.retiro > 0 and ppd_r:
                uid, _ = self._buscar(mov, _dr, dm_r, _dnr, usados_r,
                                      mismo_mes=False, solo_identidad=solo_identidad)
                if uid:
                    cfdi = por_r.get(uid.lower())
                    if cfdi:
                        usados_r.add(uid.lower())
                        self._llenar_mov(mov, cfdi, tipo="recibido")
                        nuevos.append(MovConciliado(
                            mov, cfdi, "recibido", revisar=True, nivel=nivel,
                            tipo_match=tipo_match,
                        ))
                        matched = True
            elif mov.deposito > 0 and ppd_e:
                uid, _ = self._buscar(mov, _de, dm_e, _dne, usados_e,
                                      mismo_mes=False, solo_identidad=solo_identidad)
                if uid:
                    cfdi = por_e.get(uid.lower())
                    if cfdi:
                        usados_e.add(uid.lower())
                        self._llenar_mov(mov, cfdi, tipo="emitido")
                        nuevos.append(MovConciliado(
                            mov, cfdi, "emitido", revisar=True, nivel=nivel,
                            tipo_match=tipo_match,
                        ))
                        matched = True
            if not matched:
                pendientes.append(mov)

        return nuevos, pendientes

    def _conciliar_multimov(
        self,
        sin_cfdi:       list[MovBanco],
        recibidos_banc: list[CFDI],
        emitidos_banc:  list[CFDI],
        usados_r:       set[str],
        usados_e:       set[str],
        *,
        exigir_mismo_mes: bool = True,
        tipo_match:       str  = "Multi-Movimiento",
    ) -> tuple[list["MovConciliado"], list[MovBanco]]:
        """
        Fase 2 — Varios movimientos bancarios = UN mismo CFDI (N:1).

        Un CFDI grande liquidado en varias parcialidades bancarias del mismo
        RFC/nombre y mismo mes. Hace subset-sum sobre los MOVIMIENTOS (2..6)
        cuyos importes suman el total del CFDI, exigiendo que la identidad
        (RFC o nombre del tercero) aparezca en la descripción de cada movimiento.

        retiro   → CFDIs recibidos (empresa pagó al proveedor en parcialidades)
        depósito → CFDIs emitidos  (cliente pagó una factura en parcialidades)

        exigir_mismo_mes : si False, las parcialidades pueden caer en meses
        distintos al de la factura. Solo se relaja para PPD 99 (parcialidades),
        donde el cruce entre meses ya es la conducta contable correcta —
        NUNCA para facturas bancarizadas (PUE), que conservan el candado de mes.
        """
        idx_usados: set[int] = set()   # índices de sin_cfdi consumidos aquí
        nuevos: list[MovConciliado] = []

        retiros   = [(i, m) for i, m in enumerate(sin_cfdi) if m.retiro   > 0]
        depositos = [(i, m) for i, m in enumerate(sin_cfdi) if m.deposito > 0]

        def _procesar(cfdis: list[CFDI], movs_dir, usados: set[str], tipo: str):
            for c in sorted(cfdis, key=lambda x: x.total, reverse=True):
                if c.uuid.lower() in usados:
                    continue
                target = int(round(c.total * 100))
                if target <= 0:
                    continue
                if tipo == "recibido":
                    rfc_c = c.rfc_emisor
                    nom_c = c.nombre_emisor or c.nombre_receptor
                else:
                    rfc_c = c.rfc_receptor
                    nom_c = c.nombre_receptor or c.nombre_emisor
                if c.tipo == "N":
                    nom_c = c.nombre_receptor or c.nombre_emisor
                # Candidatos: mismo mes, no usados aquí, identidad coincide,
                # importe individual < total del CFDI.
                cands = []
                for i, m in movs_dir:
                    if i in idx_usados:
                        continue
                    if exigir_mismo_mes and (m.mes or _mes_fecha(m.fecha)) != c.mes:
                        continue
                    mc = int(round(m.monto * 100))
                    if mc <= 0 or mc >= target:
                        continue
                    if not _mov_coincide_identidad(m, rfc_c, nom_c):
                        continue
                    cands.append((i, m, mc))
                if len(cands) < 2:
                    continue
                # Misma tolerancia +/-TOLERANCIA_MXN que la Fase 1.
                combo = _subset_sum_idx([mc for _, _, mc in cands], target, 6,
                                        int(round(TOLERANCIA_MXN * 100)))
                if len(combo) < 2:
                    continue
                usados.add(c.uuid.lower())
                for k in combo:
                    i, m, _ = cands[k]
                    idx_usados.add(i)
                    self._llenar_mov(m, c, tipo=tipo)
                    m.estatus = "Conciliado Multi-Movimiento"
                    nuevos.append(MovConciliado(
                        m, c, tipo, revisar=True, nivel=2,
                        tipo_match=tipo_match,
                    ))

        _procesar([c for c in recibidos_banc if c.uuid.lower() not in usados_r],
                  retiros, usados_r, "recibido")
        _procesar([c for c in emitidos_banc if c.uuid.lower() not in usados_e],
                  depositos, usados_e, "emitido")

        pendientes = [m for i, m in enumerate(sin_cfdi) if i not in idx_usados]
        return nuevos, pendientes

    def _conciliar_comisiones_banco(
        self,
        sin_cfdi:      list[MovBanco],
        recibidos_adj: list[CFDI],
        usados_r:      set[str],
    ) -> tuple[list["MovConciliado"], list[MovBanco]]:
        """
        Comisiones bancarias: la comisión y su IVA vienen como DOS movimientos
        separados en el estado de cuenta, pero la factura del banco los trae
        juntos (total = comisión + IVA). Agrupa (subset-sum) los movimientos de
        comisión pendientes que suman el total de una factura emitida por una
        institución financiera.

        No exige el nombre del banco en la descripción (el movimiento dice
        "COMISION POR TRANSFERENCIA", no nombra al banco): la identidad la dan
        dos condiciones juntas — el CFDI lo emite un BANCO y los movimientos son
        de COMISIÓN. Cruza sin candado de mes (la factura del banco puede
        timbrarse en un mes distinto al del cargo). Marca revisar=True.
        """
        banco_cfdis = [c for c in recibidos_adj
                       if c.tipo == "I" and c.uuid.lower() not in usados_r
                       and _es_emisor_banco(c)]
        comis = [m for m in sin_cfdi if m.retiro > 0 and _es_mov_comision(m)]
        if not banco_cfdis or len(comis) < 2:
            return [], sin_cfdi

        consumidos: set[int] = set()
        nuevos: list[MovConciliado] = []
        for c in sorted(banco_cfdis, key=lambda x: x.total, reverse=True):
            target = int(round(c.total * 100))
            if target <= 0:
                continue
            cands = [(m, int(round(m.monto * 100))) for m in comis
                     if id(m) not in consumidos]
            if len(cands) < 2:
                continue
            combo = _subset_sum_idx([mc for _, mc in cands], target, 8)
            if len(combo) < 2:
                continue
            usados_r.add(c.uuid.lower())
            for k in combo:
                m, _ = cands[k]
                consumidos.add(id(m))
                self._llenar_mov(m, c, tipo="recibido")
                m.estatus = "Conciliado Comisión"
                nuevos.append(MovConciliado(
                    m, c, "recibido", revisar=True, nivel=6,
                    tipo_match="Comisión Bancaria",
                ))
        if not consumidos:
            return [], sin_cfdi
        pendientes = [m for m in sin_cfdi if id(m) not in consumidos]
        return nuevos, pendientes

    def _conciliar_rep(
        self,
        sin_cfdi:      list[MovBanco],
        recibidos_adj: list[CFDI],
        emitidos_adj:  list[CFDI],
        usados_r:      set[str],
        usados_e:      set[str],
        modo:          str = "ambos",
    ) -> tuple[list["MovConciliado"], list[MovBanco]]:
        """
        Fase 4 — PPD con complemento de pago (Recepción de Pago, tipo P).

        Verifica PRIMERO el complemento (REP): casa su TOTAL PAGADO (MontoPagado)
        + RFC/nombre contra el movimiento bancario exigiendo que la **FechaPago
        del REP** caiga en el **MISMO MES** del movimiento bancario **O** a
        **±TOLERANCIA_DIAS_REP días** de él. Al cruzar vincula el REP y TODAS
        las facturas que liquida (cfdi_facturas_pagadas).

        El criterio era ±1 día, pero en la práctica el complemento se timbra
        días después del pago, así que muchos REP legítimos quedaban fuera.
        El mes es el candado de periodo; el OR de ±5 días cubre el pago de fin
        de mes que el banco refleja en los primeros días del mes siguiente.
        La identidad (RFC/nombre) + el MontoPagado exacto son los que dan la
        precisión del cruce.

        Banco → REP → Facturas. Guard anti-doble-conteo: solo REPs cuyas facturas
        están TODAS sin usar (si alguna ya cruzó por otra vía, se omite el REP).
        """
        def _elegibles(pool: list[CFDI], usados: set[str]) -> list[CFDI]:
            out = []
            for c in pool:
                if c.tipo != "P" or c.uuid.lower() in usados or not c.fecha_pago:
                    continue
                if all(u.lower() not in usados for u in c.cfdi_facturas_pagadas):
                    out.append(c)
            return out

        reps_r = _elegibles(recibidos_adj, usados_r)
        reps_e = _elegibles(emitidos_adj,  usados_e)
        if not reps_r and not reps_e:
            return [], sin_cfdi

        tol_c = int(round(TOLERANCIA_MXN * 100))

        def _match_rep(mov: MovBanco, reps: list[CFDI], usados: set[str],
                       tipo: str) -> Optional[CFDI]:
            monto_c = int(round(mov.monto * 100))
            mes_mov = mov.mes or _mes_fecha(mov.fecha)
            for rep in reps:
                if rep.uuid.lower() in usados:
                    continue
                if abs(int(round(rep.total * 100)) - monto_c) > tol_c:
                    continue
                if tipo == "recibido":
                    rfc_r, nom_r = rep.rfc_emisor, rep.nombre_emisor
                else:
                    rfc_r, nom_r = rep.rfc_receptor, rep.nombre_receptor
                # modo="rfc" exige RFC; modo="nombre" exige nombre; "ambos" = cualquiera.
                _rfc_id = rfc_r if modo in ("ambos", "rfc")    else ""
                _nom_id = nom_r if modo in ("ambos", "nombre") else ""
                if not _mov_coincide_identidad(mov, _rfc_id, _nom_id):
                    continue
                # Candado de periodo: mismo mes bancario O desfase ≤ TOLERANCIA_DIAS_REP.
                # El OR cubre el pago de fin de mes que el banco refleja en los
                # primeros días del mes siguiente (30/jun vs 01/jul).
                mes_rep = _mes_fecha(rep.fecha_pago)
                mismo_mes = bool(mes_mov and mes_rep and mes_mov == mes_rep)
                if not mismo_mes:
                    dd = _dias_entre(mov.fecha, rep.fecha_pago)
                    if dd is None or dd > TOLERANCIA_DIAS_REP:
                        continue
                return rep
            return None

        nuevos:     list[MovConciliado] = []
        pendientes: list[MovBanco]      = []

        for mov in sin_cfdi:
            rep = None
            tipo = ""
            if mov.retiro > 0 and reps_r:
                rep, tipo = _match_rep(mov, reps_r, usados_r, "recibido"), "recibido"
            elif mov.deposito > 0 and reps_e:
                rep, tipo = _match_rep(mov, reps_e, usados_e, "emitido"), "emitido"

            if rep is None:
                pendientes.append(mov)
                continue

            usados = usados_r if tipo == "recibido" else usados_e
            usados.add(rep.uuid.lower())
            for u in rep.cfdi_facturas_pagadas:
                usados.add(u.lower())

            # Trazar por separado el complemento (REP) y las facturas (CFDI)
            # que liquida: mov.uuid = factura(s) tipo I, mov.uuid_rep = REP.
            pool_fact = recibidos_adj if tipo == "recibido" else emitidos_adj
            por_uuid  = {c.uuid.lower(): c for c in pool_fact}
            facturas  = [por_uuid[u.lower()] for u in rep.cfdi_facturas_pagadas
                         if u.lower() in por_uuid]
            if facturas:
                self._llenar_mov(mov, facturas[0], tipo=tipo)
                mov.uuid_suma = ",".join(c.uuid for c in facturas[1:])
            else:
                # La factura no vino en el lote: conservar los UUIDs del REP.
                self._llenar_mov(mov, rep, tipo=tipo)
                mov.uuid_suma = ",".join(rep.cfdi_facturas_pagadas)
            mov.uuid_rep = rep.uuid
            nuevos.append(MovConciliado(
                mov, rep, tipo, revisar=True, nivel=5,
                tipo_match="PPD c/ Complemento",
            ))

        # ── Segunda pasada — REP liquidado por VARIOS movimientos (N:1) ──────
        # El MontoPagado del REP se cubre con varias parcialidades bancarias
        # (mismo tercero). La pasada 1:1 de arriba no las ve porque ninguna
        # iguala el total sola. Subset-sum de los movimientos PENDIENTES cuya
        # identidad (RFC/nombre, según `modo`) coincide con el emisor/receptor
        # del REP y suman su MontoPagado, dentro del mismo candado de periodo.
        # Caso real: REP de $27,016.40 = $20,346.40 + $6,670.00, ambos con el
        # mismo BENEFICIARIO en la descripción (MEXLIVING MBA1701182QA).
        consumidos: set[int] = set()

        def _match_rep_multi(reps: list[CFDI], usados: set[str], tipo: str) -> None:
            for rep in reps:
                if rep.uuid.lower() in usados:
                    continue
                if any(u.lower() in usados for u in rep.cfdi_facturas_pagadas):
                    continue
                target = int(round(rep.total * 100))
                if target <= 0:
                    continue
                if tipo == "recibido":
                    rfc_r, nom_r = rep.rfc_emisor, rep.nombre_emisor
                else:
                    rfc_r, nom_r = rep.rfc_receptor, rep.nombre_receptor
                _rfc_id = rfc_r if modo in ("ambos", "rfc")    else ""
                _nom_id = nom_r if modo in ("ambos", "nombre") else ""
                mes_rep = _mes_fecha(rep.fecha_pago)
                cands: list[tuple[MovBanco, int]] = []
                for m in pendientes:
                    if id(m) in consumidos:
                        continue
                    if tipo == "recibido" and not m.retiro > 0:
                        continue
                    if tipo == "emitido" and not m.deposito > 0:
                        continue
                    mc = int(round(m.monto * 100))
                    if mc <= 0 or mc >= target:
                        continue
                    if not _mov_coincide_identidad(m, _rfc_id, _nom_id):
                        continue
                    mes_mov = m.mes or _mes_fecha(m.fecha)
                    mismo_mes = bool(mes_mov and mes_rep and mes_mov == mes_rep)
                    if not mismo_mes:
                        dd = _dias_entre(m.fecha, rep.fecha_pago)
                        if dd is None or dd > TOLERANCIA_DIAS_REP:
                            continue
                    cands.append((m, mc))
                if len(cands) < 2:
                    continue
                combo = _subset_sum_idx([mc for _, mc in cands], target, 6)
                if len(combo) < 2:
                    continue
                usados.add(rep.uuid.lower())
                for u in rep.cfdi_facturas_pagadas:
                    usados.add(u.lower())
                pool_fact = recibidos_adj if tipo == "recibido" else emitidos_adj
                por_uuid  = {c.uuid.lower(): c for c in pool_fact}
                facturas  = [por_uuid[u.lower()] for u in rep.cfdi_facturas_pagadas
                             if u.lower() in por_uuid]
                for k in combo:
                    m, _ = cands[k]
                    consumidos.add(id(m))
                    if facturas:
                        self._llenar_mov(m, facturas[0], tipo=tipo)
                        m.uuid_suma = ",".join(c.uuid for c in facturas[1:])
                    else:
                        self._llenar_mov(m, rep, tipo=tipo)
                        m.uuid_suma = ",".join(rep.cfdi_facturas_pagadas)
                    m.uuid_rep = rep.uuid
                    m.estatus  = "Conciliado Multi-Movimiento"
                    nuevos.append(MovConciliado(
                        m, rep, tipo, revisar=True, nivel=5,
                        tipo_match="PPD c/ Complemento",
                    ))

        _match_rep_multi(reps_r, usados_r, "recibido")
        _match_rep_multi(reps_e, usados_e, "emitido")
        if consumidos:
            pendientes = [m for m in pendientes if id(m) not in consumidos]

        return nuevos, pendientes

    @staticmethod
    def _buscar_global_mensual(
        sin_cfdi: list[MovBanco],
        pool_r:   list[CFDI],
        pool_e:   list[CFDI],
        usados_r: set[str],
        usados_e: set[str],
    ) -> tuple[list["MovConciliado"], list[MovBanco]]:
        """
        Fase 5 (automático) — Agrupado por RFC + mes (N:M).

        Para cada clave (RFC_banco, mes), compara:
          retiros  pendientes vs CFDIs recibidos disponibles
          depósitos pendientes vs CFDIs emitidos  disponibles

        Si la suma del grupo bancario ≈ suma CFDI (±TOLERANCIA_MXN),
        concilia todo el grupo con tipo_match='Global Mensual'.
        """
        from collections import defaultdict

        nuevos:     list[MovConciliado] = []
        usados_idx: set[int]            = set()

        # ── Agrupar movimientos pendientes por (rfc_descripcion, mes) ────────
        grupos_ret: dict[tuple, list[int]] = defaultdict(list)
        grupos_dep: dict[tuple, list[int]] = defaultdict(list)

        for i, mov in enumerate(sin_cfdi):
            rfc = mov.rfc_en_descripcion()
            mes = mov.mes or _mes_fecha(mov.fecha)
            if not rfc or not mes:
                continue
            if mov.retiro > 0:
                grupos_ret[(rfc, mes)].append(i)
            elif mov.deposito > 0:
                grupos_dep[(rfc, mes)].append(i)

        # ── Agrupar CFDIs disponibles por (rfc_tercero, mes) ─────────────────
        def _cfdi_por_clave(pool: list[CFDI], usados: set[str],
                            es_recibido: bool) -> dict:
            grp: dict[tuple, list[CFDI]] = defaultdict(list)
            for c in pool:
                if c.uuid.lower() in usados:
                    continue
                rfc_t = c.rfc_emisor if es_recibido else c.rfc_receptor
                if rfc_t and c.mes:
                    grp[(rfc_t.upper(), c.mes)].append(c)
            return grp

        cfdi_rec = _cfdi_por_clave(pool_r, usados_r, True)
        cfdi_emi = _cfdi_por_clave(pool_e, usados_e, False)

        # ── Conciliar grupo cuando sumas cuadran ──────────────────────────────
        def _procesar_grupo(
            mov_idxs:    list[int],
            cfdis_grp:   list[CFDI],
            tipo:        str,
            usados_cfdi: set[str],
        ) -> None:
            sum_movs = round(sum(sin_cfdi[i].monto for i in mov_idxs), 2)
            sum_cfd  = round(sum(c.total for c in cfdis_grp), 2)
            if abs(sum_movs - sum_cfd) > TOLERANCIA_MXN:
                return

            # Ordenar descendente para mejor pareamiento
            movs_ord  = sorted(mov_idxs,  key=lambda i: sin_cfdi[i].monto, reverse=True)
            cfdis_ord = sorted(cfdis_grp, key=lambda c: c.total,            reverse=True)

            for rank, idx in enumerate(movs_ord):
                mov    = sin_cfdi[idx]
                cfdi_p = cfdis_ord[rank] if rank < len(cfdis_ord) else cfdis_ord[-1]

                # CFDIs sobrantes van en cfdi_suma del último movimiento
                extra: list[CFDI] = (
                    cfdis_ord[len(movs_ord):]
                    if rank == len(movs_ord) - 1 and len(cfdis_ord) > len(movs_ord)
                    else []
                )

                Conciliador._llenar_mov(mov, cfdi_p, tipo=tipo)
                mov.estatus   = "Global Mensual"
                mov.uuid_suma = ",".join(c.uuid for c in extra)

                usados_cfdi.add(cfdi_p.uuid.lower())
                for c in extra:
                    usados_cfdi.add(c.uuid.lower())
                usados_idx.add(idx)

                nuevos.append(MovConciliado(
                    mov, cfdi_p, tipo,
                    revisar=False, nivel=2,
                    cfdi_suma=extra,
                    tipo_match="Global Mensual",
                ))

        for key, idxs in grupos_ret.items():
            if key in cfdi_rec:
                _procesar_grupo(idxs, cfdi_rec[key], "recibido", usados_r)

        for key, idxs in grupos_dep.items():
            if key in cfdi_emi:
                _procesar_grupo(idxs, cfdi_emi[key], "emitido", usados_e)

        pendientes = [mov for i, mov in enumerate(sin_cfdi) if i not in usados_idx]
        return nuevos, pendientes

    @staticmethod
    def _llenar_mov(mov: MovBanco, cfdi: CFDI, tipo: str = "") -> None:
        mov.uuid       = cfdi.uuid
        mov.fecha_cfdi = cfdi.fecha_emision[:10]
        mov.serie_cfdi = cfdi.serie
        mov.folio_cfdi = cfdi.folio
        mov.estatus    = "Conciliado"

        if tipo == "recibido":
            if cfdi.tipo == "N":
                # NÓMINA: la empresa pagó al empleado → beneficiario = RECEPTOR (empleado)
                mov.rfc             = cfdi.rfc_receptor
                mov.nombre          = cfdi.nombre_receptor
                mov.fecha_pago_cfdi = cfdi.fecha_pago   # DD/MM/YYYY
            else:
                # RETIRO normal: empresa pagó a proveedor → beneficiario = EMISOR del CFDI
                mov.rfc    = cfdi.rfc_emisor
                mov.nombre = cfdi.nombre_emisor
        elif tipo == "emitido":
            # DEPÓSITO: la empresa cobró → pagador es el RECEPTOR del CFDI
            mov.rfc    = cfdi.rfc_receptor
            mov.nombre = cfdi.nombre_receptor
        else:
            # Fallback genérico
            mov.rfc    = cfdi.rfc_emisor or cfdi.rfc_receptor
            mov.nombre = cfdi.nombre_emisor or cfdi.nombre_receptor

    def _conciliar_solo_importe(
        self,
        sin_cfdi:      list[MovBanco],
        recibidos_adj: list[CFDI],
        emitidos_adj:  list[CFDI],
        usados_r:      set[str],
        usados_e:      set[str],
        bitacora:      list | None = None,
    ) -> tuple[list["MovConciliado"], list[MovBanco]]:
        """
        Fase 7 (OPCIONAL) — último recurso: cruza un movimiento que quedó sin
        conciliar contra CUALQUIER CFDI disponible cuyo IMPORTE coincida
        exactamente Y sea del MISMO MES que el movimiento, SIN validar identidad
        (RFC/nombre).

        Anti-duplicados: solo considera CFDIs que NO estén ya en usados_r/usados_e
        (la bitácora de comprobantes consumidos por las fases previas). Cada
        cruce marca el CFDI como usado de inmediato y anota una entrada en
        `bitacora`, así el mismo comprobante nunca concilia dos movimientos.

        Como no hay señal de tercero, TODO match aquí se marca revisar=True
        (tipo_match="Solo Importe", nivel=8).

        Un retiro busca en recibidos (la empresa pagó); un depósito en emitidos
        (la empresa cobró). Se excluyen los tipo P (complementos de pago: los
        resuelve la Fase 4) — su total no representa una factura conciliable.
        """
        def _pool(cfdis: list[CFDI], usados: set[str]) -> dict[tuple, list[CFDI]]:
            # (importe_en_centavos, mes) → [CFDIs disponibles]
            idx: dict[tuple, list[CFDI]] = {}
            for c in cfdis:
                if not c.uuid or c.uuid.lower() in usados or c.tipo == "P":
                    continue
                cents = int(round(c.total * 100))
                if cents <= 0:
                    continue
                idx.setdefault((cents, c.mes), []).append(c)
            return idx

        pool_r = _pool(recibidos_adj, usados_r)
        pool_e = _pool(emitidos_adj,  usados_e)

        nuevos:     list[MovConciliado] = []
        pendientes: list[MovBanco]      = []

        def _tomar(idx: dict[tuple, list[CFDI]], usados: set[str],
                   cents: int, mes: int) -> Optional[CFDI]:
            for c in idx.get((cents, mes), []):
                if c.uuid.lower() not in usados:
                    return c
            return None

        for mov in sin_cfdi:
            cents  = int(round(mov.monto * 100))
            mes_mv = mov.mes or _mes_fecha(mov.fecha)
            cfdi   = None
            tipo   = ""
            if mov.retiro > 0:
                cfdi, tipo = _tomar(pool_r, usados_r, cents, mes_mv), "recibido"
            elif mov.deposito > 0:
                cfdi, tipo = _tomar(pool_e, usados_e, cents, mes_mv), "emitido"

            if cfdi is None:
                pendientes.append(mov)
                continue

            usados = usados_r if tipo == "recibido" else usados_e
            usados.add(cfdi.uuid.lower())
            self._llenar_mov(mov, cfdi, tipo=tipo)
            mov.estatus = "Conciliado Solo Importe"
            nuevos.append(MovConciliado(
                mov, cfdi, tipo, revisar=True, nivel=8,
                tipo_match="Solo Importe",
            ))
            if bitacora is not None:
                bitacora.append({
                    "banco":   getattr(mov, "id_banco", "") or "",
                    "fecha":   mov.fecha,
                    "importe": round(mov.monto, 2),
                    "tipo":    tipo,
                    "uuid":    cfdi.uuid,
                    "serie":   cfdi.serie,
                    "folio":   cfdi.folio,
                    "rfc":     mov.rfc,
                    "nombre":  mov.nombre,
                })

        return nuevos, pendientes


# ── Multi-banco ───────────────────────────────────────────────────────────────

@dataclass
class ResultadoBanco:
    """Resultado de conciliación de un estado de cuenta individual."""
    nombre:       str
    conciliados:  list[MovConciliado]
    sin_cfdi:     list[MovBanco]


def conciliar_multibanco(
    bancos:          list[tuple[str, list[MovBanco]]],
    todos_recibidos: list["CFDI"],  # type: ignore[name-defined]
    todos_emitidos:  list["CFDI"],  # type: ignore[name-defined]
    grupos:          list | None = None,   # list[GrupoConciliacion] — Fase 5
    fases:           "set[str] | frozenset[str] | None" = None,
    log:             "callable | None" = None,
) -> list[ResultadoBanco]:
    """
    Concilia múltiples estados de cuenta en DOS PASADAS:

    PASADA 1 — Cruce exacto (niveles 1-3) para todos los bancos en orden.
               Rápido. Un CFDI usado en banco A no está disponible para banco B.

    PASADA 2 — Suma multi-CFDI (nivel 4) solo para los movimientos que
               quedaron sin conciliar en la pasada 1.
               Solo utiliza CFDIs que NO fueron usados en ningún cruce exacto.

    Esto garantiza que la suma nunca "robe" un CFDI que podría haber hecho
    un cruce exacto en otro estado de cuenta.

    fases : conjunto de ids de fase a ejecutar (ver FASES_IDS). None →
            FASES_DEFAULT (todas menos la Fase 7 "Solo importe", que es
            opcional). Cada fase se omite si su id no está en el conjunto.
    log   : callback opcional log(msg, tag) para escribir la bitácora de la
            Fase 7 (un renglón por cruce por-importe).
    """
    if fases is None:
        fases = FASES_DEFAULT
    # Estampar la posición original de cada movimiento en su estado de cuenta,
    # ANTES de conciliar. La hoja de Excel ordena por este índice para NO
    # reordenar los movimientos (aparecen tal cual el estado de cuenta). Se hace
    # por banco porque cada hoja es un estado de cuenta independiente.
    for _nombre, _movs in bancos:
        for _i, _m in enumerate(_movs):
            _m.orden = _i
    bitacora_f7: list = []   # anti-duplicados / auditoría de la Fase 7
    # Aplicar notas de crédito (tipo "E") una sola vez sobre el pool global.
    # Las facturas relacionadas ven reducido su total efectivo;
    # las notas absorbidas se excluyen del pool y se marcan como "usadas"
    # para que no aparezcan como CFDIs sin movimiento bancario.
    todos_recibidos_adj, absorbidos_r_global = _aplicar_notas_credito(todos_recibidos)
    todos_emitidos_adj,  absorbidos_e_global = _aplicar_notas_credito(todos_emitidos)

    # Fases 2 y 3 (igual que Fase 1 dentro de Conciliador.conciliar) solo
    # consideran CFDIs con forma de pago bancarizada — ver FORMAS_PAGO_BANCARIZADAS.
    # PPD se excluye a propósito: se resuelve en Fase 4 con el pool sin filtrar.
    # NÓMINA (tipo "N") SIEMPRE es elegible aunque su FormaPago no sea bancarizada.
    # Tipo P (REP) fuera de las fases de factura directa (1-3, 5): Fase 4 los
    # resuelve vinculando sus facturas pagadas.
    todos_recibidos_banc = [c for c in todos_recibidos_adj
                            if c.tipo != "P"
                            and (c.forma_pago in FORMAS_PAGO_BANCARIZADAS or c.tipo == "N")]
    todos_emitidos_banc  = [c for c in todos_emitidos_adj
                            if c.tipo != "P"
                            and (c.forma_pago in FORMAS_PAGO_BANCARIZADAS or c.tipo == "N")]

    usados_r: set[str] = set(absorbidos_r_global)
    usados_e: set[str] = set(absorbidos_e_global)
    resultados: list[ResultadoBanco] = []

    # ── PASADA 1: cruce exacto (niveles 1-3) ─────────────────────────────────
    for nombre, movimientos in bancos:
        disponibles_r = [c for c in todos_recibidos_adj if c.uuid.lower() not in usados_r]
        disponibles_e = [c for c in todos_emitidos_adj  if c.uuid.lower() not in usados_e]

        conc = Conciliador(movimientos, disponibles_r, disponibles_e)
        conciliados, sin_cfdi, _, _ = conc.conciliar(skip_suma=True, fases=fases)  # pasada 1: solo Fase 1

        for mc in conciliados:
            if not mc.cfdi:
                continue
            if mc.tipo == "recibido":
                usados_r.add(mc.cfdi.uuid.lower())
            elif mc.tipo == "emitido":
                usados_e.add(mc.cfdi.uuid.lower())

        resultados.append(ResultadoBanco(nombre, conciliados, sin_cfdi))

    # ── Apartar préstamos ANTES de las Fases 2-7 ─────────────────────────────
    # La pasada 1 ya los excluyó y los devolvió dentro de sin_cfdi (con su
    # estatus); aquí se apartan de nuevo para que las fases N:1/combinatorio/
    # PPD/REP/grupos/solo-importe tampoco intenten cruzarlos. Se reincorporan
    # al final.
    prestamos_por_res: list[list[MovBanco]] = []
    for res in resultados:
        prest = [m for m in res.sin_cfdi if _es_no_conciliable(m)]
        prestamos_por_res.append(prest)
        if prest:
            res.sin_cfdi = [m for m in res.sin_cfdi
                            if not _es_no_conciliable(m)]

    # ── FASE 2: PPD con complemento, por RFC (REP) ────────────────────────────
    # El REP cruza exigiendo el RFC del tercero en la descripción (identidad
    # fuerte). Va antes de las fases de suma para que reclame sus facturas antes
    # de que un subset-sum pudiera tomarlas por importe.
    _px_pc = Conciliador([], todos_recibidos, todos_emitidos)
    for res in resultados:
        if not res.sin_cfdi:
            continue
        if "ppd_con_rfc" not in fases:
            break
        mc, aun = _px_pc._conciliar_rep(
            res.sin_cfdi, todos_recibidos_adj, todos_emitidos_adj,
            usados_r, usados_e, modo="rfc")
        res.conciliados.extend(mc)
        res.sin_cfdi = aun

    # ── FASE 3: PPD con complemento, por Nombre (REP) ─────────────────────────
    # Igual que la Fase 2 pero por NOMBRE del tercero (sin RFC), sobre lo que
    # quedó. El RFC (más fuerte) siempre corre primero.
    for res in resultados:
        if not res.sin_cfdi:
            continue
        if "ppd_con_nom" not in fases:
            break
        mc, aun = _px_pc._conciliar_rep(
            res.sin_cfdi, todos_recibidos_adj, todos_emitidos_adj,
            usados_r, usados_e, modo="nombre")
        res.conciliados.extend(mc)
        res.sin_cfdi = aun

    # ── FASE 4: PPD 99 SIN complemento, por RFC ───────────────────────────────
    # Hay movimiento en banco pero el PPD 99 no tiene complemento de pago. Se
    # concilia (pendiente de regularizar) exigiendo el RFC del tercero en la
    # descripción — nunca por importe solo.
    _px_ps = Conciliador([], todos_recibidos, todos_emitidos)
    for res in resultados:
        if not res.sin_cfdi:
            continue
        if "ppd_sin_rfc" not in fases:
            break
        mc, aun = _px_ps._conciliar_ppd(
            res.sin_cfdi, todos_recibidos_adj, todos_emitidos_adj,
            usados_r, usados_e,
            con_complemento=False, tipo_match="PPD Sin Complemento",
            nivel=7, solo_identidad=True, modo="rfc",
        )
        res.conciliados.extend(mc)
        res.sin_cfdi = aun

    # ── FASE 5: PPD 99 SIN complemento, por Nombre ────────────────────────────
    # Igual que la Fase 4 pero por NOMBRE del tercero, sobre lo que quedó.
    for res in resultados:
        if not res.sin_cfdi:
            continue
        if "ppd_sin_nom" not in fases:
            break
        mc, aun = _px_ps._conciliar_ppd(
            res.sin_cfdi, todos_recibidos_adj, todos_emitidos_adj,
            usados_r, usados_e,
            con_complemento=False, tipo_match="PPD Sin Complemento",
            nivel=7, solo_identidad=True, modo="nombre",
        )
        res.conciliados.extend(mc)
        res.sin_cfdi = aun

    # ── FASE 6: Un movimiento = varios CFDIs (1:N) subset-sum ─────────────────
    # Se usa un Conciliador "proxy" solo para acceder a _buscar_suma.
    _proxy_r = Conciliador([], todos_recibidos_adj, todos_emitidos_adj)
    _proxy_e = Conciliador([], todos_recibidos_adj, todos_emitidos_adj)

    # Índice por MES (una sola vez): _buscar_suma solo usa CFDIs del mismo mes y
    # ya filtra `usados` internamente, así que en vez de reconstruir y escanear
    # TODO el pool por cada movimiento (O(movs × pool)), le pasamos solo el bucket
    # del mes del movimiento. En conciliaciones de un año esto era el mayor costo.
    from collections import defaultdict as _dd
    _idx_r = _dd(list); _idx_e = _dd(list)
    for _c in todos_recibidos_banc:
        _idx_r[_c.mes].append(_c)
    for _c in todos_emitidos_banc:
        _idx_e[_c.mes].append(_c)

    for res in resultados:
        if "combinatorio" not in fases or not res.sin_cfdi:
            continue

        nuevos_conc: list[MovConciliado] = []
        aun_sin:     list[MovBanco]      = []

        for mov in res.sin_cfdi:
            es_retiro   = mov.retiro   > 0
            es_deposito = mov.deposito > 0
            suma:      list[CFDI] = []
            tipo_suma: str        = ""

            _mes_mov = mov.mes or _mes_fecha(mov.fecha)
            if es_retiro:
                # Solo el bucket del mes; _buscar_suma filtra `usados` adentro.
                suma = _proxy_r._buscar_suma(mov, _idx_r.get(_mes_mov, []), usados_r)
                if suma:
                    tipo_suma = "recibido"
            elif es_deposito:
                suma = _proxy_e._buscar_suma(mov, _idx_e.get(_mes_mov, []), usados_e)
                if suma:
                    tipo_suma = "emitido"

            if suma:
                usados = usados_r if tipo_suma == "recibido" else usados_e
                for c in suma:
                    usados.add(c.uuid.lower())
                Conciliador._llenar_mov(mov, suma[0], tipo=tipo_suma)
                mov.estatus   = "Conciliado SUMA"
                mov.uuid_suma = ",".join(c.uuid for c in suma[1:])
                nuevos_conc.append(MovConciliado(
                    mov, suma[0], tipo_suma,
                    revisar=True, nivel=4, cfdi_suma=suma[1:],
                    tipo_match="Combinatorio",
                ))
            else:
                aun_sin.append(mov)

        res.conciliados.extend(nuevos_conc)
        res.sin_cfdi = aun_sin

    # ── FASE 7: Varios movimientos = UN mismo CFDI (N:1) ─────────────────────
    _px2 = Conciliador([], todos_recibidos_adj, todos_emitidos_adj)
    for res in resultados:
        if "multi_mov" not in fases or not res.sin_cfdi:
            continue
        mc_mm, aun_sin = _px2._conciliar_multimov(
            res.sin_cfdi, todos_recibidos_banc, todos_emitidos_banc,
            usados_r, usados_e,
        )
        res.conciliados.extend(mc_mm)
        res.sin_cfdi = aun_sin

    # ── FASE 7b: PPD 99 (factura) liquidada por VARIOS movimientos ────────────
    # Las PPD 99 están fuera del pool bancarizado, así que la Fase 7 nunca las
    # ve. Subset-sum de los movimientos contra el total de la PPD, sin candado
    # de mes (las parcialidades pueden caer en meses distintos), con identidad.
    for res in resultados:
        if "multi_mov" not in fases or not res.sin_cfdi:
            continue
        ppd_r = [c for c in todos_recibidos_adj
                 if c.metodo_pago == "PPD" and c.forma_pago == "99"
                 and c.uuid.lower() not in usados_r]
        ppd_e = [c for c in todos_emitidos_adj
                 if c.metodo_pago == "PPD" and c.forma_pago == "99"
                 and c.uuid.lower() not in usados_e]
        if ppd_r or ppd_e:
            mc_p, aun_p = _px2._conciliar_multimov(
                res.sin_cfdi, ppd_r, ppd_e, usados_r, usados_e,
                exigir_mismo_mes=False)
            res.conciliados.extend(mc_p)
            res.sin_cfdi = aun_p

    # ── FASE 7c: Comisiones bancarias (comisión + IVA => factura del banco) ───
    for res in resultados:
        if "multi_mov" not in fases or not res.sin_cfdi:
            continue
        mc_c, aun_c = _px2._conciliar_comisiones_banco(
            res.sin_cfdi, todos_recibidos_adj, usados_r)
        res.conciliados.extend(mc_c)
        res.sin_cfdi = aun_c

    # ── FASE 8: Agrupado por RFC (grupos configurados + automático) ───────────
    if "grupos" in fases:
        _px5 = Conciliador([], todos_recibidos_adj, todos_emitidos_adj)
        for res in resultados:
            if not res.sin_cfdi:
                continue
            # (a) Grupos multi-RFC configurados
            if grupos:
                pool_r_g = [c for c in todos_recibidos_adj if c.uuid.lower() not in usados_r]
                pool_e_g = [c for c in todos_emitidos_adj  if c.uuid.lower() not in usados_e]
                mc_g5, aun_sin_g5 = Conciliador._conciliar_grupos(
                    res.sin_cfdi, pool_r_g, pool_e_g, usados_r, usados_e, grupos,
                )
                res.conciliados.extend(mc_g5)
                res.sin_cfdi = aun_sin_g5
            # (b) Automático por RFC+mes (antes "Global Mensual")
            if res.sin_cfdi:
                pool_r_d = [c for c in todos_recibidos_banc if c.uuid.lower() not in usados_r]
                pool_e_d = [c for c in todos_emitidos_banc  if c.uuid.lower() not in usados_e]
                mc_gm, aun_sin_gm = Conciliador._buscar_global_mensual(
                    res.sin_cfdi, pool_r_d, pool_e_d, usados_r, usados_e
                )
                res.conciliados.extend(mc_gm)
                res.sin_cfdi = aun_sin_gm

    # ── FASE 9: Solo importe (OPCIONAL, último recurso) ───────────────────────
    # Cruza lo que siga pendiente contra cualquier CFDI no usado cuyo importe
    # coincida, SIN mes ni identidad. Anti-duplicados vía usados_r/usados_e
    # (bitácora de comprobantes consumidos). Cada cruce se anota en bitacora_f7
    # y se escribe al log. Todo match => revisar=True.
    if "solo_importe" in fases:
        _px7 = Conciliador([], todos_recibidos, todos_emitidos)
        for res in resultados:
            if not res.sin_cfdi:
                continue
            # id_banco vacío en los MovBanco → usar el nombre del estado de cuenta
            for _m in res.sin_cfdi:
                if not getattr(_m, "id_banco", ""):
                    _m.id_banco = res.nombre
            mc_f7, aun_sin_f7 = _px7._conciliar_solo_importe(
                res.sin_cfdi, todos_recibidos_adj, todos_emitidos_adj,
                usados_r, usados_e, bitacora=bitacora_f7,
            )
            res.conciliados.extend(mc_f7)
            res.sin_cfdi = aun_sin_f7

        if bitacora_f7 and log:
            log(f"Fase 7 (solo importe): {len(bitacora_f7)} cruce(s) por importe "
                f"— TODOS requieren revisión:", "warn")
            for e in bitacora_f7:
                log(f"    • {e['banco']} {e['fecha']}  ${e['importe']:,.2f}  "
                    f"[{e['tipo']}]  serie {e['serie'] or '-'} folio {e['folio'] or '-'}  "
                    f"UUID {e['uuid']}", "warn")

    # Reincorporar los préstamos como pendientes de cada banco.
    for res, prest in zip(resultados, prestamos_por_res):
        res.sin_cfdi.extend(prest)

    return resultados


# ── Conversión desde Movimiento bancario ─────────────────────────────────────

def mov_a_banco(mov,  # bancos.base.Movimiento
                mes: int = 0,
                id_banco: str = "",
                anio: int = 0) -> MovBanco:
    """Convierte un Movimiento del extractor bancario en un MovBanco.
    Si *anio* y la fecha es DD/MM (sin año), la completa a DD/MM/YYYY."""
    fecha = mov.fecha or ""
    if anio and re.fullmatch(r"\d{1,2}/\d{1,2}", fecha.strip()):
        fecha = f"{fecha.strip()}/{anio}"
    return MovBanco(
        fecha       = fecha,
        folio       = mov.folio or "",
        descripcion = (mov.descripcion or "").replace(" | ", " | "),
        retiro      = mov.retiro   or 0.0,
        deposito    = mov.deposito or 0.0,
        saldo       = mov.saldo    or 0.0,
        mes         = mes or _mes_fecha(fecha),
        id_banco    = id_banco,
    )