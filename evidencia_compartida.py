# -*- coding: utf-8 -*-
"""
evidencia_compartida.py — v5.0.1
=================================
Parche ADITIVO para ConvertidorBancos v5.0 (cruce Bancos ↔ CFDI).

Objetivo
--------
El motor de `conciliacion.py` (V5) ya concilia por importe + mes + identidad.
Este módulo aporta una **jerarquía de evidencia** independiente y reutilizable que
califica QUÉ TAN FUERTE es un cruce, SIN modificar `conciliacion.py`, la UI, ni
ningún parser de banco:

    CLAVE_REFERENCIA  →  CONTRAPARTE_RFC  →  IMPORTE_FECHA  →  NINGUNA
      (más fuerte)                                              (sin cruce)

Principios de diseño (no negociables)
-------------------------------------
1. ADITIVO: es un archivo nuevo. No importa ni modifica `conciliacion.py`,
   `ConvertidorSantander.py`, `exportar_conciliacion.py` ni los parsers de
   `bancos/`. Solo usa la librería estándar de Python.
2. El motor **opera sobre la descripción YA ENSAMBLADA** (`MovBanco.descripcion`,
   o cualquier string). No re-parsea PDFs ni conoce el layout de cada banco.
3. Las **etiquetas por banco son DATO**, no código: viven en un `PerfilBanco`
   (dict) que se le pasa al motor. Agregar un banco = agregar un perfil.
4. `evaluar_match` acepta tanto un **dict** como un **objeto CFDI** del app
   (lee `uuid/folio/serie/rfc_emisor/rfc_receptor/total/fecha` o `fecha_emision`),
   así se puede llamar desde el motor del app sin adaptadores.

Uso desde el app (ejemplo, sin tocar conciliacion.py)
-----------------------------------------------------
    from evidencia_compartida import MotorEvidencia, PERFILES, evaluar_match

    motor = MotorEvidencia(PERFILES["BBVA"])
    ev = motor.extraer(mov.descripcion, importe=mov.monto, fecha=mov.fecha)
    nivel = evaluar_match(ev, cfdi)      # cfdi = CFDI del app o un dict
    if nivel >= NivelEvidencia.CONTRAPARTE_RFC:
        ...  # cruce fuerte

Migración de esquema (opcional, para persistir el nivel en conciliaciones.db)
-----------------------------------------------------------------------------
    import sqlite3
    conn = sqlite3.connect(ruta_db)
    aplicar_migracion(conn, tabla="movimientos_cruce")   # idempotente

Ejecuta `python evidencia_compartida.py` para correr la autoprueba.
"""

from __future__ import annotations

import re
import sqlite3
from dataclasses import dataclass, field
from enum import IntEnum
from typing import Optional

VERSION = "5.0.1"

# ─────────────────────────────────────────────────────────────────────────────
#  Jerarquía de evidencia
# ─────────────────────────────────────────────────────────────────────────────


class NivelEvidencia(IntEnum):
    """Fuerza del cruce banco ↔ CFDI. Mayor valor = evidencia más fuerte.

    El orden numérico ES la jerarquía: comparar/ordenar por el valor entero da
    directamente la prioridad (CLAVE_REFERENCIA gana a CONTRAPARTE_RFC, etc.).
    """
    NINGUNA          = 0   # ninguna señal — no se debe cruzar
    IMPORTE_FECHA    = 1   # solo coincide importe + periodo (lo que hace V5 nivel 3.5)
    CONTRAPARTE_RFC  = 2   # además coincide el RFC del tercero
    CLAVE_REFERENCIA = 3   # la referencia/folio del banco apunta al CFDI (lo más fuerte)

    @property
    def etiqueta(self) -> str:
        return {
            NivelEvidencia.CLAVE_REFERENCIA: "Clave/Referencia",
            NivelEvidencia.CONTRAPARTE_RFC:  "RFC contraparte",
            NivelEvidencia.IMPORTE_FECHA:    "Importe+Fecha",
            NivelEvidencia.NINGUNA:          "Sin evidencia",
        }[self]

    @property
    def requiere_revision(self) -> bool:
        """Solo IMPORTE_FECHA (o menos) exige revisión humana: no hay identidad."""
        return self <= NivelEvidencia.IMPORTE_FECHA


# ─────────────────────────────────────────────────────────────────────────────
#  Evidencia extraída de un movimiento bancario
# ─────────────────────────────────────────────────────────────────────────────


@dataclass
class Evidencia:
    """Señales de cruce extraídas de la descripción YA ENSAMBLADA del movimiento.

    `clave_referencia` y `contraparte_rfc` salen del texto; `importe` y `fecha`
    los aporta el propio movimiento (no la descripción) y se guardan aquí para
    que `evaluar_match` tenga todo en un solo objeto.
    """
    descripcion:      str
    banco:            str = ""
    importe:          Optional[float] = None
    fecha:            str = ""            # ISO 'YYYY-MM-DD' o 'DD/MM/YYYY'
    clave_referencia: str = ""           # clave de rastreo / referencia / folio del banco
    contraparte_rfc:  str = ""           # RFC del tercero detectado en la descripción
    # tokens crudos detectados (útil para depurar / auditar)
    tokens_referencia: list[str] = field(default_factory=list)

    def tiene_clave(self) -> bool:
        return bool(self.clave_referencia)

    def tiene_rfc(self) -> bool:
        return bool(self.contraparte_rfc)


# ─────────────────────────────────────────────────────────────────────────────
#  Perfiles de banco (DATO, no código)
# ─────────────────────────────────────────────────────────────────────────────
#  Cada perfil describe DÓNDE mirar en la descripción ya ensamblada de ese banco.
#  Agregar un banco = agregar una entrada aquí; el motor no cambia.

PerfilBanco = dict  # alias legible: {"banco": str, "etiquetas_referencia": (...), ...}

PERFILES: dict[str, PerfilBanco] = {
    "BBVA": {
        "banco": "BBVA",
        "etiquetas_referencia": (
            "CLAVE DE RASTREO", "CLAVE RASTREO", "REFERENCIA", "REF.", "REF ",
            "FOLIO", "NUM. REFERENCIA",
        ),
        "etiquetas_rfc": ("RFC:", "R.F.C:", "R.F.C.", "RFC "),
        "separadores": ("|", " - ", " / ", ",", "/"),
    },
    "BANORTE": {
        "banco": "BANORTE",
        "etiquetas_referencia": (
            "CLAVE DE RASTREO", "CLAVE RASTREO", "REFERENCIA", "REF.", "REF ",
            "FOLIO", "CVE RASTREO",
        ),
        "etiquetas_rfc": ("RFC:", "R.F.C:", "RFC ", "RFC/CURP"),
        "separadores": ("|", " - ", " / ", ",", "/"),
    },
    "INBURSA": {
        "banco": "INBURSA",
        "etiquetas_referencia": (
            "CLAVE DE RASTREO", "REFERENCIA", "REF.", "FOLIO",
        ),
        "etiquetas_rfc": ("RFC", "R.F.C"),
        "separadores": ("|", " - ", " / ", ",", "/"),
    },
    "BANREGIO": {
        "banco": "BANREGIO",
        "etiquetas_referencia": (
            "CLAVE DE RASTREO", "REFERENCIA", "REF.", "FOLIO",
        ),
        "etiquetas_rfc": ("R.F.C.", "RFC:", "RFC "),
        "separadores": (",", "|", " - ", " / ", "/"),
    },
    # Perfil genérico de respaldo para bancos sin perfil propio.
    "_DEFAULT": {
        "banco": "GENERICO",
        "etiquetas_referencia": ("CLAVE DE RASTREO", "REFERENCIA", "REF.", "FOLIO"),
        "etiquetas_rfc": ("RFC:", "RFC "),
        "separadores": ("|", " - ", " / ", ",", "/"),
    },
}

# RFC mexicano: 3-4 letras (incluye Ñ y &) + 6 dígitos (fecha) + 3 alfanum (homoclave)
_RFC_RE = re.compile(r"\b([A-ZÑ&]{3,4}\d{6}[A-Z0-9]{3})\b")
# Token de clave/referencia: alfanumérico "largo" (clave de rastreo, folio, etc.)
_TOKEN_REF_RE = re.compile(r"[A-Z0-9][A-Z0-9\-]{4,}")


# ─────────────────────────────────────────────────────────────────────────────
#  Motor de evidencia
# ─────────────────────────────────────────────────────────────────────────────


class MotorEvidencia:
    """Extrae `Evidencia` de una descripción bancaria ya ensamblada.

    Se construye con un `PerfilBanco` (dato). El mismo motor sirve para cualquier
    banco: lo único que cambia es el perfil.
    """

    def __init__(self, perfil: PerfilBanco):
        if not perfil or "banco" not in perfil:
            raise ValueError("perfil de banco inválido: falta 'banco'")
        self.perfil = perfil
        self.banco = perfil.get("banco", "")
        self.etiquetas_referencia = tuple(perfil.get("etiquetas_referencia", ()))
        self.etiquetas_rfc = tuple(perfil.get("etiquetas_rfc", ()))

    # ── API pública ─────────────────────────────────────────────────────────
    def extraer(self, descripcion: str,
                importe: Optional[float] = None,
                fecha: str = "") -> Evidencia:
        desc = descripcion or ""
        up = desc.upper()

        rfc = self._extraer_rfc(up)
        clave, tokens = self._extraer_referencia(up)

        return Evidencia(
            descripcion=desc,
            banco=self.banco,
            importe=importe,
            fecha=fecha or "",
            clave_referencia=clave,
            contraparte_rfc=rfc,
            tokens_referencia=tokens,
        )

    # ── Internos ──────────────────────────────────────────────────────────────
    def _extraer_rfc(self, up: str) -> str:
        # 1) RFC tras una etiqueta ("RFC: XXX") — evidencia más limpia
        for etq in self.etiquetas_rfc:
            i = up.find(etq)
            if i >= 0:
                m = _RFC_RE.search(up, i + len(etq))
                # exigir que el RFC esté cerca de la etiqueta (mismo campo)
                if m and (m.start() - (i + len(etq))) <= 6:
                    return m.group(1)
        # 2) RFC "suelto" en cualquier parte de la descripción
        m = _RFC_RE.search(up)
        return m.group(1) if m else ""

    def _extraer_referencia(self, up: str) -> tuple[str, list[str]]:
        """Devuelve (clave_referencia principal, todos los tokens candidatos)."""
        tokens: list[str] = []
        clave = ""
        for etq in self.etiquetas_referencia:
            i = up.find(etq)
            if i < 0:
                continue
            resto = up[i + len(etq):].lstrip(" :.-#")
            m = _TOKEN_REF_RE.match(resto)
            if m:
                tok = m.group(0).strip("-")
                # descartar que el token sea en realidad un RFC
                if not _RFC_RE.fullmatch(tok):
                    tokens.append(tok)
                    if not clave:
                        clave = tok
        return clave, tokens


# ─────────────────────────────────────────────────────────────────────────────
#  Evaluación del cruce contra un CFDI (dict u objeto CFDI del app)
# ─────────────────────────────────────────────────────────────────────────────

TOLERANCIA_IMPORTE = 0.02   # ±2 centavos, igual que V5


def _get(cfdi, *nombres):
    """Lee un campo del CFDI ya sea un dict o un objeto (dataclass CFDI del app).
    Acepta varios nombres alternativos (p. ej. 'fecha' o 'fecha_emision')."""
    for n in nombres:
        if isinstance(cfdi, dict):
            if n in cfdi and cfdi[n] not in (None, ""):
                return cfdi[n]
        else:
            v = getattr(cfdi, n, None)
            if v not in (None, ""):
                return v
    return None


def _norm(s) -> str:
    return str(s if s is not None else "").strip().upper()


def _mismo_periodo(fecha_mov: str, fecha_cfdi: str) -> bool:
    """True si caen en el mismo AÑO-MES. Acepta ISO 'YYYY-MM-DD' y 'DD/MM/YYYY'."""
    def ym(f: str):
        f = (f or "").strip()
        if not f:
            return None
        m = re.match(r"(\d{4})-(\d{2})-(\d{2})", f)          # ISO
        if m:
            return int(m.group(1)), int(m.group(2))
        m = re.match(r"(\d{2})/(\d{2})/(\d{4})", f)          # DD/MM/YYYY
        if m:
            return int(m.group(3)), int(m.group(2))
        return None
    a, b = ym(fecha_mov), ym(fecha_cfdi)
    if a is None or b is None:
        return False
    return a == b


def _clave_apunta_al_cfdi(evidencia: Evidencia, cfdi) -> bool:
    """True si la referencia del banco contiene el folio, serie+folio o cola del
    UUID del CFDI (el banco 'apunta' explícitamente a ese comprobante)."""
    universo = _norm(evidencia.clave_referencia) + "|" + _norm(evidencia.descripcion)
    universo = universo.replace(" ", "")

    folio = _norm(_get(cfdi, "folio")).replace(" ", "")
    serie = _norm(_get(cfdi, "serie")).replace(" ", "")
    uuid = _norm(_get(cfdi, "uuid")).replace("-", "")

    # folio debe ser suficientemente específico para no dar falsos positivos
    if folio and len(folio) >= 3 and folio in universo:
        return True
    if serie and folio and (serie + folio) in universo:
        return True
    if uuid and len(uuid) >= 8 and uuid[-12:] in universo:
        return True
    return False


def _rfc_contraparte_coincide(evidencia: Evidencia, cfdi) -> bool:
    rfc = _norm(evidencia.contraparte_rfc)
    if not rfc:
        return False
    return rfc in (_norm(_get(cfdi, "rfc_emisor")), _norm(_get(cfdi, "rfc_receptor")))


def _importe_fecha_coincide(evidencia: Evidencia, cfdi,
                            tol: float = TOLERANCIA_IMPORTE) -> bool:
    total = _get(cfdi, "total")
    if evidencia.importe is None or total is None:
        return False
    if abs(float(evidencia.importe) - float(total)) > tol:
        return False
    # el periodo debe coincidir; si el CFDI no trae fecha, el importe solo NO basta
    return _mismo_periodo(evidencia.fecha, _norm(_get(cfdi, "fecha", "fecha_emision")))


def evaluar_match(evidencia: Evidencia, cfdi,
                  tol_importe: float = TOLERANCIA_IMPORTE) -> NivelEvidencia:
    """Devuelve el NIVEL MÁS ALTO de la jerarquía que se cumple entre el
    movimiento (su `Evidencia`) y un `cfdi` (dict u objeto CFDI del app con
    uuid/folio/serie/rfc_emisor/rfc_receptor/total/fecha|fecha_emision).

    Jerarquía:  CLAVE_REFERENCIA → CONTRAPARTE_RFC → IMPORTE_FECHA → NINGUNA.

    La clave/RFC por sí solos NO cruzan si el importe no cuadra: se exige que el
    importe coincida (±tol) en TODOS los niveles — la clave o el RFC solo ELEVAN
    la confianza de un cruce que ya cuadra en dinero. Esto evita cruzar un pago a
    un proveedor contra una factura distinta del mismo proveedor.
    """
    if not _importe_fecha_coincide(evidencia, cfdi, tol_importe):
        return NivelEvidencia.NINGUNA

    if _clave_apunta_al_cfdi(evidencia, cfdi):
        return NivelEvidencia.CLAVE_REFERENCIA
    if _rfc_contraparte_coincide(evidencia, cfdi):
        return NivelEvidencia.CONTRAPARTE_RFC
    return NivelEvidencia.IMPORTE_FECHA


# ─────────────────────────────────────────────────────────────────────────────
#  Migración de esquema (aditiva, idempotente) — para persistencia SQLite
# ─────────────────────────────────────────────────────────────────────────────
#  Estas columnas EXTIENDEN la tabla del cruce; no reemplazan ninguna de V5.

COLUMNAS_NUEVAS: list[tuple[str, str]] = [
    ("nivel_evidencia",  "INTEGER DEFAULT 0"),   # valor de NivelEvidencia
    ("evidencia_nombre", "TEXT DEFAULT ''"),      # etiqueta legible
    ("evidencia_clave",  "TEXT DEFAULT ''"),      # clave/referencia detectada
    ("evidencia_rfc",    "TEXT DEFAULT ''"),      # RFC contraparte detectado
    ("evidencia_version","TEXT DEFAULT ''"),      # versión del motor que la calculó
]


_TABLAS_PERMITIDAS = {"movimientos_cruce"}


def _tabla_valida(tabla: str) -> str:
    if tabla not in _TABLAS_PERMITIDAS:
        raise ValueError(f"Tabla no permitida: {tabla!r}")
    return tabla


def sql_migracion(tabla: str = "movimientos_cruce") -> list[str]:
    """Devuelve los ALTER TABLE (uno por columna) para agregar las columnas
    nuevas. SQLite no soporta 'ADD COLUMN IF NOT EXISTS', por eso `aplicar_migracion`
    consulta primero qué columnas ya existen."""
    tabla = _tabla_valida(tabla)
    return [
        f"ALTER TABLE {tabla} ADD COLUMN {col} {tipo};"
        for col, tipo in COLUMNAS_NUEVAS
    ]


def aplicar_migracion(conn: sqlite3.Connection,
                      tabla: str = "movimientos_cruce") -> list[str]:
    """Aplica la migración de forma IDEMPOTENTE: solo agrega las columnas que
    falten. Devuelve la lista de columnas realmente agregadas. Si la tabla no
    existe, no hace nada (devuelve [])."""
    tabla = _tabla_valida(tabla)
    cur = conn.cursor()
    cur.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name=?", (tabla,))
    if cur.fetchone() is None:
        return []

    cur.execute(f"PRAGMA table_info({tabla})")
    existentes = {row[1] for row in cur.fetchall()}   # row[1] = nombre de columna

    agregadas: list[str] = []
    for col, tipo in COLUMNAS_NUEVAS:
        if col not in existentes:
            cur.execute(f"ALTER TABLE {tabla} ADD COLUMN {col} {tipo}")
            agregadas.append(col)
    conn.commit()
    return agregadas


# ─────────────────────────────────────────────────────────────────────────────
#  Autoprueba
# ─────────────────────────────────────────────────────────────────────────────


def _autoprueba() -> bool:
    print(f"evidencia_compartida v{VERSION} - autoprueba\n" + "=" * 52)
    ok = True

    def check(nombre: str, cond: bool):
        nonlocal ok
        ok = ok and cond
        print(f"  [{'OK ' if cond else 'FALLA'}] {nombre}")

    # ── Fixtures BBVA ────────────────────────────────────────────────────────
    motor_bbva = MotorEvidencia(PERFILES["BBVA"])
    desc_bbva = ("SPEI ENVIADO STP - JUAN PEREZ LOPEZ RFC: PELJ800101AB1 "
                 "CLAVE DE RASTREO STP12026030100012345 FOLIO 1234")
    ev_bbva = motor_bbva.extraer(desc_bbva, importe=5000.00, fecha="2026-03-01")

    print("\n-- BBVA: extraccion --")
    print(f"   rfc={ev_bbva.contraparte_rfc!r}  clave={ev_bbva.clave_referencia!r}")
    check("BBVA extrae RFC contraparte", ev_bbva.contraparte_rfc == "PELJ800101AB1")
    check("BBVA extrae clave de rastreo", ev_bbva.clave_referencia.startswith("STP1"))

    cfdi_folio = dict(uuid="AAAA1111-2222-3333-4444-555566667777", serie="A",
                      folio="1234", rfc_emisor="PELJ800101AB1",
                      rfc_receptor="EMP010101AA1", total=5000.00, fecha="2026-03-05")
    check("BBVA folio referenciado -> CLAVE_REFERENCIA",
          evaluar_match(ev_bbva, cfdi_folio) == NivelEvidencia.CLAVE_REFERENCIA)

    cfdi_rfc = dict(uuid="BBBB1111-2222-3333-4444-555566667777", serie="B",
                    folio="9999", rfc_emisor="PELJ800101AB1",
                    rfc_receptor="EMP010101AA1", total=5000.00, fecha="2026-03-20")
    check("BBVA mismo RFC, otro folio -> CONTRAPARTE_RFC",
          evaluar_match(ev_bbva, cfdi_rfc) == NivelEvidencia.CONTRAPARTE_RFC)

    cfdi_otro = dict(uuid="CCCC1111-2222-3333-4444-555566667777", serie="C",
                     folio="7777", rfc_emisor="XXX010101XX1",
                     rfc_receptor="EMP010101AA1", total=5000.00, fecha="2026-03-10")
    check("BBVA otro proveedor, mismo importe -> IMPORTE_FECHA",
          evaluar_match(ev_bbva, cfdi_otro) == NivelEvidencia.IMPORTE_FECHA)

    cfdi_otromes = dict(uuid="DDDD1111-2222-3333-4444-555566667777", serie="D",
                        folio="8888", rfc_emisor="PELJ800101AB1",
                        rfc_receptor="EMP010101AA1", total=5000.00, fecha="2026-04-10")
    check("BBVA otro mes -> NINGUNA",
          evaluar_match(ev_bbva, cfdi_otromes) == NivelEvidencia.NINGUNA)

    cfdi_otroimporte = dict(uuid="EEEE1111-2222-3333-4444-555566667777", serie="E",
                            folio="1234", rfc_emisor="PELJ800101AB1",
                            rfc_receptor="EMP010101AA1", total=5001.00, fecha="2026-03-05")
    check("BBVA importe distinto -> NINGUNA",
          evaluar_match(ev_bbva, cfdi_otroimporte) == NivelEvidencia.NINGUNA)

    # ── Fixtures BANORTE ─────────────────────────────────────────────────────
    motor_bnt = MotorEvidencia(PERFILES["BANORTE"])
    desc_bnt = ("BANORTE | SPEI RECIBIDO | COMERCIALIZADORA MAKULIS SA DE CV | "
                "RFC CMA150203QX9 | CLAVE DE RASTREO 072BNET00099887766 | "
                "CONCEPTO PAGO FACTURA A-500")
    ev_bnt = motor_bnt.extraer(desc_bnt, importe=12500.50, fecha="15/03/2026")

    print("\n-- BANORTE: extraccion --")
    print(f"   rfc={ev_bnt.contraparte_rfc!r}  clave={ev_bnt.clave_referencia!r}")
    check("BANORTE extrae RFC contraparte", ev_bnt.contraparte_rfc == "CMA150203QX9")
    check("BANORTE extrae clave de rastreo",
          ev_bnt.clave_referencia == "072BNET00099887766")

    cfdi_emit = dict(uuid="1111AAAA-2222-3333-4444-555566667777", serie="A",
                     folio="500", rfc_emisor="EMP010101AA1",
                     rfc_receptor="CMA150203QX9", total=12500.50, fecha="2026-03-15")
    check("BANORTE folio en concepto -> CLAVE_REFERENCIA",
          evaluar_match(ev_bnt, cfdi_emit) == NivelEvidencia.CLAVE_REFERENCIA)

    cfdi_emit2 = dict(uuid="2222AAAA-2222-3333-4444-555566667777", serie="A",
                      folio="777", rfc_emisor="EMP010101AA1",
                      rfc_receptor="CMA150203QX9", total=12500.50, fecha="2026-03-28")
    check("BANORTE mismo cliente, otro folio -> CONTRAPARTE_RFC",
          evaluar_match(ev_bnt, cfdi_emit2) == NivelEvidencia.CONTRAPARTE_RFC)

    ev_pobre = motor_bnt.extraer("BANORTE | ABONO | DEPOSITO EN EFECTIVO",
                                 importe=12500.50, fecha="15/03/2026")
    check("BANORTE sin identidad -> IMPORTE_FECHA",
          evaluar_match(ev_pobre, cfdi_emit2) == NivelEvidencia.IMPORTE_FECHA)

    # ── Compatibilidad con OBJETO CFDI del app (duck typing) ─────────────────
    print("\n-- Objeto CFDI del app (fecha_emision) --")

    class _CFDIapp:
        def __init__(self, **kw):
            self.__dict__.update(kw)

    cfdi_obj = _CFDIapp(uuid="9999AAAA-2222-3333-4444-555566667777", serie="A",
                        folio="1234", rfc_emisor="PELJ800101AB1",
                        rfc_receptor="EMP010101AA1", total=5000.00,
                        fecha_emision="2026-03-05")  # el app usa fecha_emision
    check("objeto CFDI (fecha_emision) -> CLAVE_REFERENCIA",
          evaluar_match(ev_bbva, cfdi_obj) == NivelEvidencia.CLAVE_REFERENCIA)

    # ── Migración SQLite (idempotente) ───────────────────────────────────────
    print("\n-- Migracion SQLite --")
    conn = sqlite3.connect(":memory:")
    conn.execute("CREATE TABLE movimientos_cruce (id INTEGER PRIMARY KEY, importe REAL)")
    ag1 = aplicar_migracion(conn, "movimientos_cruce")
    ag2 = aplicar_migracion(conn, "movimientos_cruce")   # segunda vez: no debe agregar nada
    cols = {r[1] for r in conn.execute("PRAGMA table_info(movimientos_cruce)")}
    print(f"   agregadas 1a vez: {ag1}")
    print(f"   agregadas 2a vez: {ag2}  (debe ser [])")
    check("migracion agrega las 5 columnas", len(ag1) == len(COLUMNAS_NUEVAS))
    check("migracion es idempotente", ag2 == [])
    check("todas las columnas nuevas presentes",
          all(c in cols for c, _ in COLUMNAS_NUEVAS))
    check("migracion no falla si la tabla no existe",
          aplicar_migracion(conn, "tabla_inexistente") == [])
    conn.close()

    check("jerarquia CLAVE > RFC > IMPORTE > NINGUNA",
          NivelEvidencia.CLAVE_REFERENCIA > NivelEvidencia.CONTRAPARTE_RFC
          > NivelEvidencia.IMPORTE_FECHA > NivelEvidencia.NINGUNA)

    print("\n" + "=" * 52)
    print("RESULTADO:", "TODAS PASAN [OK]" if ok else "HAY FALLAS [X]")
    return ok


if __name__ == "__main__":
    import sys
    sys.exit(0 if _autoprueba() else 1)
