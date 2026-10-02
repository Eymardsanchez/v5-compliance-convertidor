# -*- coding: utf-8 -*-
"""
Capa de persistencia de la conciliación (Etapa 2 — BD incremental).

Objetivo
--------
Guardar cada fila conciliada con un **ID ESTABLE** que NUNCA se reusa (aunque el
usuario borre la fila del Excel). Ese ID es la llave de sincronización entre el
Excel revisado y el programa: al reimportar, se hace match por ID y no se
sobrescribe información equivocada.

Diseño
------
- ``Persistencia``           : interfaz abstracta (para poder cambiar de motor).
- ``FirebirdPersistencia``   : implementación con Firebird (servidor local).

En Firebird el ID estable se logra con un **GENERADOR** (secuencia): entrega
números crecientes que jamás retroceden aunque se borren registros → deja huecos,
que es exactamente el comportamiento pedido.

La BD vive en local (``%PROGRAMDATA%\\ConvertidorBancos``), NUNCA en OneDrive.

Control de concurrencia (evitar duplicidad / pérdida de datos)
--------------------------------------------------------------
Si dos instancias del programa (o dos usuarios sobre una BD compartida) editan la
misma fila, sin control uno pisa al otro en silencio ("last-writer-wins") o la UI
se cuelga esperando el lock. Para evitarlo:

1. **TPB explícito** por conexión: ``read_committed + rec_version + NOWAIT``. Da
   lecturas siempre frescas y hace que un choque de escritura FALLE de inmediato
   (código Firebird 335544451 / SQLSTATE 40001) en vez de colgar. Se fuerza vía
   TPB porque el servidor en producción es Firebird **2.5** (verificado con
   ``ENGINE_VERSION``): ``ReadConsistency`` (FB4+) no existe aquí.
2. **Columna de versión** (``VERSION_REG``) + ``FECHA_MODIFICACION`` en las tablas
   editables (FILA, MOV_BANCO). El UPDATE es condicional por versión
   (``WHERE ID=? AND VERSION_REG=?``); si la versión ya cambió, se detecta el
   conflicto (``ConflictoConcurrencia``) — garantía independiente de la versión
   del motor. Índice sobre ``FECHA_MODIFICACION`` para el checkpoint incremental.
3. **Bitácora de auditoría** (``AUDITORIA_CONFLICTO``): cada conflicto queda
   registrado (quién, cuándo, qué versión ganó) para trazabilidad de pre-auditoría.
"""

from __future__ import annotations

import os
from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional


# ── Modelo de fila persistida ────────────────────────────────────────────────
@dataclass
class FilaConciliacion:
    id:              int   = 0     # ID ESTABLE (asignado por el generador)
    conciliacion_id: int   = 0
    rfc:             str   = ""
    banco:           str   = ""
    fecha:           str   = ""    # ISO YYYY-MM-DD
    importe:         float = 0.0
    concepto:        str   = ""
    uuid_cfdi:       str   = ""
    uuid_rep:        str   = ""
    tipo_match:      str   = ""
    revisar:         bool  = False
    cuenta:          str   = ""    # cuenta contable (COI)
    tercero:         str   = ""
    # Campos para regenerar pólizas desde la BD (autosuficiente, sin los XML):
    tipo:            str   = ""    # "recibido" | "emitido"
    uuid_suma:       str   = ""    # UUIDs adicionales (suma multi-CFDI), separados por coma
    deposito:        float = 0.0
    retiro:          float = 0.0
    # Control de concurrencia (no viene del Excel; lo administra la BD):
    version_reg:        int = 0    # se incrementa en cada UPDATE; llave optimista
    fecha_modificacion: str = ""   # ISO del último cambio


# ── Modelo de CFDI persistido (lo que la póliza necesita, sin volver al XML) ──
@dataclass
class CFDIData:
    uuid:            str   = ""    # PK
    tipo:            str   = ""    # I/E/N/P
    rfc_emisor:      str   = ""
    rfc_receptor:    str   = ""
    nombre_emisor:   str   = ""
    nombre_receptor: str   = ""
    subtotal:        float = 0.0
    total:           float = 0.0
    descuento:       float = 0.0
    iva:             float = 0.0
    iva_retenido:    float = 0.0
    isr_retenido:    float = 0.0
    clave_prod_serv: str   = ""
    fecha_emision:   str   = ""


# ── Modelo de movimiento bancario CONVERTIDO (registro PDF→Excel) ─────────────
@dataclass
class MovimientoBanco:
    id:             int   = 0     # ID ESTABLE (llave de sync con el Excel)
    conversion_id:  int   = 0
    banco:          str   = ""
    fecha:          str   = ""
    folio:          str   = ""
    descripcion:    str   = ""
    deposito:       float = 0.0
    retiro:         float = 0.0
    saldo:          float = 0.0
    # Control de concurrencia:
    version_reg:        int = 0
    fecha_modificacion: str = ""


# ── Conflicto de concurrencia ────────────────────────────────────────────────
class ConflictoConcurrencia(Exception):
    """Se lanza cuando el registro que se intenta escribir ya fue modificado por
    otro (la versión guardada no coincide con la que el usuario tenía cargada).

    Lleva ambos valores para que la capa de aplicación se los muestre al usuario y
    lo deje decidir: sobrescribir, descartar su cambio, o fusionar.
    """

    def __init__(self, tabla: str, registro_id: int, local, guardado,
                 version_local, version_guardada):
        self.tabla            = tabla
        self.registro_id      = registro_id
        self.local            = local            # lo que el usuario intentó guardar
        self.guardado         = guardado         # lo que hay ahora en la BD
        self.version_local    = version_local
        self.version_guardada = version_guardada
        super().__init__(
            f"Conflicto de concurrencia en {tabla} ID={registro_id}: "
            f"tu versión ({version_local}) ya no es la actual ({version_guardada}). "
            f"Otro proceso modificó este registro.")


def _es_update_conflict(exc) -> bool:
    """True si la excepción de Firebird es el conflicto de actualización
    concurrente (isc_update_conflict = 335544451) o un deadlock por lock (40001)."""
    if 335544451 in tuple(getattr(exc, "gds_codes", ()) or ()):
        return True
    return str(getattr(exc, "sqlstate", "") or "") == "40001"


import re as _re_mes

_MESES_ABBR = ["", "Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago",
               "Sep", "Oct", "Nov", "Dic"]


def _mes_de(fecha_str) -> int:
    """Mes (1-12) de una fecha en texto: 'YYYY-MM-DD' o 'DD/MM/AAAA'. 0 si no aplica."""
    s = str(fecha_str or "").strip()
    m = _re_mes.match(r"^\d{4}-(\d{2})-\d{2}", s)
    if m:
        return int(m.group(1))
    m = _re_mes.match(r"^\d{1,2}[/-](\d{1,2})[/-]\d{2,4}", s)
    if m:
        return int(m.group(1))
    return 0


def _fmt_meses(nums) -> str:
    ab = [_MESES_ABBR[n] for n in nums if 1 <= n <= 12]
    return ", ".join(ab) if ab else "—"


# ── Interfaz ─────────────────────────────────────────────────────────────────
class Persistencia(ABC):
    """Contrato mínimo. Cualquier motor (Firebird/MySQL/SQLite) lo implementa."""

    @abstractmethod
    def inicializar(self) -> None:
        """Crea la BD y el esquema si no existen. Idempotente."""

    @abstractmethod
    def siguiente_id(self) -> int:
        """Devuelve el próximo ID estable (monotónico, nunca reusado)."""

    @abstractmethod
    def guardar_conciliacion(self, rfc: str, periodo: str,
                             filas: list[FilaConciliacion]) -> int:
        """Crea una conciliación y guarda sus filas. Devuelve conciliacion_id.
        A cada fila sin id se le asigna uno del generador (in-place)."""

    @abstractmethod
    def leer_filas(self, conciliacion_id: int) -> list[FilaConciliacion]:
        """Lee las filas de una conciliación."""

    @abstractmethod
    def upsert_fila(self, fila: FilaConciliacion) -> None:
        """Inserta o actualiza una fila POR ID (no sobrescribe otras)."""

    @abstractmethod
    def cerrar(self) -> None:
        ...


# ── Utilidades de configuración Firebird ─────────────────────────────────────
def _detectar_fbclient() -> Optional[str]:
    """Busca un fbclient.dll de 64-bit utilizable. Prioriza System32 (FB5)."""
    candidatos = [
        r"C:\Windows\System32\fbclient.dll",
        r"C:\Program Files\Firebird\Firebird_5_0\fbclient.dll",
        r"C:\Program Files\Firebird\Firebird_5_0\bin\fbclient.dll",
        r"C:\Program Files\Firebird\Firebird_3_0\bin\fbclient.dll",
        r"C:\Program Files\Firebird\Firebird_2_5\bin\fbclient.dll",
    ]
    for c in candidatos:
        if os.path.exists(c):
            return c
    return None


def _ruta_bd_default() -> str:
    """BD local, fuera de OneDrive. Usa %PROGRAMDATA% (accesible al servicio)."""
    base = os.environ.get("PROGRAMDATA", r"C:\ProgramData")
    carpeta = os.path.join(base, "ConvertidorBancos")
    os.makedirs(carpeta, exist_ok=True)
    return os.path.join(carpeta, "conciliacion.fdb")


# ── Implementación Firebird ──────────────────────────────────────────────────
class FirebirdPersistencia(Persistencia):
    """Persistencia sobre Firebird (servidor local en 3050)."""

    def __init__(self, ruta_fdb: str | None = None, fbclient: str | None = None,
                 host: str | None = None, user: str | None = None,
                 password: str | None = None, port: int | None = None,
                 charset: str | None = None):
        # Los parámetros no dados se toman de la config del GUI (db_config.json).
        try:
            import db_config as _dc
            cfg = _dc.cargar()
        except Exception:
            cfg = {}
        self.ruta_fdb = ruta_fdb if ruta_fdb is not None else (cfg.get("database") or _ruta_bd_default())
        self.fbclient = fbclient if fbclient is not None else (cfg.get("fbclient") or _detectar_fbclient())
        self.host     = host      if host     is not None else (cfg.get("host") or "localhost")
        self.user     = user      if user     is not None else (cfg.get("user") or "SYSDBA")
        self.password = password  if password is not None else (cfg.get("password") or "masterkey")
        self.port     = int(port  if port     is not None else (cfg.get("port") or 3050))
        self.charset  = charset   if charset  is not None else (cfg.get("charset") or "UTF8")
        # Usuario que opera esta sesión (para la bitácora de conflictos). Lo puede
        # fijar la UI con `persist.usuario = "..."`.
        self.usuario  = os.environ.get("USERNAME", "") or ""
        self._con     = None

    # -- conexión ------------------------------------------------------------
    def _dsn(self) -> str:
        # host/port:ruta → modo servidor (más robusto y reparable con gfix/gbak)
        if not self.host:
            return self.ruta_fdb                       # embebido (sin host)
        if self.port and self.port != 0:
            return f"{self.host}/{self.port}:{self.ruta_fdb}"
        return f"{self.host}:{self.ruta_fdb}"

    def _forzar_tpb(self) -> None:
        """Fija el TPB por defecto de la conexión: read_committed + rec_version +
        NOWAIT. Lecturas frescas + choque de escritura que falla rápido (no cuelga).
        Se fuerza por conexión porque el servidor es FB 2.5 (sin ReadConsistency)."""
        try:
            from firebird.driver import TPB, Isolation
            tpb = TPB(isolation=Isolation.READ_COMMITTED_RECORD_VERSION,
                      lock_timeout=0).get_buffer()
            self._con.default_tpb = tpb
            try:
                self._con.main_transaction.default_tpb = tpb
            except Exception:
                pass
        except Exception:
            # Si el driver no soporta el TPB pedido, seguimos con el default: la
            # columna de versión sigue detectando el conflicto igual.
            pass

    def inicializar(self) -> None:
        import firebird.driver as fb
        if self.fbclient:
            fb.driver_config.fb_client_library.value = self.fbclient

        nueva = not os.path.exists(self.ruta_fdb)
        if nueva:
            self._con = fb.create_database(
                self._dsn(), user=self.user, password=self.password, charset=self.charset)
        else:
            self._con = fb.connect(
                self._dsn(), user=self.user, password=self.password, charset=self.charset)
        self._forzar_tpb()
        self._crear_esquema()

    def _existe_objeto(self, cur, nombre: str, tipo: str = "TABLE") -> bool:
        if tipo == "TABLE":
            cur.execute(
                "SELECT 1 FROM RDB$RELATIONS WHERE RDB$RELATION_NAME = ? "
                "AND RDB$VIEW_BLR IS NULL", [nombre.upper()])
        elif tipo == "INDEX":
            cur.execute(
                "SELECT 1 FROM RDB$INDICES WHERE RDB$INDEX_NAME = ?", [nombre.upper()])
        else:  # GENERATOR
            cur.execute(
                "SELECT 1 FROM RDB$GENERATORS WHERE RDB$GENERATOR_NAME = ?",
                [nombre.upper()])
        return cur.fetchone() is not None

    def _crear_esquema(self) -> None:
        cur = self._con.cursor()
        # Generadores (IDs estables, nunca reusados → dejan huecos)
        for gen in ("GEN_CONCILIACION_ID", "GEN_FILA_ID",
                    "GEN_CONVERSION_ID", "GEN_MOV_ID", "GEN_AUDIT_ID",
                    "GEN_MODIF_ID"):
            if not self._existe_objeto(cur, gen, "GENERATOR"):
                cur.execute(f"CREATE SEQUENCE {gen}")
        # Tabla de conciliaciones (una por corrida)
        if not self._existe_objeto(cur, "CONCILIACION"):
            cur.execute("""
                CREATE TABLE CONCILIACION (
                    ID        INTEGER      NOT NULL PRIMARY KEY,
                    RFC       VARCHAR(20),
                    PERIODO   VARCHAR(20),
                    CREADO    TIMESTAMP
                )""")
        # Tabla de filas conciliadas (ID estable = llave de sync con el Excel)
        if not self._existe_objeto(cur, "FILA"):
            cur.execute("""
                CREATE TABLE FILA (
                    ID              INTEGER       NOT NULL PRIMARY KEY,
                    CONCILIACION_ID INTEGER,
                    RFC             VARCHAR(20),
                    BANCO           VARCHAR(60),
                    FECHA           VARCHAR(10),
                    IMPORTE         DOUBLE PRECISION,
                    CONCEPTO        VARCHAR(500),
                    UUID_CFDI       VARCHAR(200),
                    UUID_REP        VARCHAR(200),
                    TIPO_MATCH      VARCHAR(40),
                    REVISAR         SMALLINT,
                    CUENTA          VARCHAR(40),
                    TERCERO         VARCHAR(200),
                    TIPO            VARCHAR(12),
                    UUID_SUMA       VARCHAR(2000),
                    DEPOSITO        DOUBLE PRECISION,
                    RETIRO          DOUBLE PRECISION,
                    VERSION_REG        INTEGER,
                    FECHA_MODIFICACION TIMESTAMP
                )""")
        # Tabla CFDI: todo lo que la póliza necesita, para regenerar SIN los XML
        if not self._existe_objeto(cur, "CFDI"):
            cur.execute("""
                CREATE TABLE CFDI (
                    UUID            VARCHAR(40)   NOT NULL PRIMARY KEY,
                    TIPO            VARCHAR(4),
                    RFC_EMISOR      VARCHAR(20),
                    RFC_RECEPTOR    VARCHAR(20),
                    NOMBRE_EMISOR   VARCHAR(300),
                    NOMBRE_RECEPTOR VARCHAR(300),
                    SUBTOTAL        DOUBLE PRECISION,
                    TOTAL           DOUBLE PRECISION,
                    DESCUENTO       DOUBLE PRECISION,
                    IVA             DOUBLE PRECISION,
                    IVA_RETENIDO    DOUBLE PRECISION,
                    ISR_RETENIDO    DOUBLE PRECISION,
                    CLAVE_PROD_SERV VARCHAR(20),
                    FECHA_EMISION   VARCHAR(30)
                )""")
        # Registro de CONVERSIONES PDF→Excel (una por estado de cuenta convertido)
        if not self._existe_objeto(cur, "CONVERSION"):
            cur.execute("""
                CREATE TABLE CONVERSION (
                    ID       INTEGER      NOT NULL PRIMARY KEY,
                    ARCHIVO  VARCHAR(400),
                    BANCO    VARCHAR(60),
                    RFC      VARCHAR(20),
                    PERIODO  VARCHAR(30),
                    CREADO   TIMESTAMP
                )""")
        # Movimientos bancarios convertidos (ID estable = llave de sync con el Excel)
        if not self._existe_objeto(cur, "MOV_BANCO"):
            cur.execute("""
                CREATE TABLE MOV_BANCO (
                    ID            INTEGER       NOT NULL PRIMARY KEY,
                    CONVERSION_ID INTEGER,
                    BANCO         VARCHAR(60),
                    FECHA         VARCHAR(20),
                    FOLIO         VARCHAR(60),
                    DESCRIPCION   VARCHAR(600),
                    DEPOSITO      DOUBLE PRECISION,
                    RETIRO        DOUBLE PRECISION,
                    SALDO         DOUBLE PRECISION,
                    VERSION_REG        INTEGER,
                    FECHA_MODIFICACION TIMESTAMP
                )""")
        # Catálogo CENTRAL de empresas (RFC → nombre + regímenes). Fuente única
        # compartida por todos los módulos (antes vivía en %APPDATA%\empresas.bin).
        if not self._existe_objeto(cur, "EMPRESA"):
            cur.execute("""
                CREATE TABLE EMPRESA (
                    RFC               VARCHAR(20)  NOT NULL PRIMARY KEY,
                    NOMBRE            VARCHAR(300),
                    REGIMENES         VARCHAR(400),
                    REGIMENES_NOMBRES VARCHAR(1000),
                    CREADO            TIMESTAMP,
                    MODIFICADO        TIMESTAMP
                )""")
        # Informacion PROCESADA de Obligaciones Fiscales (para reportes posteriores)
        if not self._existe_objeto(cur, "PRESENTACION"):
            cur.execute("""
                CREATE TABLE PRESENTACION (
                    RFC                VARCHAR(20)  NOT NULL,
                    ANIO               INTEGER      NOT NULL,
                    MES                INTEGER      NOT NULL,
                    OBLIG_ID           VARCHAR(60)  NOT NULL,
                    ARCHIVO            VARCHAR(500),
                    FECHA_PRESENTACION VARCHAR(20),
                    NUMERO_OPERACION   VARCHAR(60),
                    USUARIO            VARCHAR(80),
                    CREADO             TIMESTAMP,
                    PRIMARY KEY (RFC, ANIO, MES, OBLIG_ID)
                )""")
        # Resultado del analisis de Contribuciones/Saldos (IVA acumulado)
        if not self._existe_objeto(cur, "SALDO_IVA"):
            cur.execute("""
                CREATE TABLE SALDO_IVA (
                    RFC          VARCHAR(20)  NOT NULL,
                    ANIO         INTEGER      NOT NULL,
                    MES          INTEGER      NOT NULL,
                    OPERACION    VARCHAR(60),
                    FECHA        VARCHAR(30),
                    A_CARGO      DOUBLE PRECISION,
                    A_FAVOR      DOUBLE PRECISION,
                    ACREDITADO   DOUBLE PRECISION,
                    REMANENTE    DOUBLE PRECISION,
                    CREADO       TIMESTAMP,
                    PRIMARY KEY (RFC, ANIO, MES)
                )""")
        # Resultado del analisis de Contribuciones/Saldos (pagos)
        if not self._existe_objeto(cur, "SALDO_PAGO"):
            cur.execute("""
                CREATE TABLE SALDO_PAGO (
                    RFC            VARCHAR(20)  NOT NULL,
                    ANIO           INTEGER      NOT NULL,
                    MES            INTEGER      NOT NULL,
                    OPERACION      VARCHAR(60)  NOT NULL,
                    TIPO           VARCHAR(20),
                    VENCIMIENTO    VARCHAR(30),
                    CONTRIBUCION   DOUBLE PRECISION,
                    ACTUALIZACION  DOUBLE PRECISION,
                    RECARGOS       DOUBLE PRECISION,
                    IMPORTE_PAGAR  DOUBLE PRECISION,
                    IMPORTE_PAGADO DOUBLE PRECISION,
                    FECHA_PAGO     VARCHAR(30),
                    ESTATUS        VARCHAR(40),
                    CREADO         TIMESTAMP,
                    PRIMARY KEY (RFC, ANIO, MES, OPERACION)
                )""")
        # Bitácora de auditoría de conflictos de concurrencia (trazabilidad)
        if not self._existe_objeto(cur, "AUDITORIA_CONFLICTO"):
            cur.execute("""
                CREATE TABLE AUDITORIA_CONFLICTO (
                    ID               INTEGER      NOT NULL PRIMARY KEY,
                    TABLA            VARCHAR(40),
                    REGISTRO_ID      INTEGER,
                    USUARIO          VARCHAR(80),
                    DECISION         VARCHAR(20),
                    VERSION_LOCAL    VARCHAR(20),
                    VERSION_GUARDADA VARCHAR(20),
                    VALOR_LOCAL      VARCHAR(1000),
                    VALOR_GUARDADO   VARCHAR(1000),
                    CREADO           TIMESTAMP
                )""")
        # Bitacora de MODIFICACIONES de valores (monto/concepto) — tabla APARTE
        # para NO mezclar con la conciliacion original. La FILA original queda intacta
        # en sus montos/concepto; aqui se guarda el valor nuevo, el anterior y una
        # observacion, con su propio ID.
        if not self._existe_objeto(cur, "FILA_MODIFICACION"):
            cur.execute("""
                CREATE TABLE FILA_MODIFICACION (
                    ID           INTEGER      NOT NULL PRIMARY KEY,
                    ORIGEN_ID    INTEGER,
                    RFC          VARCHAR(20),
                    BANCO        VARCHAR(60),
                    FECHA        VARCHAR(10),
                    CONCEPTO     VARCHAR(500),
                    RETIRO       DOUBLE PRECISION,
                    DEPOSITO     DOUBLE PRECISION,
                    CONCEPTO_ANT VARCHAR(500),
                    RETIRO_ANT   DOUBLE PRECISION,
                    DEPOSITO_ANT DOUBLE PRECISION,
                    OBSERVACION  VARCHAR(500),
                    USUARIO      VARCHAR(80),
                    CREADO       TIMESTAMP
                )""")
        self._con.commit()
        # Para BDs creadas antes de estos campos: agregarlos si faltan (idempotente).
        self._asegurar_columnas(cur)
        self._con.commit()

    def _columnas(self, cur, tabla: str) -> set:
        cur.execute(
            "SELECT RDB$FIELD_NAME FROM RDB$RELATION_FIELDS WHERE RDB$RELATION_NAME = ?",
            [tabla.upper()])
        return {r[0].strip().upper() for r in cur.fetchall()}

    def _asegurar_columnas(self, cur) -> None:
        # FILA: columnas de negocio + control de concurrencia
        cols = self._columnas(cur, "FILA")
        faltan = {
            "TIPO":               "VARCHAR(12)",
            "UUID_SUMA":          "VARCHAR(2000)",
            "DEPOSITO":           "DOUBLE PRECISION",
            "RETIRO":             "DOUBLE PRECISION",
            "VERSION_REG":        "INTEGER",
            "FECHA_MODIFICACION": "TIMESTAMP",
        }
        for col, tipo in faltan.items():
            if col not in cols:
                cur.execute(f"ALTER TABLE FILA ADD {col} {tipo}")
        # MOV_BANCO: control de concurrencia
        if self._existe_objeto(cur, "MOV_BANCO"):
            cols_m = self._columnas(cur, "MOV_BANCO")
            for col, tipo in (("VERSION_REG", "INTEGER"),
                              ("FECHA_MODIFICACION", "TIMESTAMP")):
                if col not in cols_m:
                    cur.execute(f"ALTER TABLE MOV_BANCO ADD {col} {tipo}")
        self._con.commit()
        # Backfill: en filas migradas VERSION_REG queda NULL. El UPDATE optimista
        # compara `VERSION_REG = 0` y NULL nunca casa con 0 → daría un conflicto
        # falso. Igualamos NULL→0 (idempotente) para que la llave optimista funcione.
        for tabla in ("FILA", "MOV_BANCO"):
            try:
                if self._existe_objeto(cur, tabla):
                    cur.execute(f"UPDATE {tabla} SET VERSION_REG = 0 "
                                f"WHERE VERSION_REG IS NULL")
            except Exception:
                pass
        self._con.commit()
        # Índices sobre FECHA_MODIFICACION (checkpoint incremental). Requieren que
        # la columna ya exista; por eso van tras el commit del ALTER.
        for idx, tabla in (("IDX_FILA_FMOD", "FILA"),
                           ("IDX_MOV_FMOD", "MOV_BANCO")):
            try:
                if (self._existe_objeto(cur, tabla)
                        and not self._existe_objeto(cur, idx, "INDEX")):
                    cur.execute(f"CREATE INDEX {idx} ON {tabla} (FECHA_MODIFICACION)")
            except Exception:
                pass

    # -- utilidades de concurrencia -----------------------------------------
    @staticmethod
    def _ahora() -> datetime:
        return datetime.now()

    def _rollback_seguro(self) -> None:
        try:
            self._con.rollback()
        except Exception:
            pass

    # -- IDs estables --------------------------------------------------------
    def siguiente_id(self) -> int:
        cur = self._con.cursor()
        cur.execute("SELECT NEXT VALUE FOR GEN_FILA_ID FROM RDB$DATABASE")
        return int(cur.fetchone()[0])

    def _siguiente_conciliacion_id(self) -> int:
        cur = self._con.cursor()
        cur.execute("SELECT NEXT VALUE FOR GEN_CONCILIACION_ID FROM RDB$DATABASE")
        return int(cur.fetchone()[0])

    # -- guardar -------------------------------------------------------------
    def guardar_conciliacion(self, rfc: str, periodo: str,
                             filas: list[FilaConciliacion]) -> int:
        cur = self._con.cursor()
        conc_id = self._siguiente_conciliacion_id()
        cur.execute(
            "INSERT INTO CONCILIACION (ID, RFC, PERIODO, CREADO) VALUES (?, ?, ?, ?)",
            [conc_id, rfc, periodo, datetime.now()])
        for fila in filas:
            if not fila.id:
                fila.id = self.siguiente_id()
            fila.conciliacion_id = conc_id
            self._insert_fila(cur, fila)
        self._con.commit()
        return conc_id

    _FILA_COLS = ("ID", "CONCILIACION_ID", "RFC", "BANCO", "FECHA", "IMPORTE",
                  "CONCEPTO", "UUID_CFDI", "UUID_REP", "TIPO_MATCH", "REVISAR",
                  "CUENTA", "TERCERO", "TIPO", "UUID_SUMA", "DEPOSITO", "RETIRO")

    def _fila_values(self, f: FilaConciliacion) -> list:
        # Recorte defensivo al ancho de CADA columna VARCHAR de FILA: evita el
        # error Firebird -303 (string truncation). En especial FECHA es VARCHAR(10)
        # y el Excel puede entregar un datetime ('2026-08-20 00:00:00', 19 chars)
        # cuando la celda quedo como fecha real; [:10] deja la fecha ISO correcta.
        return [f.id, f.conciliacion_id, (f.rfc or "")[:20], (f.banco or "")[:60],
                (f.fecha or "")[:10], f.importe,
                (f.concepto or "")[:500], (f.uuid_cfdi or "")[:200], (f.uuid_rep or "")[:200],
                (f.tipo_match or "")[:40], 1 if f.revisar else 0, (f.cuenta or "")[:40],
                (f.tercero or "")[:200], (f.tipo or "")[:12], (f.uuid_suma or "")[:2000],
                f.deposito, f.retiro]

    def _insert_fila(self, cur, f: FilaConciliacion) -> None:
        # Inserta las columnas de negocio + versión inicial (1) y fecha de cambio.
        ph = ",".join("?" * len(self._FILA_COLS))
        cols = self._FILA_COLS + ("VERSION_REG", "FECHA_MODIFICACION")
        ahora = self._ahora()
        cur.execute(
            f"INSERT INTO FILA ({', '.join(cols)}) VALUES ({ph},?,?)",
            self._fila_values(f) + [1, ahora])
        f.version_reg = 1
        f.fecha_modificacion = ahora.isoformat(sep=" ", timespec="seconds")

    # -- leer / upsert -------------------------------------------------------
    def _row_to_fila(self, r) -> FilaConciliacion:
        # _FILA_SEL = 17 cols de negocio (0..16) + VERSION_REG(17) + FECHA_MODIFICACION(18)
        ver = int(r[17] or 0) if len(r) > 17 else 0
        fm  = r[18] if len(r) > 18 and r[18] else None
        return FilaConciliacion(
            id=r[0], conciliacion_id=r[1], rfc=r[2] or "", banco=r[3] or "",
            fecha=r[4] or "", importe=r[5] or 0.0, concepto=r[6] or "",
            uuid_cfdi=r[7] or "", uuid_rep=r[8] or "", tipo_match=r[9] or "",
            revisar=bool(r[10]), cuenta=r[11] or "", tercero=r[12] or "",
            tipo=r[13] or "", uuid_suma=r[14] or "",
            deposito=r[15] or 0.0, retiro=r[16] or 0.0,
            version_reg=ver, fecha_modificacion=(str(fm) if fm else ""))

    _FILA_SEL = ", ".join(_FILA_COLS) + ", VERSION_REG, FECHA_MODIFICACION"

    def leer_filas(self, conciliacion_id: int) -> list[FilaConciliacion]:
        cur = self._con.cursor()
        cur.execute(
            f"SELECT {self._FILA_SEL} FROM FILA WHERE CONCILIACION_ID = ? "
            "ORDER BY ID", [conciliacion_id])
        return [self._row_to_fila(r) for r in cur.fetchall()]

    def leer_filas_por_rfc(self, rfc: str) -> list[FilaConciliacion]:
        """Todas las filas conciliadas de una empresa (para REIMPRIMIR desde la BD
        sin depender del Excel). Ordenadas por ID."""
        cur = self._con.cursor()
        cur.execute(
            f"SELECT {self._FILA_SEL} FROM FILA WHERE RFC = ? ORDER BY ID",
            [(rfc or "").strip().upper()])
        return [self._row_to_fila(r) for r in cur.fetchall()]

    def leer_conciliaciones(self, rfc: str) -> list:
        """Lista las conciliaciones guardadas de una empresa (mas recientes primero),
        con conteo de filas y de conciliadas (con CFDI). Para la consulta del GUI."""
        cur = self._con.cursor()
        cur.execute("SELECT ID, RFC, PERIODO, CREADO FROM CONCILIACION WHERE RFC = ? "
                    "ORDER BY ID DESC", [(rfc or "").strip().upper()])
        concs = [dict(zip(("id", "rfc", "periodo", "creado"), r))
                 for r in cur.fetchall()]
        for c in concs:
            cur.execute("SELECT COUNT(*), "
                        "SUM(CASE WHEN UUID_CFDI IS NOT NULL AND UUID_CFDI <> '' "
                        "THEN 1 ELSE 0 END) FROM FILA WHERE CONCILIACION_ID = ?",
                        [c["id"]])
            r = cur.fetchone()
            c["n_filas"] = int(r[0] or 0)
            c["n_conciliadas"] = int((r[1] or 0))
            cur.execute("SELECT DISTINCT FECHA FROM FILA WHERE CONCILIACION_ID = ?",
                        [c["id"]])
            meses = sorted({_mes_de(row[0]) for row in cur.fetchall()} - {0})
            c["meses"] = _fmt_meses(meses)
        return concs

    def eliminar_conciliacion(self, conciliacion_id: int) -> int:
        """Borra una conciliacion guardada con todas sus filas y su historial de
        modificaciones (no hay FK con cascade, se borra en orden). Devuelve el
        numero de filas borradas."""
        cur = self._con.cursor()
        cid = int(conciliacion_id)
        cur.execute("SELECT COUNT(*) FROM FILA WHERE CONCILIACION_ID = ?", [cid])
        n = int((cur.fetchone() or [0])[0] or 0)
        cur.execute("DELETE FROM FILA_MODIFICACION WHERE ORIGEN_ID IN "
                    "(SELECT ID FROM FILA WHERE CONCILIACION_ID = ?)", [cid])
        cur.execute("DELETE FROM FILA WHERE CONCILIACION_ID = ?", [cid])
        cur.execute("DELETE FROM CONCILIACION WHERE ID = ?", [cid])
        self._con.commit()
        return n

    def leer_fila(self, fila_id: int) -> Optional[FilaConciliacion]:
        cur = self._con.cursor()
        cur.execute(
            f"SELECT {self._FILA_SEL} FROM FILA WHERE ID = ?", [fila_id])
        r = cur.fetchone()
        return self._row_to_fila(r) if r else None

    def upsert_fila(self, f: FilaConciliacion, version_esperada=None,
                    usuario: str | None = None) -> str:
        """Inserta o actualiza una fila POR ID con control de concurrencia.

        - Nueva → INSERT (versión 1).
        - Existente → UPDATE condicional por versión. Si ``version_esperada`` se
          da y ya no coincide con la versión guardada, o si la carrera se detecta
          en el propio UPDATE, se lanza ``ConflictoConcurrencia`` con ambos valores.

        Devuelve 'nuevo' | 'actualizado'.
        """
        import firebird.driver as fb
        cur = self._con.cursor()
        cur.execute("SELECT VERSION_REG FROM FILA WHERE ID = ?", [f.id])
        row = cur.fetchone()
        if row is None:
            self._insert_fila(cur, f)
            self._con.commit()
            return "nuevo"

        version_actual = int(row[0] or 0)
        # 1) Chequeo optimista explícito (el usuario cargó una versión y ya cambió)
        if version_esperada is not None and int(version_esperada) != version_actual:
            self._rollback_seguro()
            guardado = self.leer_fila(f.id)
            raise ConflictoConcurrencia("FILA", f.id, f, guardado,
                                        version_esperada, version_actual)

        # 2) UPDATE condicional por versión (defensa ante carrera simultánea)
        sets = ", ".join(f"{c}=?" for c in self._FILA_COLS if c != "ID")
        vals = [v for c, v in zip(self._FILA_COLS, self._fila_values(f)) if c != "ID"]
        nueva_ver = version_actual + 1
        ahora = self._ahora()
        try:
            cur.execute(
                f"UPDATE FILA SET {sets}, VERSION_REG=?, FECHA_MODIFICACION=? "
                f"WHERE ID=? AND VERSION_REG=?",
                vals + [nueva_ver, ahora, f.id, version_actual])
        except fb.DatabaseError as e:
            self._rollback_seguro()
            if _es_update_conflict(e):
                guardado = self.leer_fila(f.id)
                raise ConflictoConcurrencia("FILA", f.id, f, guardado,
                                            version_actual, "concurrente")
            raise
        if cur.rowcount == 0:
            # Alguien cambió la versión entre el SELECT y el UPDATE.
            self._rollback_seguro()
            guardado = self.leer_fila(f.id)
            raise ConflictoConcurrencia("FILA", f.id, f, guardado,
                                        version_actual, "cambió")
        self._con.commit()
        f.version_reg = nueva_ver
        f.fecha_modificacion = ahora.isoformat(sep=" ", timespec="seconds")
        return "actualizado"

    # -- CFDI (autosuficiencia para pólizas sin los XML) --------------------
    _CFDI_COLS = ("UUID", "TIPO", "RFC_EMISOR", "RFC_RECEPTOR", "NOMBRE_EMISOR",
                  "NOMBRE_RECEPTOR", "SUBTOTAL", "TOTAL", "DESCUENTO", "IVA",
                  "IVA_RETENIDO", "ISR_RETENIDO", "CLAVE_PROD_SERV", "FECHA_EMISION")

    def guardar_cfdis(self, cfdis: list[CFDIData], solo_nuevos: bool = False) -> dict:
        """Guarda CFDIs por UUID. Con `solo_nuevos=True` OMITE los UUID que ya
        existen (no los actualiza) — para "subir XML" sin duplicar; con False es
        upsert idempotente. Devuelve {nuevos, actualizados, omitidos}."""
        cur = self._con.cursor()
        ph = ",".join("?" * len(self._CFDI_COLS))
        sets = ", ".join(f"{c}=?" for c in self._CFDI_COLS if c != "UUID")
        stats = dict(nuevos=0, actualizados=0, omitidos=0)
        for c in cfdis:
            if not c.uuid:
                continue
            vals = [c.uuid[:40], c.tipo[:4], c.rfc_emisor[:20], c.rfc_receptor[:20],
                    c.nombre_emisor[:300], c.nombre_receptor[:300], c.subtotal,
                    c.total, c.descuento, c.iva, c.iva_retenido, c.isr_retenido,
                    c.clave_prod_serv[:20], c.fecha_emision[:30]]
            cur.execute("SELECT 1 FROM CFDI WHERE UUID = ?", [c.uuid[:40]])
            if cur.fetchone() is not None:
                if solo_nuevos:
                    stats["omitidos"] += 1
                    continue
                cur.execute(f"UPDATE CFDI SET {sets} WHERE UUID=?", vals[1:] + [c.uuid[:40]])
                stats["actualizados"] += 1
            else:
                cur.execute(f"INSERT INTO CFDI ({', '.join(self._CFDI_COLS)}) VALUES ({ph})", vals)
                stats["nuevos"] += 1
        self._con.commit()
        return stats

    def obtener_cfdi(self, uuid: str) -> Optional[CFDIData]:
        cur = self._con.cursor()
        cur.execute(
            f"SELECT {', '.join(self._CFDI_COLS)} FROM CFDI WHERE UUID = ?",
            [(uuid or "")[:40]])
        r = cur.fetchone()
        if not r:
            return None
        return CFDIData(
            uuid=r[0] or "", tipo=r[1] or "", rfc_emisor=r[2] or "",
            rfc_receptor=r[3] or "", nombre_emisor=r[4] or "",
            nombre_receptor=r[5] or "", subtotal=r[6] or 0.0, total=r[7] or 0.0,
            descuento=r[8] or 0.0, iva=r[9] or 0.0, iva_retenido=r[10] or 0.0,
            isr_retenido=r[11] or 0.0, clave_prod_serv=r[12] or "",
            fecha_emision=r[13] or "")

    def leer_todos_cfdis(self) -> list[CFDIData]:
        cur = self._con.cursor()
        cur.execute(f"SELECT {', '.join(self._CFDI_COLS)} FROM CFDI")
        out = []
        for r in cur.fetchall():
            out.append(CFDIData(
                uuid=r[0] or "", tipo=r[1] or "", rfc_emisor=r[2] or "",
                rfc_receptor=r[3] or "", nombre_emisor=r[4] or "",
                nombre_receptor=r[5] or "", subtotal=r[6] or 0.0, total=r[7] or 0.0,
                descuento=r[8] or 0.0, iva=r[9] or 0.0, iva_retenido=r[10] or 0.0,
                isr_retenido=r[11] or 0.0, clave_prod_serv=r[12] or "",
                fecha_emision=r[13] or ""))
        return out

    # -- Conversiones PDF→Excel (registro de movimientos convertidos) --------
    _MOV_COLS = ("ID", "CONVERSION_ID", "BANCO", "FECHA", "FOLIO",
                 "DESCRIPCION", "DEPOSITO", "RETIRO", "SALDO")

    def _mov_values(self, m: "MovimientoBanco") -> list:
        return [m.id, m.conversion_id, (m.banco or "")[:60], (m.fecha or "")[:20],
                (m.folio or "")[:60], (m.descripcion or "")[:600],
                float(m.deposito or 0), float(m.retiro or 0), float(m.saldo or 0)]

    def _row_to_mov(self, r) -> "MovimientoBanco":
        # _MOV_SEL = 9 cols (0..8) + VERSION_REG(9) + FECHA_MODIFICACION(10)
        ver = int(r[9] or 0) if len(r) > 9 else 0
        fm  = r[10] if len(r) > 10 and r[10] else None
        return MovimientoBanco(
            id=r[0], conversion_id=r[1], banco=r[2] or "", fecha=r[3] or "",
            folio=r[4] or "", descripcion=r[5] or "", deposito=r[6] or 0.0,
            retiro=r[7] or 0.0, saldo=r[8] or 0.0,
            version_reg=ver, fecha_modificacion=(str(fm) if fm else ""))

    _MOV_SEL = ", ".join(_MOV_COLS) + ", VERSION_REG, FECHA_MODIFICACION"

    def _insert_mov(self, cur, m: "MovimientoBanco") -> None:
        ph = ",".join("?" * len(self._MOV_COLS))
        cols = self._MOV_COLS + ("VERSION_REG", "FECHA_MODIFICACION")
        ahora = self._ahora()
        cur.execute(
            f"INSERT INTO MOV_BANCO ({', '.join(cols)}) VALUES ({ph},?,?)",
            self._mov_values(m) + [1, ahora])
        m.version_reg = 1
        m.fecha_modificacion = ahora.isoformat(sep=" ", timespec="seconds")

    def guardar_conversion(self, archivo: str, banco: str, rfc: str, periodo: str,
                           movs: list["MovimientoBanco"]) -> int:
        """Crea una conversión y guarda sus movimientos con ID estable (in-place)."""
        cur = self._con.cursor()
        cur.execute("SELECT NEXT VALUE FOR GEN_CONVERSION_ID FROM RDB$DATABASE")
        conv_id = int(cur.fetchone()[0])
        cur.execute(
            "INSERT INTO CONVERSION (ID, ARCHIVO, BANCO, RFC, PERIODO, CREADO) "
            "VALUES (?,?,?,?,?,?)",
            [conv_id, (archivo or "")[:400], (banco or "")[:60], (rfc or "")[:20],
             (periodo or "")[:30], datetime.now()])
        for m in movs:
            if not m.id:
                cur.execute("SELECT NEXT VALUE FOR GEN_MOV_ID FROM RDB$DATABASE")
                m.id = int(cur.fetchone()[0])
            m.conversion_id = conv_id
            self._insert_mov(cur, m)
        self._con.commit()
        return conv_id

    def leer_movimiento(self, mov_id: int) -> Optional["MovimientoBanco"]:
        cur = self._con.cursor()
        cur.execute(
            f"SELECT {self._MOV_SEL} FROM MOV_BANCO WHERE ID = ?", [mov_id])
        r = cur.fetchone()
        return self._row_to_mov(r) if r else None

    def leer_movimientos(self, conversion_id: int) -> list["MovimientoBanco"]:
        cur = self._con.cursor()
        cur.execute(
            f"SELECT {self._MOV_SEL} FROM MOV_BANCO "
            "WHERE CONVERSION_ID = ? ORDER BY ID", [conversion_id])
        return [self._row_to_mov(r) for r in cur.fetchall()]

    def upsert_movimiento_si_cambio(self, m: "MovimientoBanco",
                                    version_esperada=None) -> str:
        """Inserta o actualiza un movimiento POR ID, pero SOLO escribe si el
        contenido cambió. Con control de concurrencia por versión: si otro proceso
        modificó el movimiento (versión distinta), lanza ``ConflictoConcurrencia``.
        Devuelve 'nuevo' | 'actualizado' | 'sin_cambio'."""
        import firebird.driver as fb
        prev = self.leer_movimiento(m.id) if m.id else None
        if prev is None:
            cur = self._con.cursor()
            if not m.id:
                cur.execute("SELECT NEXT VALUE FOR GEN_MOV_ID FROM RDB$DATABASE")
                m.id = int(cur.fetchone()[0])
            self._insert_mov(cur, m)
            self._con.commit()
            return "nuevo"
        # Comparar campos editables (no conversion_id)
        def _clave(x):
            return (x.banco, x.fecha, x.folio, x.descripcion,
                    round(x.deposito or 0, 2), round(x.retiro or 0, 2),
                    round(x.saldo or 0, 2))
        if _clave(prev) == _clave(m):
            return "sin_cambio"

        version_actual = int(prev.version_reg or 0)
        if version_esperada is not None and int(version_esperada) != version_actual:
            self._rollback_seguro()
            guardado = self.leer_movimiento(m.id)
            raise ConflictoConcurrencia("MOV_BANCO", m.id, m, guardado,
                                        version_esperada, version_actual)

        m.conversion_id = m.conversion_id or prev.conversion_id
        cur = self._con.cursor()
        sets = ", ".join(f"{c}=?" for c in self._MOV_COLS if c != "ID")
        vals = [v for c, v in zip(self._MOV_COLS, self._mov_values(m)) if c != "ID"]
        nueva_ver = version_actual + 1
        ahora = self._ahora()
        try:
            cur.execute(
                f"UPDATE MOV_BANCO SET {sets}, VERSION_REG=?, FECHA_MODIFICACION=? "
                f"WHERE ID=? AND VERSION_REG=?",
                vals + [nueva_ver, ahora, m.id, version_actual])
        except fb.DatabaseError as e:
            self._rollback_seguro()
            if _es_update_conflict(e):
                guardado = self.leer_movimiento(m.id)
                raise ConflictoConcurrencia("MOV_BANCO", m.id, m, guardado,
                                            version_actual, "concurrente")
            raise
        if cur.rowcount == 0:
            self._rollback_seguro()
            guardado = self.leer_movimiento(m.id)
            raise ConflictoConcurrencia("MOV_BANCO", m.id, m, guardado,
                                        version_actual, "cambió")
        self._con.commit()
        m.version_reg = nueva_ver
        m.fecha_modificacion = ahora.isoformat(sep=" ", timespec="seconds")
        return "actualizado"

    # -- Bitácora de auditoría de conflictos --------------------------------
    def registrar_conflicto(self, tabla: str, registro_id: int, decision: str,
                            version_local, version_guardada,
                            valor_local: str = "", valor_guardado: str = "",
                            usuario: str | None = None) -> int:
        """Deja constancia de un conflicto de concurrencia y de cómo se resolvió
        (quién lo intentó, qué decisión ganó, cuándo). Devuelve el ID de auditoría."""
        cur = self._con.cursor()
        cur.execute("SELECT NEXT VALUE FOR GEN_AUDIT_ID FROM RDB$DATABASE")
        aid = int(cur.fetchone()[0])
        cur.execute(
            "INSERT INTO AUDITORIA_CONFLICTO (ID, TABLA, REGISTRO_ID, USUARIO, "
            "DECISION, VERSION_LOCAL, VERSION_GUARDADA, VALOR_LOCAL, VALOR_GUARDADO, "
            "CREADO) VALUES (?,?,?,?,?,?,?,?,?,?)",
            [aid, (tabla or "")[:40], int(registro_id or 0),
             (usuario or self.usuario or "")[:80], (decision or "")[:20],
             str(version_local)[:20], str(version_guardada)[:20],
             (valor_local or "")[:1000], (valor_guardado or "")[:1000],
             datetime.now()])
        self._con.commit()
        return aid

    def leer_auditoria(self, limite: int = 200) -> list[dict]:
        """Lee la bitácora de conflictos (más recientes primero)."""
        cur = self._con.cursor()
        cur.execute(
            "SELECT FIRST ? ID, TABLA, REGISTRO_ID, USUARIO, DECISION, "
            "VERSION_LOCAL, VERSION_GUARDADA, VALOR_LOCAL, VALOR_GUARDADO, CREADO "
            "FROM AUDITORIA_CONFLICTO ORDER BY ID DESC", [int(limite)])
        campos = ("id", "tabla", "registro_id", "usuario", "decision",
                  "version_local", "version_guardada", "valor_local",
                  "valor_guardado", "creado")
        return [dict(zip(campos, r)) for r in cur.fetchall()]

    # -- Catálogo central de empresas ---------------------------------------
    def guardar_empresa(self, rfc: str, nombre: str, regimenes=None,
                        regimenes_nombres=None) -> None:
        """Upsert por RFC. regimenes/regimenes_nombres: listas (se guardan como CSV).
        Si se pasan None, conserva los que ya tenga la fila."""
        rfc = (rfc or "").strip().upper()
        if not rfc:
            return
        cur = self._con.cursor()
        cur.execute("SELECT REGIMENES, REGIMENES_NOMBRES, CREADO FROM EMPRESA WHERE RFC=?",
                    [rfc])
        row = cur.fetchone()
        def _csv(v, prev):
            if v is None:
                return prev or ""
            return ",".join(str(x) for x in v)
        ahora = datetime.now()
        if row is None:
            cur.execute(
                "INSERT INTO EMPRESA (RFC, NOMBRE, REGIMENES, REGIMENES_NOMBRES, "
                "CREADO, MODIFICADO) VALUES (?,?,?,?,?,?)",
                [rfc, (nombre or "")[:300], _csv(regimenes, "")[:400],
                 _csv(regimenes_nombres, "")[:1000], ahora, ahora])
        else:
            cur.execute(
                "UPDATE EMPRESA SET NOMBRE=?, REGIMENES=?, REGIMENES_NOMBRES=?, "
                "MODIFICADO=? WHERE RFC=?",
                [(nombre or "")[:300], _csv(regimenes, row[0])[:400],
                 _csv(regimenes_nombres, row[1])[:1000], ahora, rfc])
        self._con.commit()

    def leer_empresa(self, rfc: str) -> Optional[dict]:
        rfc = (rfc or "").strip().upper()
        if not rfc:
            return None
        cur = self._con.cursor()
        cur.execute("SELECT RFC, NOMBRE, REGIMENES, REGIMENES_NOMBRES FROM EMPRESA "
                    "WHERE RFC=?", [rfc])
        r = cur.fetchone()
        if not r:
            return None
        return {"rfc": r[0], "nombre": r[1] or "",
                "regimenes": [x for x in (r[2] or "").split(",") if x],
                "regimenes_nombres": [x for x in (r[3] or "").split(",") if x]}

    def leer_empresas(self) -> list[dict]:
        cur = self._con.cursor()
        cur.execute("SELECT RFC, NOMBRE, REGIMENES, REGIMENES_NOMBRES FROM EMPRESA")
        out = []
        for r in cur.fetchall():
            out.append({"rfc": r[0], "nombre": r[1] or "",
                        "regimenes": [x for x in (r[2] or "").split(",") if x],
                        "regimenes_nombres": [x for x in (r[3] or "").split(",") if x]})
        return out

    def eliminar_empresa(self, rfc: str) -> bool:
        rfc = (rfc or "").strip().upper()
        cur = self._con.cursor()
        cur.execute("DELETE FROM EMPRESA WHERE RFC=?", [rfc])
        n = cur.rowcount
        self._con.commit()
        return bool(n)

    # -- Bitacora de modificaciones de valores (monto/concepto) --------------
    def registrar_modificacion(self, origen_id, rfc, banco, fecha, concepto,
                               retiro, deposito, concepto_ant, retiro_ant,
                               deposito_ant, observacion="", usuario="") -> int:
        """Guarda en FILA_MODIFICACION un cambio de valor detectado al subir el Excel.
        NO toca la FILA original. Devuelve el ID nuevo de la modificacion."""
        cur = self._con.cursor()
        cur.execute("SELECT NEXT VALUE FOR GEN_MODIF_ID FROM RDB$DATABASE")
        mid = int(cur.fetchone()[0])
        cur.execute(
            "INSERT INTO FILA_MODIFICACION (ID, ORIGEN_ID, RFC, BANCO, FECHA, "
            "CONCEPTO, RETIRO, DEPOSITO, CONCEPTO_ANT, RETIRO_ANT, DEPOSITO_ANT, "
            "OBSERVACION, USUARIO, CREADO) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            [mid, int(origen_id or 0), (rfc or "")[:20], (banco or "")[:60],
             (fecha or "")[:10], (concepto or "")[:500], float(retiro or 0),
             float(deposito or 0), (concepto_ant or "")[:500], float(retiro_ant or 0),
             float(deposito_ant or 0), (observacion or "")[:500], (usuario or "")[:80],
             datetime.now()])
        self._con.commit()
        return mid

    def leer_modificaciones(self, rfc, origen_id=None) -> list:
        """Modificaciones registradas de una empresa (o de una fila origen)."""
        cur = self._con.cursor()
        if origen_id is not None:
            cur.execute(
                "SELECT ID, ORIGEN_ID, BANCO, FECHA, CONCEPTO, RETIRO, DEPOSITO, "
                "CONCEPTO_ANT, RETIRO_ANT, DEPOSITO_ANT, OBSERVACION, USUARIO, CREADO "
                "FROM FILA_MODIFICACION WHERE RFC=? AND ORIGEN_ID=? ORDER BY ID DESC",
                [(rfc or "").strip().upper(), int(origen_id)])
        else:
            cur.execute(
                "SELECT ID, ORIGEN_ID, BANCO, FECHA, CONCEPTO, RETIRO, DEPOSITO, "
                "CONCEPTO_ANT, RETIRO_ANT, DEPOSITO_ANT, OBSERVACION, USUARIO, CREADO "
                "FROM FILA_MODIFICACION WHERE RFC=? ORDER BY ID DESC",
                [(rfc or "").strip().upper()])
        campos = ("id", "origen_id", "banco", "fecha", "concepto", "retiro",
                  "deposito", "concepto_ant", "retiro_ant", "deposito_ant",
                  "observacion", "usuario", "creado")
        return [dict(zip(campos, r)) for r in cur.fetchall()]

    def cerrar(self) -> None:
        if self._con is not None:
            try:
                self._con.close()
            except Exception:
                pass
            self._con = None


    # -- Informacion procesada de Obligaciones (para reportes) ---------------
    def guardar_presentacion(self, rfc, anio, mes, oblig_id, archivo="",
                             fecha_presentacion="", numero_operacion="",
                             usuario="") -> None:
        """Upsert de una obligacion presentada (espejo en FDB de compliance.json)."""
        rfc = (rfc or "").strip().upper()
        cur = self._con.cursor()
        cur.execute("SELECT 1 FROM PRESENTACION WHERE RFC=? AND ANIO=? AND MES=? "
                    "AND OBLIG_ID=?", [rfc, int(anio), int(mes), oblig_id])
        vals = [(archivo or "")[:500], (fecha_presentacion or "")[:20],
                (numero_operacion or "")[:60], (usuario or "")[:80], datetime.now()]
        if cur.fetchone():
            cur.execute("UPDATE PRESENTACION SET ARCHIVO=?, FECHA_PRESENTACION=?, "
                        "NUMERO_OPERACION=?, USUARIO=?, CREADO=? WHERE RFC=? AND "
                        "ANIO=? AND MES=? AND OBLIG_ID=?",
                        vals + [rfc, int(anio), int(mes), oblig_id])
        else:
            cur.execute("INSERT INTO PRESENTACION (ARCHIVO, FECHA_PRESENTACION, "
                        "NUMERO_OPERACION, USUARIO, CREADO, RFC, ANIO, MES, OBLIG_ID) "
                        "VALUES (?,?,?,?,?,?,?,?,?)",
                        vals + [rfc, int(anio), int(mes), oblig_id])
        self._con.commit()

    def eliminar_presentacion(self, rfc, anio, mes, oblig_id) -> bool:
        cur = self._con.cursor()
        cur.execute("DELETE FROM PRESENTACION WHERE RFC=? AND ANIO=? AND MES=? AND "
                    "OBLIG_ID=?", [(rfc or "").strip().upper(), int(anio), int(mes),
                                   oblig_id])
        n = cur.rowcount
        self._con.commit()
        return bool(n)

    def leer_presentaciones(self, rfc, anio, mes) -> list:
        cur = self._con.cursor()
        cur.execute("SELECT OBLIG_ID, ARCHIVO, FECHA_PRESENTACION, NUMERO_OPERACION, "
                    "USUARIO, CREADO FROM PRESENTACION WHERE RFC=? AND ANIO=? AND MES=? "
                    "ORDER BY OBLIG_ID",
                    [(rfc or "").strip().upper(), int(anio), int(mes)])
        campos = ("oblig_id", "archivo", "fecha_presentacion", "numero_operacion",
                  "usuario", "creado")
        return [dict(zip(campos, r)) for r in cur.fetchall()]

    # -- Resultado procesado de Contribuciones/Saldos -------------------------
    def guardar_saldo_iva(self, rfc, anio, mes, operacion="", fecha="", a_cargo=0.0,
                          a_favor=0.0, acreditado=0.0, remanente=0.0) -> None:
        rfc = (rfc or "").strip().upper()
        cur = self._con.cursor()
        cur.execute("SELECT 1 FROM SALDO_IVA WHERE RFC=? AND ANIO=? AND MES=?",
                    [rfc, int(anio), int(mes)])
        vals = [(operacion or "")[:60], (fecha or "")[:30], float(a_cargo or 0),
                float(a_favor or 0), float(acreditado or 0), float(remanente or 0),
                datetime.now()]
        if cur.fetchone():
            cur.execute("UPDATE SALDO_IVA SET OPERACION=?, FECHA=?, A_CARGO=?, "
                        "A_FAVOR=?, ACREDITADO=?, REMANENTE=?, CREADO=? "
                        "WHERE RFC=? AND ANIO=? AND MES=?",
                        vals + [rfc, int(anio), int(mes)])
        else:
            cur.execute("INSERT INTO SALDO_IVA (OPERACION, FECHA, A_CARGO, A_FAVOR, "
                        "ACREDITADO, REMANENTE, CREADO, RFC, ANIO, MES) "
                        "VALUES (?,?,?,?,?,?,?,?,?,?)", vals + [rfc, int(anio), int(mes)])
        self._con.commit()

    def guardar_saldo_pago(self, rfc, anio, mes, operacion, tipo="", vencimiento="",
                           contribucion=0.0, actualizacion=0.0, recargos=0.0,
                           importe_pagar=0.0, importe_pagado=0.0, fecha_pago="",
                           estatus="") -> None:
        rfc = (rfc or "").strip().upper()
        operacion = (operacion or "")[:60] or "-"
        cur = self._con.cursor()
        cur.execute("SELECT 1 FROM SALDO_PAGO WHERE RFC=? AND ANIO=? AND MES=? AND "
                    "OPERACION=?", [rfc, int(anio), int(mes), operacion])
        vals = [(tipo or "")[:20], (vencimiento or "")[:30], float(contribucion or 0),
                float(actualizacion or 0), float(recargos or 0), float(importe_pagar or 0),
                float(importe_pagado or 0), (fecha_pago or "")[:30], (estatus or "")[:40],
                datetime.now()]
        if cur.fetchone():
            cur.execute("UPDATE SALDO_PAGO SET TIPO=?, VENCIMIENTO=?, CONTRIBUCION=?, "
                        "ACTUALIZACION=?, RECARGOS=?, IMPORTE_PAGAR=?, IMPORTE_PAGADO=?, "
                        "FECHA_PAGO=?, ESTATUS=?, CREADO=? WHERE RFC=? AND ANIO=? AND "
                        "MES=? AND OPERACION=?",
                        vals + [rfc, int(anio), int(mes), operacion])
        else:
            cur.execute("INSERT INTO SALDO_PAGO (TIPO, VENCIMIENTO, CONTRIBUCION, "
                        "ACTUALIZACION, RECARGOS, IMPORTE_PAGAR, IMPORTE_PAGADO, "
                        "FECHA_PAGO, ESTATUS, CREADO, RFC, ANIO, MES, OPERACION) "
                        "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                        vals + [rfc, int(anio), int(mes), operacion])
        self._con.commit()

    def leer_saldos_iva(self, rfc, anio, mes=None) -> list:
        cur = self._con.cursor()
        sql = ("SELECT ANIO, MES, OPERACION, FECHA, A_CARGO, A_FAVOR, ACREDITADO, "
               "REMANENTE FROM SALDO_IVA WHERE RFC=? AND ANIO=?")
        par = [(rfc or "").strip().upper(), int(anio)]
        if mes:
            sql += " AND MES=?"
            par.append(int(mes))
        cur.execute(sql + " ORDER BY ANIO, MES", par)
        campos = ("anio", "mes", "operacion", "fecha", "a_cargo", "a_favor",
                  "acreditado", "remanente")
        return [dict(zip(campos, r)) for r in cur.fetchall()]

    def leer_saldos_pago(self, rfc, anio, mes=None) -> list:
        cur = self._con.cursor()
        sql = ("SELECT ANIO, MES, OPERACION, TIPO, VENCIMIENTO, CONTRIBUCION, "
               "ACTUALIZACION, RECARGOS, IMPORTE_PAGAR, IMPORTE_PAGADO, FECHA_PAGO, "
               "ESTATUS FROM SALDO_PAGO WHERE RFC=? AND ANIO=?")
        par = [(rfc or "").strip().upper(), int(anio)]
        if mes:
            sql += " AND MES=?"
            par.append(int(mes))
        cur.execute(sql + " ORDER BY ANIO, MES", par)
        campos = ("anio", "mes", "operacion", "tipo", "vencimiento", "contribucion",
                  "actualizacion", "recargos", "importe_pagar", "importe_pagado",
                  "fecha_pago", "estatus")
        return [dict(zip(campos, r)) for r in cur.fetchall()]

# ── Fábrica (para futuros backends) ──────────────────────────────────────────
def crear_persistencia(motor: str = "firebird", **kw) -> Persistencia:
    motor = (motor or "firebird").lower()
    if motor == "firebird":
        return FirebirdPersistencia(**kw)
    raise ValueError(f"Motor de persistencia no soportado aún: {motor}")
