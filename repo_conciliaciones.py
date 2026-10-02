"""
repo_conciliaciones.py — Historial persistente de conciliaciones bancarias.

Guarda el resultado completo de una conciliación (CFDIs, movimientos, matches)
para poder reabrirlo sin reprocesar XMLs ni estados de cuenta, y para generar
Pólizas COI aunque la app se haya cerrado entre la conciliación y las pólizas.

Almacenamiento: %APPDATA%\\ConvertidorBancos\\conciliaciones.db (SQLite)
El payload (JSON con CFDIs/movimientos) se cifra por fila con XOR-SHA256
atado al MAC de la máquina — mismo esquema que empresas.py / catalogo_*.py.

API pública:
  guardar_conciliacion(rfc, resultados, recibidos, emitidos, destino) -> int
  listar_conciliaciones(rfc)                                          -> list[dict]
  cargar_conciliacion(id_registro)                                    -> dict | None
  eliminar_conciliacion(id_registro)                                  -> None
"""

from __future__ import annotations

import hashlib
import json
import os
import sqlite3
import uuid
from dataclasses import asdict
from datetime import datetime
from pathlib import Path
from typing import Optional

from cfdi_parser import CFDI
from conciliacion import MovBanco, MovConciliado, ResultadoBanco


# ── Persistencia: carpeta y cifrado (mismo esquema que empresas.py) ───────────

def _app_dir() -> Path:
    p = Path(os.environ.get("APPDATA", Path.home())) / "ConvertidorBancos"
    p.mkdir(parents=True, exist_ok=True)
    return p


def _db_path() -> Path:
    return _app_dir() / "conciliaciones.db"


def _cipher_key() -> bytes:
    return hashlib.sha256(str(uuid.getnode()).encode()).digest()


def _xor(data: bytes, key: bytes) -> bytes:
    return bytes(b ^ key[i % len(key)] for i, b in enumerate(data))


def _cifrar(payload: dict) -> bytes:
    raw = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    return _xor(raw, _cipher_key())


def _descifrar(blob: bytes) -> dict:
    raw = _xor(blob, _cipher_key())
    return json.loads(raw.decode("utf-8"))


def _conectar() -> sqlite3.Connection:
    con = sqlite3.connect(_db_path())
    con.execute("""
        CREATE TABLE IF NOT EXISTS conciliaciones (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            rfc             TEXT    NOT NULL,
            fecha_guardado  TEXT    NOT NULL,
            destino_excel   TEXT    NOT NULL DEFAULT '',
            n_recibidos     INTEGER NOT NULL DEFAULT 0,
            n_emitidos      INTEGER NOT NULL DEFAULT 0,
            n_conciliados   INTEGER NOT NULL DEFAULT 0,
            n_pendientes    INTEGER NOT NULL DEFAULT 0,
            payload         BLOB    NOT NULL
        )
    """)
    con.execute("CREATE INDEX IF NOT EXISTS idx_conc_rfc ON conciliaciones(rfc)")
    return con


# ── Serialización de dataclasses ───────────────────────────────────────────────

def _dict_a_cfdi(d: Optional[dict]) -> Optional[CFDI]:
    return CFDI(**d) if d else None


def _dict_a_movbanco(d: dict) -> MovBanco:
    return MovBanco(**d)


def _dict_a_movconciliado(d: dict) -> MovConciliado:
    return MovConciliado(
        banco      = _dict_a_movbanco(d["banco"]),
        cfdi       = _dict_a_cfdi(d["cfdi"]),
        tipo       = d.get("tipo", ""),
        revisar    = d.get("revisar", False),
        nivel      = d.get("nivel", 0),
        cfdi_suma  = [CFDI(**c) for c in d.get("cfdi_suma", [])],
        tipo_match = d.get("tipo_match", ""),
    )


def _dict_a_resultado(d: dict) -> ResultadoBanco:
    return ResultadoBanco(
        nombre      = d["nombre"],
        conciliados = [_dict_a_movconciliado(mc) for mc in d["conciliados"]],
        sin_cfdi    = [_dict_a_movbanco(mb) for mb in d["sin_cfdi"]],
    )


# ── API pública ───────────────────────────────────────────────────────────────

def guardar_conciliacion(
    rfc:        str,
    resultados: list[ResultadoBanco],
    recibidos:  list[CFDI],
    emitidos:   list[CFDI],
    destino:    str,
) -> int:
    """
    Guarda el resultado completo de una conciliación ya exportada con éxito.
    Retorna el id de la fila insertada.

    Llamar SOLO después de exportar el Excel con éxito — nunca con un
    resultado parcial. El caller debe envolver esta llamada en try/except:
    si falla el guardado (disco lleno, antivirus, etc.) el Excel ya generado
    sigue siendo válido y la conciliación no debe interrumpirse por esto.
    """
    n_conciliados = sum(len(r.conciliados) for r in resultados)
    n_pendientes  = sum(len(r.sin_cfdi)    for r in resultados)

    payload = {
        "resultados": [asdict(r) for r in resultados],
        "recibidos":  [asdict(c) for c in recibidos],
        "emitidos":   [asdict(c) for c in emitidos],
        "destino":    destino,
        "rfc":        rfc,
    }

    con = _conectar()
    try:
        cur = con.execute(
            """INSERT INTO conciliaciones
               (rfc, fecha_guardado, destino_excel, n_recibidos, n_emitidos,
                n_conciliados, n_pendientes, payload)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                rfc.strip().upper(),
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                destino,
                len(recibidos),
                len(emitidos),
                n_conciliados,
                n_pendientes,
                _cifrar(payload),
            ),
        )
        con.commit()
        return cur.lastrowid
    finally:
        con.close()


def listar_conciliaciones(rfc: str) -> list[dict]:
    """
    Retorna metadata liviana (sin el payload pesado) de las conciliaciones
    guardadas para un RFC, más reciente primero.

    Cada dict: id, fecha_guardado, destino_excel, n_recibidos, n_emitidos,
    n_conciliados, n_pendientes.

    Si la base de datos no existe o está corrupta, retorna [] en vez de
    lanzar — el historial vacío no debe impedir usar la app.
    """
    try:
        con = _conectar()
    except sqlite3.Error:
        return []
    try:
        rows = con.execute(
            """SELECT id, fecha_guardado, destino_excel, n_recibidos,
                      n_emitidos, n_conciliados, n_pendientes
               FROM conciliaciones
               WHERE rfc = ?
               ORDER BY id DESC""",
            (rfc.strip().upper(),),
        ).fetchall()
    except sqlite3.Error:
        return []
    finally:
        con.close()

    cols = ["id", "fecha_guardado", "destino_excel", "n_recibidos",
            "n_emitidos", "n_conciliados", "n_pendientes"]
    return [dict(zip(cols, row)) for row in rows]


def cargar_conciliacion(id_registro: int) -> Optional[dict]:
    """
    Reconstruye el resultado completo de una conciliación guardada, con la
    misma forma que VentanaConciliacion._ultimo_resultado:
    {"resultados": list[ResultadoBanco], "recibidos": list[CFDI],
     "emitidos": list[CFDI], "destino": str, "rfc": str}

    Retorna None si el id no existe o el registro está corrupto.
    """
    try:
        con = _conectar()
    except sqlite3.Error:
        return None
    try:
        row = con.execute(
            "SELECT payload FROM conciliaciones WHERE id = ?",
            (id_registro,),
        ).fetchone()
    except sqlite3.Error:
        return None
    finally:
        con.close()

    if not row:
        return None

    try:
        payload = _descifrar(row[0])
        return {
            "resultados": [_dict_a_resultado(r) for r in payload["resultados"]],
            "recibidos":  [CFDI(**c) for c in payload["recibidos"]],
            "emitidos":   [CFDI(**c) for c in payload["emitidos"]],
            "destino":    payload["destino"],
            "rfc":        payload["rfc"],
        }
    except Exception:
        return None


def eliminar_conciliacion(id_registro: int) -> None:
    """Elimina una entrada del historial. No lanza si el id no existe."""
    con = _conectar()
    try:
        con.execute("DELETE FROM conciliaciones WHERE id = ?", (id_registro,))
        con.commit()
    finally:
        con.close()


def obtener_ultima_conciliacion(rfc: str) -> Optional[dict]:
    """
    Retorna el resultado completo (misma forma que cargar_conciliacion) de
    la conciliación MÁS RECIENTE guardada para un RFC, o None si no hay
    ninguna o el RFC no tiene historial.

    Se usa como base para fusionar una conciliación nueva con la anterior:
    como cada fila guardada ya incluye lo fusionado con SU predecesora, solo
    hace falta la más reciente — no todo el historial.
    """
    registros = listar_conciliaciones(rfc)
    if not registros:
        return None
    return cargar_conciliacion(registros[0]["id"])
