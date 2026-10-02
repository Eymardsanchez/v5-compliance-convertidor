"""
Gestion de usuarios del sistema — ConvertidorBancos
====================================================
Almacenamiento: %APPDATA%\\ConvertidorBancos\\usuarios.bin
Cifrado: XOR-SHA256 con MAC address.
Contrasenas: PBKDF2-HMAC-SHA256, 200 000 iteraciones + salt aleatorio de 32 bytes.
Recuperacion: codigo de 12 chars hex (XXXX-XXXX-XXXX) generado al crear usuario,
              almacenado como SHA-256 con salt. Mostrado UNA sola vez.

API publica:
  hay_usuarios()                            -> bool
  crear_usuario(user, pwd, creado_por)      -> (bool, str, str)  # (ok, msg, recovery_code)
  verificar_usuario(user, pwd)              -> bool
  cambiar_password(user, old, new)          -> (bool, str)
  recuperar_password(user, code, new_pwd)   -> (bool, str)
  regenerar_codigo(user, pwd)               -> (bool, str, str)  # (ok, msg, nuevo_code)
  listar_usuarios()                         -> list[str]
  eliminar_usuario(user)                    -> (bool, str)
"""

import hashlib
import json
import os
import secrets
import uuid
from pathlib import Path
from typing import List, Tuple


# ── Persistencia ──────────────────────────────────────────────────────────────

def _app_dir() -> Path:
    p = Path(os.environ.get("APPDATA", Path.home())) / "ConvertidorBancos"
    p.mkdir(parents=True, exist_ok=True)
    return p

def _path() -> Path:
    return _app_dir() / "usuarios.bin"

def _cipher_key() -> bytes:
    """Clave de cifrado derivada del MAC address (consistente entre versiones)."""
    return hashlib.sha256(str(uuid.getnode()).encode()).digest()

def _xor(data: bytes, key: bytes) -> bytes:
    return bytes(b ^ key[i % len(key)] for i, b in enumerate(data))

def _load() -> dict:
    try:
        if not _path().exists():
            return {}
        raw = _xor(_path().read_bytes(), _cipher_key())
        return json.loads(raw.decode("utf-8"))
    except Exception:
        return {}

def _save(data: dict) -> None:
    _path().write_bytes(_xor(json.dumps(data).encode("utf-8"), _cipher_key()))


# ── Hashing ───────────────────────────────────────────────────────────────────

def _hash_pwd(password: str, salt: str) -> str:
    """PBKDF2-HMAC-SHA256 con 200 000 iteraciones. Resistente a fuerza bruta."""
    return hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        salt.encode("utf-8"),
        200_000,
    ).hex()

def _hash_recovery(code: str, salt: str) -> str:
    """SHA-256 con salt para el codigo de recuperacion.
    El codigo tiene alta entropia propia (48 bits), no necesita PBKDF2."""
    clean = code.replace("-", "").upper()
    return hashlib.sha256(f"{salt}{clean}".encode("utf-8")).hexdigest()

def _generar_codigo() -> str:
    """Genera codigo de recuperacion: XXXX-XXXX-XXXX (12 chars hex mayusculas)."""
    raw = secrets.token_hex(6).upper()   # 12 hex chars
    return f"{raw[:4]}-{raw[4:8]}-{raw[8:]}"


# ── API publica ───────────────────────────────────────────────────────────────

def hay_usuarios() -> bool:
    """True si ya existe al menos un usuario registrado."""
    return bool(_load())


def crear_usuario(
    username: str,
    password: str,
    creado_por: str = "",
) -> Tuple[bool, str, str]:
    """
    Registra un nuevo usuario y genera su codigo de recuperacion.
    Retorna (exito, mensaje, recovery_code).
    recovery_code se muestra UNA sola vez — guarda su hash en el almacen.
    En caso de error retorna (False, mensaje, "").
    """
    username = username.strip().lower()
    if not username or not password:
        return False, "El usuario y la contrasena son obligatorios.", ""
    if len(username) < 3:
        return False, "El usuario debe tener al menos 3 caracteres.", ""
    if len(password) < 6:
        return False, "La contrasena debe tener al menos 6 caracteres.", ""

    data = _load()
    if username in data:
        return False, f"El usuario '{username}' ya existe.", ""

    pwd_salt = secrets.token_hex(32)
    rec_salt = secrets.token_hex(16)
    rec_code = _generar_codigo()

    data[username] = {
        "salt":         pwd_salt,
        "hash":         _hash_pwd(password, pwd_salt),
        "creado_por":   creado_por,
        "rec_salt":     rec_salt,
        "rec_hash":     _hash_recovery(rec_code, rec_salt),
    }
    _save(data)
    return True, f"Usuario '{username}' creado correctamente.", rec_code


def verificar_usuario(username: str, password: str) -> bool:
    """True si las credenciales son correctas."""
    data  = _load()
    entry = data.get(username.strip().lower())
    if not entry:
        return False
    return _hash_pwd(password, entry["salt"]) == entry["hash"]


def cambiar_password(
    username: str,
    pwd_actual: str,
    pwd_nueva: str,
) -> Tuple[bool, str]:
    """Cambia la contrasena validando la actual primero."""
    if not verificar_usuario(username, pwd_actual):
        return False, "Contrasena actual incorrecta."
    if len(pwd_nueva) < 6:
        return False, "La nueva contrasena debe tener al menos 6 caracteres."
    data  = _load()
    salt  = secrets.token_hex(32)
    uname = username.strip().lower()
    data[uname]["salt"] = salt
    data[uname]["hash"] = _hash_pwd(pwd_nueva, salt)
    _save(data)
    return True, "Contrasena actualizada correctamente."


def recuperar_password(
    username: str,
    recovery_code: str,
    new_password: str,
) -> Tuple[bool, str]:
    """
    Resetea la contrasena usando el codigo de recuperacion.
    El codigo se invalida tras el uso exitoso (se regenera uno nuevo
    que se retorna para que el usuario lo guarde).
    Retorna (exito, mensaje).
    """
    uname = username.strip().lower()
    data  = _load()
    entry = data.get(uname)

    if not entry:
        return False, "Usuario no encontrado."
    if not entry.get("rec_hash"):
        return False, "Este usuario no tiene codigo de recuperacion configurado."
    if len(new_password) < 6:
        return False, "La nueva contrasena debe tener al menos 6 caracteres."

    rec_salt = entry["rec_salt"]
    if _hash_recovery(recovery_code, rec_salt) != entry["rec_hash"]:
        return False, "Codigo de recuperacion incorrecto."

    # Resetear contrasena
    new_salt = secrets.token_hex(32)
    data[uname]["salt"] = new_salt
    data[uname]["hash"] = _hash_pwd(new_password, new_salt)

    # Invalidar codigo usado (generar nuevo hash vacio para que no sirva de nuevo)
    new_rec_salt = secrets.token_hex(16)
    new_rec_code = _generar_codigo()
    data[uname]["rec_salt"] = new_rec_salt
    data[uname]["rec_hash"] = _hash_recovery(new_rec_code, new_rec_salt)
    data[uname]["_nuevo_codigo"] = new_rec_code  # temporal, se lee y borra en la UI

    _save(data)
    return True, "Contrasena restablecida correctamente."


def tomar_nuevo_codigo(username: str) -> str:
    """
    Retorna y borra el codigo temporal generado tras recuperar_password().
    Llama justo despues de recuperar_password() exitoso para mostrar el nuevo codigo.
    """
    uname = username.strip().lower()
    data  = _load()
    entry = data.get(uname, {})
    code  = entry.pop("_nuevo_codigo", "")
    if code:
        _save(data)
    return code


def regenerar_codigo(
    username: str,
    password: str,
) -> Tuple[bool, str, str]:
    """
    Genera un nuevo codigo de recuperacion (requiere autenticacion con pwd actual).
    Retorna (exito, mensaje, nuevo_codigo).
    Util si el usuario perdio su codigo y aun recuerda la contrasena.
    """
    if not verificar_usuario(username, password):
        return False, "Contrasena incorrecta.", ""

    uname    = username.strip().lower()
    data     = _load()
    rec_salt = secrets.token_hex(16)
    rec_code = _generar_codigo()

    data[uname]["rec_salt"] = rec_salt
    data[uname]["rec_hash"] = _hash_recovery(rec_code, rec_salt)
    _save(data)
    return True, "Nuevo codigo generado correctamente.", rec_code


def listar_usuarios() -> List[str]:
    """Retorna la lista de nombres de usuario registrados."""
    return sorted(_load().keys())


def eliminar_usuario(username: str) -> Tuple[bool, str]:
    """Elimina un usuario. No permite eliminar el ultimo."""
    data  = _load()
    uname = username.strip().lower()
    if uname not in data:
        return False, "Usuario no encontrado."
    if len(data) <= 1:
        return False, "No se puede eliminar el unico usuario del sistema."
    del data[uname]
    _save(data)
    return True, f"Usuario '{uname}' eliminado."
