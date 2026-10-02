# -*- coding: utf-8 -*-
"""
compliance_engine.py — Motor de Obligaciones Fiscales/Laborales
===============================================================
Catálogo de obligaciones, configuración por empresa y generación
del cuerpo del correo de recordatorio.

Config se guarda en %APPDATA%\\ConvertidorBancos\\compliance.json
Estructura:
  {
    "RFC123": {
        "destinatario": "proveedor@x.com",
        "cc":           "contador@x.com,director@x.com",
        "asunto_template": "Obligaciones {empresa} – {mes} {anio} | Entrega antes del día {dia}",
        "obligaciones_activas": ["isr_declaracion", "iva_declaracion", ...],
        "carpeta_raiz": "C:\\Clientes\\EmpresaXYZ",
        "notas_adicionales": ""
    },
    ...
  }
"""

import json
import os
import calendar
from datetime import date, datetime
from pathlib  import Path

# ── Catálogo de obligaciones ──────────────────────────────────────────────────
OBLIGACIONES_CATALOGO: list[dict] = [
    {"id": "listado_trabajadores",          "nombre": "Listado de Trabajadores (Excel)",                       "periodicidad": "Mensual",       "meses_aplica": list(range(1, 13))},
    {"id": "recibos_nomina",                "nombre": "Recibos de Nómina (.xml)",                              "periodicidad": "Mensual",       "meses_aplica": list(range(1, 13))},
    {"id": "isr_declaracion",               "nombre": "ISR – Declaración Mensual (PDF)",                       "periodicidad": "Mensual",       "meses_aplica": list(range(1, 13))},
    {"id": "isr_acuse",                     "nombre": "Acuse de pago ISR Retenciones (PDF)",                   "periodicidad": "Mensual",       "meses_aplica": list(range(1, 13))},
    {"id": "iva_declaracion",               "nombre": "IVA – Declaración Mensual Propio (PDF)",                "periodicidad": "Mensual",       "meses_aplica": list(range(1, 13))},
    {"id": "iva_acuse",                     "nombre": "Acuse de pago IVA Propio (PDF)",                        "periodicidad": "Mensual",       "meses_aplica": list(range(1, 13))},
    {"id": "isr_ret_declaracion",           "nombre": "ISR Retenciones – Declaración Mensual (PDF)",           "periodicidad": "Mensual",       "meses_aplica": list(range(1, 13))},
    {"id": "iva_ret_declaracion",           "nombre": "IVA Retenciones – Declaración Mensual (PDF)",           "periodicidad": "Mensual",       "meses_aplica": list(range(1, 13))},
    {"id": "ieps_declaracion",              "nombre": "IEPS – Declaración Mensual (PDF)",                      "periodicidad": "Mensual",       "meses_aplica": list(range(1, 13))},
    {"id": "ieps_ret_declaracion",          "nombre": "IEPS Retenciones – Declaración Mensual (PDF)",          "periodicidad": "Mensual",       "meses_aplica": list(range(1, 13))},
    {"id": "opinion_sat",                   "nombre": "Opinión del Cumplimiento SAT (PDF)",                    "periodicidad": "Mensual",       "meses_aplica": list(range(1, 13))},
    {"id": "ism_declaracion",               "nombre": "ISM – Declaración Impuesto Sobre Nóminas",              "periodicidad": "Mensual",       "meses_aplica": list(range(1, 13))},
    {"id": "ss_mensual",                    "nombre": "Seguridad Social – Cédula Mensual (SUA)",               "periodicidad": "Mensual",       "meses_aplica": list(range(1, 13))},
    {"id": "ss_bimestral",                  "nombre": "Seguridad Social – Cédula Bimestral (SUA)",             "periodicidad": "Bimestral",     "meses_aplica": [2, 4, 6, 8, 10, 12]},
    {"id": "factura_imss",                  "nombre": "Factura IMSS (.XML)",                                   "periodicidad": "Mensual",       "meses_aplica": list(range(1, 13))},
    {"id": "factura_infonavit",             "nombre": "Factura INFONAVIT (.XML)",                              "periodicidad": "Mensual",       "meses_aplica": list(range(1, 13))},
    {"id": "opinion_imss",                  "nombre": "Opinión del Cumplimiento IMSS",                         "periodicidad": "Mensual",       "meses_aplica": list(range(1, 13))},
    {"id": "constancia_infonavit",          "nombre": "Constancia Situación Fiscal INFONAVIT",                 "periodicidad": "Mensual",       "meses_aplica": list(range(1, 13))},
    {"id": "acuse_imss_cuatrimestral",      "nombre": "Acuse Cuatrimestral IMSS (CSOI)",                       "periodicidad": "Cuatrimestral", "meses_aplica": [4, 8, 12]},
    {"id": "acuse_infonavit_cuatrimestral", "nombre": "Acuse Cuatrimestral INFONAVIT (SESIB)",                 "periodicidad": "Cuatrimestral", "meses_aplica": [4, 8, 12]},
    {"id": "declaracion_anual_pm",          "nombre": "Declaración Anual Personas Morales (PDF+acuse)",        "periodicidad": "Anual",         "meses_aplica": [3]},
    {"id": "declaracion_anual_pf",          "nombre": "Declaración Anual Personas Físicas (PDF+acuse)",        "periodicidad": "Anual",         "meses_aplica": [4]},
    {"id": "diot",                          "nombre": "DIOT – Operaciones con Terceros",                       "periodicidad": "Mensual",       "meses_aplica": list(range(1, 13))},
    {"id": "cfdi_nomina_xml",               "nombre": "CFDIs de Nómina timbrados (.zip XML)",                  "periodicidad": "Mensual",       "meses_aplica": list(range(1, 13))},
    {"id": "ptu",                           "nombre": "Cálculo PTU (Participación de Utilidades)",             "periodicidad": "Anual",         "meses_aplica": [5, 6]},
    {"id": "declaracion_info_sueldos",      "nombre": "Declaración Informativa de Sueldos (DIMM)",             "periodicidad": "Anual",         "meses_aplica": [2]},
]

TODOS_IDS = [o["id"] for o in OBLIGACIONES_CATALOGO]

MESES_ES = [
    "", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
]

ASUNTO_DEFAULT = "Obligaciones {empresa} – {mes} {anio} | Entrega antes del día {dia}"

# ── Persistencia ──────────────────────────────────────────────────────────────
def _app_dir() -> Path:
    p = Path(os.environ.get("APPDATA", Path.home())) / "ConvertidorBancos"
    p.mkdir(parents=True, exist_ok=True)
    return p

def _path() -> Path:
    return _app_dir() / "compliance.json"

def cargar_compliance() -> dict:
    """Carga el dict {RFC → config_compliance}."""
    try:
        if _path().exists():
            return json.loads(_path().read_text(encoding="utf-8"))
    except Exception:
        pass
    return {}

def guardar_compliance(data: dict) -> None:
    _path().write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

def obtener_config(rfc: str) -> dict:
    """Devuelve la config compliance de un RFC (crea default si no existe)."""
    data = cargar_compliance()
    if rfc not in data:
        data[rfc] = {
            "destinatario":          "",
            "cc":                    "",
            "asunto_template":       ASUNTO_DEFAULT,
            "obligaciones_activas":  TODOS_IDS.copy(),
            "carpeta_raiz":          "",
            "notas_adicionales":     "",
        }
        guardar_compliance(data)
    return data[rfc]

def guardar_config(rfc: str, cfg: dict) -> None:
    data = cargar_compliance()
    data[rfc] = cfg
    guardar_compliance(data)

# ── Presentaciones de declaraciones (marca automática en Obligaciones) ─────────
# Cuando se sube una declaración (ISR/IVA/IEPS/retenciones/anual/informativa…) se
# marca la obligación como PRESENTADA para ese periodo, guardando el archivo, la
# fecha de presentación y el número de operación. Se persiste dentro de la config
# de la empresa: cfg["presentaciones"]["YYYY-MM"][oblig_id] = {…}.
def _clave_periodo(anio: int, mes: int) -> str:
    return f"{int(anio):04d}-{int(mes):02d}"

def marcar_presentacion(rfc: str, anio: int, mes: int, oblig_id: str,
                        archivo: str = "", fecha_presentacion: str = "",
                        numero_operacion: str = "", usuario: str = "") -> dict:
    """Marca una obligación como presentada en el periodo (anio, mes). Idempotente
    por (periodo, oblig_id): vuelve a subir → actualiza los datos. Devuelve el
    registro guardado."""
    from datetime import datetime as _dt
    cfg = obtener_config(rfc)
    pres = cfg.setdefault("presentaciones", {})
    periodo = pres.setdefault(_clave_periodo(anio, mes), {})
    reg = {
        "archivo":            archivo or "",
        "fecha_presentacion": fecha_presentacion or "",
        "numero_operacion":   numero_operacion or "",
        "usuario":            usuario or "",
        "subido_en":          _dt.now().isoformat(sep=" ", timespec="seconds"),
    }
    periodo[oblig_id] = reg
    guardar_config(rfc, cfg)
    _fdb_guardar_presentacion(rfc, anio, mes, oblig_id, reg)
    return reg


# ── Espejo a la FDB de la informacion PROCESADA (para reportes posteriores) ───
# Las presentaciones viven en compliance.json (fuente operativa) y ADEMAS se
# guardan en la tabla PRESENTACION de Firebird, para dejar almacenado lo ya
# procesado y poder generar reportes despues. Si Firebird no responde, la app
# sigue funcionando con el JSON.
_persist_fdb_cache = None


def _persist_fdb():
    global _persist_fdb_cache
    if _persist_fdb_cache is not None:
        return _persist_fdb_cache
    try:
        from persistencia import FirebirdPersistencia
        _p = FirebirdPersistencia()
        _p.inicializar()
        _persist_fdb_cache = _p
        return _p
    except Exception:
        return None


def _reset_fdb():
    global _persist_fdb_cache
    try:
        if _persist_fdb_cache is not None:
            _persist_fdb_cache.cerrar()
    except Exception:
        pass
    _persist_fdb_cache = None


def _fdb_guardar_presentacion(rfc, anio, mes, oblig_id, reg):
    p = _persist_fdb()
    if p is None:
        return
    try:
        p.guardar_presentacion(rfc, anio, mes, oblig_id,
                               archivo=reg.get("archivo", ""),
                               fecha_presentacion=reg.get("fecha_presentacion", ""),
                               numero_operacion=reg.get("numero_operacion", ""),
                               usuario=reg.get("usuario", ""))
    except Exception:
        _reset_fdb()


def _fdb_eliminar_presentacion(rfc, anio, mes, oblig_id):
    p = _persist_fdb()
    if p is None:
        return
    try:
        p.eliminar_presentacion(rfc, anio, mes, oblig_id)
    except Exception:
        _reset_fdb()

def obtener_presentaciones(rfc: str, anio: int, mes: int) -> dict:
    """Devuelve {oblig_id: {archivo, fecha_presentacion, numero_operacion, …}} de
    las declaraciones ya presentadas en ese periodo (dict vacío si no hay)."""
    cfg = obtener_config(rfc)
    return dict(cfg.get("presentaciones", {}).get(_clave_periodo(anio, mes), {}))

def quitar_presentacion(rfc: str, anio: int, mes: int, oblig_id: str) -> bool:
    """Quita la marca de presentada (no borra el archivo físico). True si existía."""
    cfg = obtener_config(rfc)
    periodo = cfg.get("presentaciones", {}).get(_clave_periodo(anio, mes), {})
    if oblig_id in periodo:
        del periodo[oblig_id]
        guardar_config(rfc, cfg)
        _fdb_eliminar_presentacion(rfc, anio, mes, oblig_id)
        return True
    return False


def sincronizar_presentaciones(rfc: str, anio: int, mes: int) -> list[str]:
    """Auto-marca/desmarca según el ARCHIVO en disco: si el PDF registrado de una
    presentación ya no existe (se borró o movió), la desmarca y la mueve a la
    bitácora (``cfg["presentaciones_historial"]``) conservando fecha/operación.
    Se llama al refrescar la pestaña. Devuelve la lista de oblig_id desmarcados.

    Solo desmarca cuando el registro tiene una ruta de archivo verificable; si la
    ruta está vacía (empresa sin carpeta configurada) NO se toca (no se puede
    comprobar). Escribe a disco únicamente si hubo cambios."""
    from datetime import datetime as _dt
    cfg = obtener_config(rfc)
    clave = _clave_periodo(anio, mes)
    pres = cfg.get("presentaciones", {}).get(clave, {})
    if not pres:
        return []
    hist = cfg.setdefault("presentaciones_historial", [])
    desmarcados: list[str] = []
    for oid in list(pres.keys()):
        reg = pres.get(oid) or {}
        arch = reg.get("archivo", "") or ""
        if arch and not os.path.exists(arch):
            h = dict(reg)
            h.update({"periodo": clave, "oblig_id": oid,
                      "desmarcado_en": _dt.now().isoformat(sep=" ", timespec="seconds"),
                      "motivo": "archivo no encontrado"})
            hist.append(h)
            del pres[oid]
            desmarcados.append(oid)
            _fdb_eliminar_presentacion(rfc, anio, mes, oid)
    if desmarcados:
        guardar_config(rfc, cfg)
    return desmarcados


def leer_historial_presentaciones(rfc: str) -> list[dict]:
    """Bitácora de presentaciones desmarcadas (más recientes primero)."""
    cfg = obtener_config(rfc)
    hist = list(cfg.get("presentaciones_historial", []))
    return list(reversed(hist))

# ── Lógica de obligaciones ────────────────────────────────────────────────────
def dia_limite_habil(anio: int, mes: int, dia: int = 27) -> date:
    """Retrocede al último día hábil si el día cae en fin de semana."""
    d = date(anio, mes, min(dia, calendar.monthrange(anio, mes)[1]))
    while d.weekday() >= 5:
        d = date(d.year, d.month, d.day - 1)
    return d

def obligaciones_del_mes(cfg_rfc: dict, mes: int) -> list[dict]:
    """Lista obligaciones activas para el mes dado."""
    activas = set(cfg_rfc.get("obligaciones_activas", TODOS_IDS))
    return [o for o in OBLIGACIONES_CATALOGO
            if o["id"] in activas and mes in o["meses_aplica"]]

# ── Explorador de archivos por mes ───────────────────────────────────────────
def carpeta_mes(cfg_rfc: dict, anio: int, mes: int) -> str | None:
    """Ruta a la subcarpeta YYYY-MM si existe, None si no."""
    raiz = cfg_rfc.get("carpeta_raiz", "")
    if not raiz or not os.path.isdir(raiz):
        return None
    ruta = os.path.join(raiz, f"{anio:04d}-{mes:02d}")
    return ruta if os.path.isdir(ruta) else None

def asegurar_carpeta_mes(cfg_rfc: dict, anio: int, mes: int) -> str | None:
    """Como carpeta_mes pero la CREA si no existe. None si la empresa no tiene
    carpeta_raiz configurada o no es accesible (para poder avisar en la UI)."""
    raiz = cfg_rfc.get("carpeta_raiz", "")
    if not raiz or not os.path.isdir(raiz):
        return None
    ruta = os.path.join(raiz, f"{anio:04d}-{mes:02d}")
    try:
        os.makedirs(ruta, exist_ok=True)
        return ruta
    except Exception:
        return None

def listar_archivos_mes(cfg_rfc: dict, anio: int, mes: int) -> list[dict]:
    """Lista archivos en la subcarpeta YYYY-MM."""
    ruta = carpeta_mes(cfg_rfc, anio, mes)
    if not ruta:
        return []
    result = []
    for nombre in sorted(os.listdir(ruta)):
        ruta_f = os.path.join(ruta, nombre)
        if os.path.isfile(ruta_f):
            result.append({
                "nombre": nombre,
                "ruta":   ruta_f,
                "tamano": os.path.getsize(ruta_f),
                "mtime":  datetime.fromtimestamp(os.path.getmtime(ruta_f)).strftime("%d/%m/%Y %H:%M"),
            })
    return result

def formato_tamano(b: int) -> str:
    if b < 1024:       return f"{b} B"
    if b < 1_048_576:  return f"{b/1024:.1f} KB"
    return f"{b/1_048_576:.1f} MB"

# ── Generación del cuerpo del correo ─────────────────────────────────────────
def generar_cuerpo(
    empresa_nombre: str,
    rfc:            str,
    cfg_rfc:        dict,
    mes:            int,
    anio:           int,
    archivos_adj:   list[str] | None = None,
) -> str:
    obs       = obligaciones_del_mes(cfg_rfc, mes)
    limite    = dia_limite_habil(anio, mes, 27)
    rfc_str   = f" (RFC: {rfc})" if rfc else ""
    lineas    = [f"  • {o['nombre']} ({o['periodicidad']})" for o in obs]
    lista     = "\n".join(lineas) or "  (Sin obligaciones configuradas para este mes)"
    notas     = cfg_rfc.get("notas_adicionales", "").strip()

    cuerpo = (
        f"Estimado proveedor,\n\n"
        f"Por medio del presente correo, {empresa_nombre}{rfc_str} le solicita "
        f"amablemente enviar la documentación correspondiente a sus obligaciones "
        f"fiscales y laborales del mes de {MESES_ES[mes]} {anio}.\n\n"
        f"FECHA LÍMITE DE ENTREGA: {limite.strftime('%d de %B de %Y')}\n\n"
        f"DOCUMENTOS REQUERIDOS:\n{lista}\n"
    )
    if archivos_adj:
        cuerpo += f"\nARCHIVOS ADJUNTOS EN ESTE CORREO ({len(archivos_adj)}):\n"
        for a in archivos_adj:
            cuerpo += f"  📎 {os.path.basename(a)}\n"
    if notas:
        cuerpo += f"\nNOTAS ADICIONALES:\n{notas}\n"
    cuerpo += (
        f"\nFavor de enviar los documentos antes de la fecha indicada "
        f"para evitar penalizaciones contractuales.\n\n"
        f"Atentamente,\n{empresa_nombre}\n"
        f"Departamento de Recursos Humanos y Cumplimiento Fiscal\n\n"
        f"--- Generado por ConvertidorBancos v5.0 — Compliance + Convertidor ---\n"
    )
    return cuerpo

def generar_asunto(cfg_rfc: dict, empresa_nombre: str, mes: int, anio: int) -> str:
    tpl    = cfg_rfc.get("asunto_template", ASUNTO_DEFAULT)
    limite = dia_limite_habil(anio, mes, 27)
    try:
        return tpl.format(
            empresa=empresa_nombre,
            mes=MESES_ES[mes],
            anio=anio,
            dia=limite.day,
        )
    except Exception:
        return f"Obligaciones {empresa_nombre} – {MESES_ES[mes]} {anio}"


# ── Años de trabajo (lista personalizada, GLOBAL para toda la app) ────────────
# El usuario captura los años que necesita (p.ej. "2015, 2018, 2024-2026") en la
# sub-ventana de Config Empresa; esa lista alimenta los combos de Año de TODAS las
# pestañas (Panel General, Obligaciones, Contribuciones/Saldos, Archivos).
import re as _re

ANIO_MIN, ANIO_MAX = 2015, 2040


def _path_anios() -> Path:
    return _app_dir() / "anios.json"


def parsear_anios(texto: str) -> list:
    """'2015, 2018, 2024-2026' → [2015, 2018, 2024, 2025, 2026].
    Acepta comas/;, rangos con - – o 'a'. Ignora basura y acota a ANIO_MIN..ANIO_MAX."""
    out = set()
    for parte in _re.split(r"[,;]", texto or ""):
        parte = parte.strip()
        if not parte:
            continue
        m = _re.match(r"^(\d{4})\s*(?:-|–|a|A)\s*(\d{4})$", parte)
        if m:
            a, b = int(m.group(1)), int(m.group(2))
            if a > b:
                a, b = b, a
            for y in range(a, b + 1):
                out.add(y)
        elif _re.match(r"^\d{4}$", parte):
            out.add(int(parte))
    return sorted(y for y in out if ANIO_MIN <= y <= ANIO_MAX)


def formatear_anios(lista) -> str:
    """[2015,2024,2025,2026] → '2015, 2024-2026' (compacta rangos consecutivos)."""
    ys = sorted({int(y) for y in (lista or [])})
    if not ys:
        return ""
    partes, ini, prev = [], ys[0], ys[0]
    for y in ys[1:]:
        if y == prev + 1:
            prev = y
            continue
        partes.append(str(ini) if ini == prev else f"{ini}-{prev}")
        ini = prev = y
    partes.append(str(ini) if ini == prev else f"{ini}-{prev}")
    return ", ".join(partes)


def cargar_anios() -> list:
    """Años configurados por el usuario. [] si no hay configuración."""
    try:
        p = _path_anios()
        if p.exists():
            data = json.loads(p.read_text(encoding="utf-8"))
            ys = [int(y) for y in data.get("anios", [])]
            ys = [y for y in ys if ANIO_MIN <= y <= ANIO_MAX]
            if ys:
                return sorted(set(ys))
    except Exception:
        pass
    return []


def guardar_anios(lista) -> list:
    """Guarda la lista de años (acotada y sin duplicados). Devuelve la lista final."""
    ys = sorted({int(y) for y in (lista or []) if ANIO_MIN <= int(y) <= ANIO_MAX})
    try:
        _path_anios().write_text(json.dumps({"anios": ys}, ensure_ascii=False, indent=2),
                                 encoding="utf-8")
    except Exception:
        pass
    return ys


def anios_disponibles() -> list:
    """Años que se muestran en TODOS los combos de la app. Si no hay configuración,
    usa el rango completo permitido (2015–2040)."""
    ys = cargar_anios()
    return ys if ys else list(range(ANIO_MIN, ANIO_MAX + 1))
