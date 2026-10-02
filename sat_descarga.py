"""
sat_descarga.py — Descarga Masiva SAT para ConciliaRFC V5
Deal Global Consulting — Eymard Raymundo Sanchez Lima

Integra el servicio oficial SAT (Descarga Masiva) usando la e.firma del
contribuyente para descargar XMLs de CFDIs emitidos y recibidos de forma
automática. Este módulo sustituye la carga manual de XMLs en V5.

Dependencia: satcfdi >= 26.7.2 (pip install satcfdi)

Flujo SAT:
    1. Autenticar con e.firma → token temporal
    2. Solicitar descarga → IdSolicitud
    3. Polling hasta TERMINADA → IdsPaquetes
    4. Descargar cada paquete (ZIP base64) → extraer XMLs

Límites SAT (críticos para producción):
    - Max 10 solicitudes por día por RFC
    - Max 200,000 CFDIs por solicitud
    - Paginación automática por quincena si se supera el tope
"""

import base64
import io
import json
import logging
import os
import time
import zipfile
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Callable, Optional

from satcfdi.models import Signer
from satcfdi.pacs.sat import (
    SAT,
    EstadoComprobante,
    EstadoSolicitud,
    TipoDescargaMasivaTerceros,
)

logger = logging.getLogger(__name__)

# ─── Configuración ────────────────────────────────────────────────────────────
TIMEOUT_SEGUNDOS = 720       # 12 min máx esperando al SAT
POLL_INTERVAL_SEG = 60       # Revisar estado cada 60 segundos
MAX_SOLICITUDES_DIA = 10     # Límite SAT por RFC/día


# ─── Tipos de resultado ───────────────────────────────────────────────────────
@dataclass
class DescargaSATResult:
    emitidos: list[Path] = field(default_factory=list)
    recibidos: list[Path] = field(default_factory=list)
    # XMLs obtenidos por búsqueda directa de UUID (recover_comprobante_uuid_request).
    # No se clasifican aquí como emitido/recibido — esa clasificación real la
    # hace clasificar_cfdis() al cargarlos al conciliador, según el RFC.
    por_uuid: list[Path] = field(default_factory=list)
    errores: list[str] = field(default_factory=list)

    @property
    def total_cfdi(self) -> int:
        return len(self.emitidos) + len(self.recibidos) + len(self.por_uuid)

    @property
    def ok(self) -> bool:
        return not self.errores

    def resumen(self) -> str:
        partes = []
        if self.emitidos or self.recibidos:
            partes.append(f"{len(self.emitidos)} emitidos, {len(self.recibidos)} recibidos")
        if self.por_uuid:
            partes.append(f"{len(self.por_uuid)} por UUID")
        return (
            f"✅ {self.total_cfdi} CFDIs descargados"
            + (f" ({', '.join(partes)})" if partes else "")
            + (f" | ⚠ {len(self.errores)} error(es)" if self.errores else "")
        )


class DescargaSATError(Exception):
    """Error controlado en la comunicación con el SAT."""


# ─── Mejora #2: manejo de errores de RED (no confundir con rechazo del SAT) ────
try:
    import requests as _rq
    _ERR_RED: tuple = (ConnectionError, TimeoutError, OSError, _rq.exceptions.RequestException)
except Exception:                                            # noqa
    _ERR_RED = (ConnectionError, TimeoutError, OSError)


def _red_guard(desc: str, fn, *args, **kw):
    """Envuelve una llamada SOAP: un fallo de red se reporta CLARO como problema
    de conexión, NUNCA como rechazo/cuota del SAT (5002). No reintenta un 5002."""
    try:
        return fn(*args, **kw)
    except _ERR_RED as e:                                    # noqa
        raise DescargaSATError(
            f"No se pudo conectar al SAT ({desc}). Revisa tu conexión a internet — "
            f"esto NO es la cuota del SAT, es un problema de red ({type(e).__name__})."
        )


# ─── Mejora #1: contador de solicitudes por RFC/día (protege TU cuota) ─────────
def _ruta_cuota() -> Path:
    base = Path(os.environ.get("APPDATA") or Path.home()) / "ConvertidorBancos"
    try:
        base.mkdir(parents=True, exist_ok=True)
    except OSError:
        pass
    return base / "sat_cuota.json"


def _cargar_cuota() -> dict:
    try:
        return json.loads(_ruta_cuota().read_text(encoding="utf-8"))
    except Exception:                                        # noqa
        return {}


def solicitudes_hoy(rfc: str) -> int:
    """Cuántas solicitudes SAT se han hecho HOY para este RFC (aceptadas)."""
    if not rfc:
        return 0
    return int(_cargar_cuota().get(rfc, {}).get(date.today().isoformat(), 0))


def _registrar_solicitud(rfc: str) -> None:
    """Suma 1 al contador del día (solo tras una solicitud ACEPTADA por el SAT)."""
    if not rfc:
        return
    hoy = date.today().isoformat()
    data = _cargar_cuota()
    n = int(data.get(rfc, {}).get(hoy, 0)) + 1
    data[rfc] = {hoy: n}   # conserva solo el día de hoy (no crece indefinido)
    try:
        _ruta_cuota().write_text(json.dumps(data), encoding="utf-8")
    except Exception:                                        # noqa
        pass


# ─── Carga de FIEL ────────────────────────────────────────────────────────────
def cargar_fiel(
    cer_path: str | Path,
    key_path: str | Path,
    password: str,
) -> Signer:
    """
    Carga la e.firma (FIEL) desde archivos .cer y .key del SAT.

    Args:
        cer_path:  Ruta al archivo .cer (certificado público, DER)
        key_path:  Ruta al archivo .key (clave privada cifrada, DER)
        password:  Contraseña de la clave privada

    Returns:
        Signer listo para usar con SAT Descarga Masiva

    Raises:
        DescargaSATError: Archivo no encontrado o contraseña incorrecta
    """
    cer_path = Path(cer_path)
    key_path = Path(key_path)

    if not cer_path.exists():
        raise DescargaSATError(f"Archivo .cer no encontrado: {cer_path}")
    if not key_path.exists():
        raise DescargaSATError(f"Archivo .key no encontrado: {key_path}")

    try:
        signer = Signer.load(
            certificate=cer_path.read_bytes(),
            key=key_path.read_bytes(),
            password=password.encode("utf-8"),
        )
    except Exception as exc:
        raise DescargaSATError(
            f"No se pudo cargar la FIEL. "
            f"Verifica que la contraseña sea correcta y los archivos correspondan al mismo RFC. "
            f"Detalle: {exc}"
        ) from exc

    expira, _ = validar_vigencia_fiel(signer)
    logger.info(
        "FIEL cargada — RFC: %s | Vigente hasta: %s",
        signer.rfc,
        _fiel_expiry(signer).strftime("%d/%m/%Y"),
    )
    return signer


def _fiel_expiry(signer: Signer):
    """Devuelve datetime UTC de expiración del certificado."""
    from datetime import datetime, timezone
    raw = signer.certificate.get_notAfter().decode("ascii")
    return datetime.strptime(raw, "%Y%m%d%H%M%SZ").replace(tzinfo=timezone.utc)


def validar_vigencia_fiel(signer: Signer) -> tuple[bool, int]:
    """
    Verifica si la e.firma está vigente.

    Returns:
        (vigente: bool, dias_restantes: int)
    """
    from datetime import datetime, timezone
    hoy = datetime.now(timezone.utc)
    expira = _fiel_expiry(signer)
    dias = (expira - hoy).days
    return dias > 0, dias


# ─── Descarga por mes (punto de entrada principal para V5) ───────────────────
def descargar_mes(
    signer: Signer,
    anio: int,
    mes: int,
    carpeta_destino: str | Path,
    callback: Optional[Callable[[str], None]] = None,
) -> DescargaSATResult:
    """
    Descarga todos los CFDIs de un mes completo (emitidos + recibidos).
    Punto de entrada principal para la integración con V5.

    Args:
        signer:           Signer cargado con cargar_fiel()
        anio:             Año del período (ej: 2026)
        mes:              Mes del período (1-12)
        carpeta_destino:  Carpeta raíz donde crear emitidos/ y recibidos/
        callback:         Función opcional para reportar progreso al UI

    Returns:
        DescargaSATResult con listas de Paths de XMLs descargados
    """
    fecha_ini = date(anio, mes, 1)
    fecha_fin = _ultimo_dia_mes(anio, mes)

    return descargar_rango(
        signer=signer,
        fecha_inicial=fecha_ini,
        fecha_final=fecha_fin,
        carpeta_destino=carpeta_destino,
        callback=callback,
    )


_TIPOS_VALIDOS = {"emitidos", "recibidos"}


# ─── Descarga por rango de fechas ─────────────────────────────────────────────
def descargar_rango(
    signer: Signer,
    fecha_inicial: date,
    fecha_final: date,
    carpeta_destino: str | Path,
    callback: Optional[Callable[[str], None]] = None,
    tipos: Optional[set[str]] = None,
) -> DescargaSATResult:
    """
    Descarga CFDIs para un rango de fechas arbitrario. Por default trae
    emitidos y recibidos (cada uno es una solicitud SAT independiente);
    pasar `tipos={"emitidos"}` o `tipos={"recibidos"}` para pedir solo uno
    y ahorrar tiempo, ancho de banda y una unidad del límite diario SAT.
    Maneja paginación automática si el SAT reporta tope máximo (cod 5003).

    Args:
        signer:          Signer cargado con cargar_fiel()
        fecha_inicial:   Fecha inicio del período
        fecha_final:     Fecha fin del período
        carpeta_destino: Carpeta raíz destino
        callback:        Función de progreso (msg: str) para la UI
        tipos:           {"emitidos"}, {"recibidos"} o ambos (default: ambos)

    Returns:
        DescargaSATResult con lo solicitado (emitidos y/o recibidos)
    """
    tipos = tipos if tipos else set(_TIPOS_VALIDOS)
    invalidos = tipos - _TIPOS_VALIDOS
    if invalidos:
        raise ValueError(f"tipos inválido(s): {invalidos}. Usa 'emitidos'/'recibidos'.")

    carpeta = Path(carpeta_destino)
    carpeta.mkdir(parents=True, exist_ok=True)

    vigente, dias = validar_vigencia_fiel(signer)
    if not vigente:
        raise DescargaSATError("La e.firma está vencida. Renuévala en el SAT.")
    if dias < 30:
        _log(callback, f"⚠ AVISO: tu e.firma vence en {dias} días. Renuévala pronto.")

    sat = SAT(signer=signer)
    resultado = DescargaSATResult()

    _log(callback,
         f"[SAT] RFC: {signer.rfc} | Período: {fecha_inicial} → {fecha_final} | "
         f"Tipos: {', '.join(sorted(tipos))}")

    # ── Emitidos ──
    if "emitidos" in tipos:
        if solicitudes_hoy(signer.rfc) >= MAX_SOLICITUDES_DIA:
            _msg = (f"Ya usaste {MAX_SOLICITUDES_DIA} solicitudes hoy para {signer.rfc}. "
                    f"Espera a mañana — es protección de TU cuota diaria del SAT.")
            resultado.errores.append(f"Emitidos: {_msg}")
            _log(callback, f"[SAT] ⛔ {_msg}")
        else:
            _log(callback, "[SAT] Solicitando CFDIs emitidos...")
            try:
                xmls = _ciclo_descarga(
                    sat=sat,
                    tipo="emitidos",
                    fecha_inicial=fecha_inicial,
                    fecha_final=fecha_final,
                    carpeta=carpeta / "emitidos",
                    callback=callback,
                )
                _registrar_solicitud(signer.rfc)   # el SAT la aceptó (no lanzó)
                resultado.emitidos = xmls
            except DescargaSATError as exc:
                resultado.errores.append(f"Emitidos: {exc}")
                _log(callback, f"[SAT] ⚠ Error emitidos: {exc}")

    # ── Recibidos ──
    if "recibidos" in tipos:
        if solicitudes_hoy(signer.rfc) >= MAX_SOLICITUDES_DIA:
            _msg = (f"Ya usaste {MAX_SOLICITUDES_DIA} solicitudes hoy para {signer.rfc}. "
                    f"Espera a mañana — es protección de TU cuota diaria del SAT.")
            resultado.errores.append(f"Recibidos: {_msg}")
            _log(callback, f"[SAT] ⛔ {_msg}")
        else:
            _log(callback, "[SAT] Solicitando CFDIs recibidos...")
            try:
                xmls = _ciclo_descarga(
                    sat=sat,
                    tipo="recibidos",
                    fecha_inicial=fecha_inicial,
                    fecha_final=fecha_final,
                    carpeta=carpeta / "recibidos",
                    callback=callback,
                )
                _registrar_solicitud(signer.rfc)   # el SAT la aceptó (no lanzó)
                resultado.recibidos = xmls
            except DescargaSATError as exc:
                resultado.errores.append(f"Recibidos: {exc}")
                _log(callback, f"[SAT] ⚠ Error recibidos: {exc}")

    _log(callback, f"[SAT] {resultado.resumen()}")
    return resultado


# ─── Descarga por UUID (Folio Fiscal) específico ──────────────────────────────
def descargar_por_uuid(
    signer: Signer,
    uuids: list[str],
    carpeta_destino: str | Path,
    callback: Optional[Callable[[str], None]] = None,
) -> DescargaSATResult:
    """
    Descarga uno o más CFDIs puntuales por su Folio Fiscal (UUID), usando el
    servicio SAT recover_comprobante_uuid_request.

    ⚠ IMPORTANTE — costo en cuota SAT: cada UUID de esta lista genera una
    solicitud independiente y consume UNA unidad del límite de
    MAX_SOLICITUDES_DIA por RFC/día. Para lotes grandes o un mes completo,
    usar descargar_rango()/descargar_mes() en vez de esta función — cada
    tipo pedido (emitidos y/o recibidos) es una solicitud que cubre TODO
    el rango de fechas, en vez de una solicitud por cada CFDI individual.

    Args:
        signer:          Signer cargado con cargar_fiel()
        uuids:           Lista de Folios Fiscales (UUID) a buscar
        carpeta_destino: Carpeta raíz; los XMLs se guardan en
                         carpeta_destino/por_uuid/
        callback:        Función de progreso (msg: str) para la UI

    Returns:
        DescargaSATResult con los XMLs encontrados en `.por_uuid` y
        cualquier UUID no localizado/erróneo reflejado en `.errores`.
    """
    carpeta = Path(carpeta_destino) / "por_uuid"
    carpeta.mkdir(parents=True, exist_ok=True)

    vigente, dias = validar_vigencia_fiel(signer)
    if not vigente:
        raise DescargaSATError("La e.firma está vencida. Renuévala en el SAT.")
    if dias < 30:
        _log(callback, f"⚠ AVISO: tu e.firma vence en {dias} días. Renuévala pronto.")

    sat = SAT(signer=signer)
    resultado = DescargaSATResult()

    folios = [u.strip() for u in uuids if u.strip()]
    _log(callback, f"[SAT] RFC: {signer.rfc} | Búsqueda por UUID: {len(folios)} folio(s)")
    _log(callback, f"[SAT] ⚠ Cada UUID consume 1 de las {MAX_SOLICITUDES_DIA} solicitudes/día.")

    for i, folio in enumerate(folios, 1):
        _log(callback, f"[SAT] ({i}/{len(folios)}) Solicitando folio {folio}...")
        try:
            resp = sat.recover_comprobante_uuid_request(folio=folio)
            cod          = resp.get("CodEstatus", "")
            id_solicitud = resp.get("IdSolicitud", "")

            if not id_solicitud:
                raise DescargaSATError(
                    f"SAT no devolvió ID de solicitud para el folio {folio}. "
                    f"CodEstatus={cod}. Verifica que el UUID exista y te pertenezca."
                )

            ids_paquetes = _esperar_solicitud(sat, id_solicitud, f"UUID {folio[:8]}…", callback)
            if not ids_paquetes:
                _log(callback, f"[SAT] Sin resultado para el folio {folio}.")
                continue

            xmls = _descargar_paquetes(sat, ids_paquetes, carpeta, f"uuid {folio[:8]}…", callback)
            resultado.por_uuid.extend(xmls)

        except DescargaSATError as exc:
            resultado.errores.append(f"UUID {folio}: {exc}")
            _log(callback, f"[SAT] ⚠ Error en folio {folio}: {exc}")

    _log(callback, f"[SAT] {resultado.resumen()}")
    return resultado


def _rango_dia_completo(fecha_inicial: date, fecha_final: date) -> tuple[datetime, datetime]:
    """
    Convierte fechas (sin hora) a datetime con límites de día completo.

    satcfdi serializa un `date` puro sin componente de hora (ej. "2026-01-31"),
    y el SAT lo interpreta como medianoche — el portal de Consulta CFDI del
    SAT expone esto explícitamente con "Hora Inicial 00:00:00 / Hora Final
    23:59:59". Sin este ajuste, `fecha_final` trunca casi todo el último día
    del rango, provocando que CFDIs de esa fecha (típicamente recibidos,
    emitidos más tarde en el día) no aparezcan en la solicitud.
    """
    return (
        datetime.combine(fecha_inicial, datetime.min.time()),
        datetime.combine(fecha_final, datetime.max.time()),
    )


# ─── Ciclo completo: solicitar → polling → descargar ─────────────────────────
def _ciclo_descarga(
    sat: SAT,
    tipo: str,
    fecha_inicial: date,
    fecha_final: date,
    carpeta: Path,
    callback: Callable,
) -> list[Path]:
    """
    Ejecuta el ciclo completo SAT para un tipo (emitidos o recibidos):
    Solicitar → Polling → Descargar paquetes → Extraer XMLs
    """
    carpeta.mkdir(parents=True, exist_ok=True)

    fecha_ini_dt, fecha_fin_dt = _rango_dia_completo(fecha_inicial, fecha_final)

    # 1. SOLICITAR
    # EstadoComprobante=VIGENTE explícito: el SAT (Web Service v1.5) rechaza
    # con CodEstatus=301 "No se permite la descarga de xml que se encuentren
    # cancelados" cuando este parámetro se omite (default implícito = Todos,
    # que incluye cancelados). Para conciliación bancaria solo interesan los
    # vigentes de cualquier forma — un CFDI cancelado no debe conciliar.
    if tipo == "emitidos":
        resp = _red_guard(
            "solicitud emitidos", sat.recover_comprobante_emitted_request,
            fecha_inicial=fecha_ini_dt,
            fecha_final=fecha_fin_dt,
            tipo_solicitud=TipoDescargaMasivaTerceros.CFDI,
            estado_comprobante=EstadoComprobante.VIGENTE,
        )
    else:
        resp = _red_guard(
            "solicitud recibidos", sat.recover_comprobante_received_request,
            fecha_inicial=fecha_ini_dt,
            fecha_final=fecha_fin_dt,
            tipo_solicitud=TipoDescargaMasivaTerceros.CFDI,
            estado_comprobante=EstadoComprobante.VIGENTE,
        )

    cod = resp.get("CodEstatus", "")
    id_solicitud = resp.get("IdSolicitud", "")

    # Log completo de lo que el SAT respondió — no solo CodEstatus, para poder
    # diagnosticar con el detalle real (Mensaje/MensajeDetalle) si el SAT lo incluye.
    _log(callback, f"[SAT] Respuesta cruda {tipo}: {resp}")

    # 5003 = tope máximo → paginar por quincena
    if cod == "5003":
        _log(callback, f"[SAT] Tope máximo alcanzado en {tipo}. Paginando por quincenas...")
        return _descargar_paginado(sat, tipo, fecha_inicial, fecha_final, carpeta, callback)

    # 5005 = solicitud duplicada (mismos parámetros ya en cola)
    if cod == "5005":
        _log(callback, f"[SAT] Solicitud duplicada en {tipo} — reutilizando solicitud existente...")
        if not id_solicitud:
            raise DescargaSATError(
                "SAT reportó solicitud duplicada pero no devolvió el ID. "
                "Espera unos minutos y reintenta."
            )

    # Cualquier CodEstatus fuera de los éxitos conocidos (5000 = aceptada,
    # 5003 = tope máximo ya manejado arriba, 5005 = duplicada ya manejada
    # arriba) es un RECHAZO del SAT — aunque la respuesta incluya
    # IdSolicitud/RfcSolicitante, el SAT los repite incluso en errores
    # (confirmado en vivo con CodEstatus=5002). Sin este check, el código
    # solo validaba "¿vino IdSolicitud?" y reportaba "Solicitud aceptada"
    # para un rechazo real.
    if cod not in ("5000", "5005"):
        _MENSAJES_COD = {
            "5002": "El SAT agotó las solicitudes de por vida para este RFC. "
                    "No es un problema de conexión ni de la app — el SAT "
                    "rechazó la solicitud por límite de cuota; no se puede "
                    "reintentar hasta que el SAT libere cupo para este RFC.",
            "5004": "El SAT no encontró la solicitud referenciada (CodEstatus 5004).",
        }
        detalle = _MENSAJES_COD.get(cod, resp.get("Mensaje", "Sin mensaje del SAT."))
        raise DescargaSATError(
            f"SAT rechazó la solicitud de {tipo} (CodEstatus={cod}): {detalle}"
        )

    if not id_solicitud:
        raise DescargaSATError(
            f"SAT no devolvió ID de solicitud para {tipo}. "
            f"Respuesta completa: {resp}"
        )

    _log(callback, f"[SAT] {tipo.capitalize()} — Solicitud aceptada. ID: {id_solicitud}")

    # 2. POLLING
    ids_paquetes = _esperar_solicitud(sat, id_solicitud, tipo, callback)

    if not ids_paquetes:
        _log(callback, f"[SAT] Sin CFDIs {tipo} en el período seleccionado.")
        return []

    # 3. DESCARGAR PAQUETES → EXTRAER XMLs
    return _descargar_paquetes(sat, ids_paquetes, carpeta, tipo, callback)


def _esperar_solicitud(
    sat: SAT,
    id_solicitud: str,
    tipo: str,
    callback: Callable,
) -> list[str]:
    """
    Polling hasta que la solicitud SAT esté TERMINADA o falle.
    Returns lista de IdsPaquetes.
    """
    inicio = time.time()

    while time.time() - inicio < TIMEOUT_SEGUNDOS:
        time.sleep(POLL_INTERVAL_SEG)

        status = _red_guard("verificación de estado", sat.recover_comprobante_status, id_solicitud)
        estado_num = status.get("EstadoSolicitud", 0)
        num_cfdis = status.get("NumeroCFDIs", 0)

        try:
            estado_nombre = EstadoSolicitud(estado_num).name
        except ValueError:
            estado_nombre = str(estado_num)

        elapsed = int(time.time() - inicio)
        _log(
            callback,
            f"[SAT] {tipo.capitalize()} — {estado_nombre} | "
            f"{num_cfdis} CFDIs | {elapsed}s transcurridos...",
        )

        if estado_num == EstadoSolicitud.TERMINADA:
            ids = status.get("IdsPaquetes", [])
            _log(
                callback,
                f"[SAT] {tipo.capitalize()} listo: "
                f"{num_cfdis} CFDIs en {len(ids)} paquete(s)",
            )
            return ids

        if estado_num in (
            EstadoSolicitud.ERROR,
            EstadoSolicitud.RECHAZADA,
            EstadoSolicitud.VENCIDA,
        ):
            cod = status.get("CodigoEstadoSolicitud", "")
            raise DescargaSATError(
                f"SAT rechazó la solicitud {tipo}: "
                f"Estado={estado_nombre}, CodEstado={cod}"
            )

    raise DescargaSATError(
        f"Timeout esperando respuesta SAT para {tipo} "
        f"(más de {TIMEOUT_SEGUNDOS // 60} minutos). "
        "El SAT puede estar saturado. Reintenta más tarde."
    )


def _descargar_paquetes(
    sat: SAT,
    ids_paquetes: list[str],
    carpeta: Path,
    tipo: str,
    callback: Callable,
) -> list[Path]:
    """
    Descarga y extrae todos los paquetes ZIP del SAT.
    Deduplicación por nombre de archivo (UUID del CFDI).
    """
    xmls_descargados = []

    for i, id_paquete in enumerate(ids_paquetes, 1):
        _log(callback, f"[SAT] Descargando paquete {i}/{len(ids_paquetes)}: {id_paquete}")

        try:
            _header, paquete_b64 = _red_guard(
                "descarga de paquete", sat.recover_comprobante_download, id_paquete)
        except Exception as exc:
            _log(callback, f"[SAT] ⚠ Error descargando paquete {id_paquete}: {exc}")
            continue

        if not paquete_b64:
            _log(callback, f"[SAT] ⚠ Paquete vacío: {id_paquete}")
            continue

        zip_bytes = base64.b64decode(paquete_b64)
        nuevos = 0

        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
            xmls_en_zip = [n for n in zf.namelist() if n.lower().endswith(".xml")]

            carpeta_res = str(carpeta.resolve())
            for nombre_xml in xmls_en_zip:
                destino = (carpeta / nombre_xml).resolve()
                if not str(destino).startswith(carpeta_res):
                    _log(callback, f"[SAT] ⚠ Entrada XML fuera de carpeta, ignorada: {nombre_xml}")
                    continue
                destino.parent.mkdir(parents=True, exist_ok=True)
                if not destino.exists():
                    destino.write_bytes(zf.read(nombre_xml))
                    nuevos += 1
                xmls_descargados.append(destino)

        _log(
            callback,
            f"[SAT] Paquete {i}: {len(xmls_en_zip)} XMLs "
            f"({nuevos} nuevos, {len(xmls_en_zip) - nuevos} ya existían)",
        )

    return xmls_descargados


def _descargar_paginado(
    sat: SAT,
    tipo: str,
    fecha_ini: date,
    fecha_fin: date,
    carpeta: Path,
    callback: Callable,
) -> list[Path]:
    """
    Paginación por quincenas cuando el período supera 200K CFDIs.
    Divide el rango en bloques de ~15 días y los descarga secuencialmente.
    """
    todos_xmls = []
    cursor = fecha_ini

    while cursor <= fecha_fin:
        bloque_fin = min(cursor + timedelta(days=14), fecha_fin)
        _log(callback, f"[SAT] Paginando {tipo}: {cursor} → {bloque_fin}")

        xmls = _ciclo_descarga(sat, tipo, cursor, bloque_fin, carpeta, callback)
        todos_xmls.extend(xmls)

        cursor = bloque_fin + timedelta(days=1)
        if cursor <= fecha_fin:
            # Pausa entre solicitudes para no agotar el límite diario
            _log(callback, "[SAT] Pausa 5s entre bloques...")
            time.sleep(5)

    return todos_xmls


# ─── Utilidades ───────────────────────────────────────────────────────────────
def _ultimo_dia_mes(anio: int, mes: int) -> date:
    if mes == 12:
        return date(anio + 1, 1, 1) - timedelta(days=1)
    return date(anio, mes + 1, 1) - timedelta(days=1)


def _log(callback: Optional[Callable], msg: str) -> None:
    logger.info(msg)
    if callback:
        callback(msg)


# ─── Integración directa con V5 ───────────────────────────────────────────────
def descarga_v5(
    cer_path: str,
    key_path: str,
    password: str,
    anio: int,
    mes: int,
    carpeta_xml_base: str,
    callback: Optional[Callable[[str], None]] = None,
) -> dict:
    """
    Función de entrada única para V5.
    Reemplaza la carga manual de XMLs. Llamar desde el controlador de V5.

    Args:
        cer_path:         Ruta al .cer (puede ser relativa o absoluta)
        key_path:         Ruta al .key
        password:         Contraseña del .key
        anio:             Año del período
        mes:              Mes del período (1-12)
        carpeta_xml_base: Carpeta raíz de XMLs del cliente en V5
                          Ejemplo: "C:/ConciliaRFC/clientes/SALE770801RH9/xml"
                          Se crearán subcarpetas: .../xml/emitidos/ y .../xml/recibidos/
        callback:         Función de progreso para la UI (opcional)

    Returns:
        dict con claves:
            ok (bool): True si descarga exitosa
            total_cfdi (int): Total de CFDIs descargados
            emitidos (list[str]): Paths de XMLs emitidos
            recibidos (list[str]): Paths de XMLs recibidos
            errores (list[str]): Lista de errores si los hubo
            resumen (str): Mensaje legible para el usuario
    """
    try:
        signer = cargar_fiel(cer_path, key_path, password)
        resultado = descargar_mes(
            signer=signer,
            anio=anio,
            mes=mes,
            carpeta_destino=Path(carpeta_xml_base) / f"{anio:04d}" / f"{mes:02d}",
            callback=callback,
        )
        return {
            "ok": resultado.ok,
            "total_cfdi": resultado.total_cfdi,
            "emitidos": [str(p) for p in resultado.emitidos],
            "recibidos": [str(p) for p in resultado.recibidos],
            "errores": resultado.errores,
            "resumen": resultado.resumen(),
        }
    except DescargaSATError as exc:
        msg = str(exc)
        _log(callback, f"[SAT] ❌ {msg}")
        return {
            "ok": False,
            "total_cfdi": 0,
            "emitidos": [],
            "recibidos": [],
            "errores": [msg],
            "resumen": f"❌ Error: {msg}",
        }
