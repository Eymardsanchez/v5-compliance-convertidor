"""
sat_descarga_portal.py — Canal de respaldo: robot sobre el portal web del SAT
(Descarga masiva de CFDI), vía Playwright + FIEL.

Complementa a sat_descarga.py (Web Service SOAP, canal principal recomendado
— hasta 200,000 XML por solicitud). Este módulo es el respaldo manual para
cuando el Web Service no está disponible: el portal tiene su propia cuota,
separada de la del Web Service (confirmado con tutorial oficial de OneFacture,
ver CLAUDE.md sección "Canal Portal").

Login (verificado en vivo, 10 jul 2026, solo lectura de estructura, sin
autenticar): https://cfdiau.sat.gob.mx/nidp/app/login?id=SATx509Custom
    #fileCertificate / #filePrivateKey  input[type=file]  → .cer / .key
    #privateKeyPassword                 input[type=password]
    #submit                             compartido con el form de contraseña
                                         (oculto al cambiar a e.firma) — usar
                                         SIEMPRE el #submit VISIBLE.
No pide captcha (a diferencia del login por RFC+contraseña).

Navegación posterior al login (portal real, NO cfdiau): tras autenticar, el
SAT aterriza en portalcfdi.facturaelectronica.sat.gob.mx — un portal ASP.NET
WebForms clásico con dos páginas separadas:
    ConsultaEmisor.aspx    → Emitidos, rango de fechas
    ConsultaReceptor.aspx  → Recibidos, UN DÍA a la vez (se itera día por día)
Los IDs de control y las peculiaridades del formulario (radio que exige clic
NATIVO, inputs de fecha disabled-por-diseño que se fuerzan por JS, resultados
vía AJAX UpdatePanel) se contrastaron contra "sat-dm" (github.com/IsraelCV5/
sat-dm — MIT, Israel Castro), un proyecto de terceros que documenta estos
mismos detalles a partir de pruebas reales contra el portal (may. 2026). Se
adoptaron los HECHOS verificados sobre la estructura pública del portal
(URLs, ids de control, gotchas de ASP.NET) — no se adoptó ninguna técnica de
ese proyecto orientada a evadir límites o controles anti-abuso del SAT.

Límite del portal: ~2,000 XML/día (separado de la cuota del Web Service SOAP)
y descargas individuales (una petición HTTP por CFDI, no un paquete ZIP como
el Web Service) — de ahí el corte defensivo tras fallos seguidos en
`_descargar_filas`.

Resiliencia a cambios de layout — dos capas, solo para los puntos más
"semánticos" del flujo (login, radio de fechas, botón buscar):
    1. SELECTORES: cada campo lógico tiene una lista de selectores de
       respaldo, probados en orden — el primero que exista, se usa. Costo
       cero, sin red adicional.
    2. Si TODOS los selectores conocidos de un campo fallan, se intenta
       auto-reparación con la API de Anthropic (_reparar_con_ia): se manda
       el HTML de la página pidiendo un selector CSS nuevo. Si funciona, se
       usa para esa corrida — pero SIEMPRE se registra en
       sat_portal_selectores.log (éxito o fracaso) para que una sesión de
       Claude Code lo convierta en entrada permanente de SELECTORES; la
       auto-reparación en vivo es un parche de emergencia, no un reemplazo
       del arreglo real. Sin API key configurada, o si la IA también falla,
       se aborta con DescargaSATError y el log completo del intento — nunca
       falla en silencio.
    Los ids de control MUY específicos de ASP.NET (dropdowns de año/mes/día/
    hora, inputs de fecha forzados por JS) no pasan por esta tabla — son
    demasiado granulares para que un selector de IA los distinga de forma
    fiable; si el SAT los cambia, fallan con un error claro y explícito en
    vez de intentar adivinar.

Dependencia: playwright>=1.60 (requiere una vez por máquina:
`playwright install chromium`), anthropic>=0.40 (opcional — solo se usa si
hay API key guardada; sin ella la auto-reparación simplemente no aplica).
"""

from __future__ import annotations

import hashlib
import json
import logging
import os
import re
import time
import uuid
import zipfile
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Callable, Optional

from playwright.sync_api import sync_playwright, Page, TimeoutError as PlaywrightTimeoutError

from sat_descarga import DescargaSATResult, DescargaSATError

logger = logging.getLogger(__name__)

# ─── Configuración ────────────────────────────────────────────────────────────
LOGIN_URL_FIEL = "https://cfdiau.sat.gob.mx/nidp/app/login?id=SATx509Custom"
PORTAL_URL = "https://portalcfdi.facturaelectronica.sat.gob.mx"
LIMITE_XML_DIA_PORTAL = 2000  # cuota del portal — separada de la del Web Service
LOGIN_TIMEOUT_MS = 180_000    # 3 min — la firma client-side de la e.firma puede tardar
DOWNLOAD_TIMEOUT_MS = 30_000
FALLOS_SEGUIDOS_MAX = 3       # cortar tras 3 descargas fallidas seguidas (¿cuota agotada?)
_TIPOS_VALIDOS = {"emitidos", "recibidos"}


# ─── Guardado cifrado de la API key de Anthropic ──────────────────────────────
# Mismo esquema que repo_conciliaciones.py: XOR-SHA256 atado al MAC de la
# máquina. La API key solo habilita la auto-reparación de selectores; su
# ausencia no rompe el resto del canal Portal.

def _app_dir() -> Path:
    p = Path(os.environ.get("APPDATA", Path.home())) / "ConvertidorBancos"
    p.mkdir(parents=True, exist_ok=True)
    return p


def _api_key_path() -> Path:
    return _app_dir() / "anthropic_key.bin"


def _cipher_key() -> bytes:
    return hashlib.sha256(str(uuid.getnode()).encode()).digest()


def _xor(data: bytes, key: bytes) -> bytes:
    return bytes(b ^ key[i % len(key)] for i, b in enumerate(data))


def guardar_api_key_anthropic(api_key: str) -> None:
    """Cifra y guarda la API key de Anthropic para auto-reparación de selectores."""
    blob = _xor(api_key.strip().encode("utf-8"), _cipher_key())
    _api_key_path().write_bytes(blob)


def cargar_api_key_anthropic() -> Optional[str]:
    """Devuelve la API key guardada, o None si no hay una configurada/legible."""
    path = _api_key_path()
    if not path.exists():
        return None
    try:
        return _xor(path.read_bytes(), _cipher_key()).decode("utf-8")
    except Exception:
        return None


def eliminar_api_key_anthropic() -> None:
    path = _api_key_path()
    if path.exists():
        path.unlink()


def hay_api_key_configurada() -> bool:
    return bool(cargar_api_key_anthropic())


# ─── Tabla de selectores resilientes (puntos semánticos del flujo) ────────────
# Cada campo lógico: lista de selectores Playwright (CSS o texto) probados en
# orden. El primero que exista en la página se usa.
SELECTORES: dict[str, list[str]] = {
    # Login — verificados contra el DOM real del SAT (10 jul 2026).
    "login_cer_input": ["#fileCertificate", "input[type=file]#fileCertificate"],
    "login_key_input": ["#filePrivateKey", "input[type=file]#filePrivateKey"],
    "login_password_input": [
        "#privateKeyPassword",
        "input[type=password]#privateKeyPassword",
    ],
    # #submit está DUPLICADO en la página (uno por cada pestaña de login);
    "login_submit": ["#submit:visible", "#submit", "text=Enviar"],
    "login_error": ["#msgError", ".msg-error", "#errormsg"],

    # Navegación del portal (portalcfdi.facturaelectronica.sat.gob.mx) —
    # contrastados contra sat-dm (ver docstring del módulo). Solo los puntos
    # más "semánticos"; los dropdowns de fecha van directo en el código.
    "radio_fechas": ["#ctl00_MainContent_RdoFechas"],
    "boton_buscar_portal": ["#ctl00_MainContent_BtnBusqueda"],
    "panel_sin_resultados": ["#ctl00_MainContent_PnlNoResultados"],
    "spinner_busqueda": ["#ctl00_MainContent_UpdateProgress1"],
    "filas_resultado": ["input.ListaFolios"],
}

_DESCRIPCIONES_CAMPO = {
    "login_cer_input": "el campo de archivo (input type=file) para subir el certificado .cer de la e.firma",
    "login_key_input": "el campo de archivo (input type=file) para subir la llave privada .key de la e.firma",
    "login_password_input": "el campo de contraseña de la llave privada de la e.firma",
    "login_submit": "el botón VISIBLE para enviar el formulario de login de e.firma (hay otro oculto, no usar ese)",
    "login_error": "un mensaje de error de login/autenticación fallida",
    "radio_fechas": "el radio button para buscar CFDIs 'Por fecha' (habilita los controles de fecha)",
    "boton_buscar_portal": "el botón 'Buscar CFDI' que envía la búsqueda",
    "panel_sin_resultados": "el panel que el portal muestra cuando la búsqueda no encontró CFDIs",
    "spinner_busqueda": "el indicador de carga que se muestra mientras la búsqueda AJAX procesa",
    "filas_resultado": "las filas (checkboxes) de la tabla de resultados de CFDIs encontrados",
}

LOG_SELECTORES_PATH = _app_dir() / "sat_portal_selectores.log"


def _log_selector_roto(campo: str, contexto: dict, resultado: str) -> None:
    linea = {
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "campo": campo,
        "contexto": contexto,
        "resultado": resultado,
    }
    with LOG_SELECTORES_PATH.open("a", encoding="utf-8") as f:
        f.write(json.dumps(linea, ensure_ascii=False) + "\n")
    logger.warning("[Portal SAT] selector roto — campo=%s resultado=%s", campo, resultado)


def _reparar_con_ia(page: Page, campo: str):
    """
    Último recurso cuando ningún selector conocido de `campo` existe en la
    página: pide a la API de Anthropic un selector CSS nuevo a partir del
    HTML actual. Requiere API key guardada (guardar_api_key_anthropic).

    SIEMPRE deja rastro en LOG_SELECTORES_PATH, funcione o no — el objetivo
    no es solo pasar esta corrida, sino que quede evidencia para convertir
    el hallazgo en una entrada permanente de SELECTORES.
    """
    contexto = {
        "descripcion": _DESCRIPCIONES_CAMPO.get(campo, campo),
        "selectores_probados": SELECTORES.get(campo, []),
        "url": page.url,
    }

    api_key = cargar_api_key_anthropic()
    if not api_key:
        _log_selector_roto(campo, contexto, resultado="sin API key de Anthropic configurada")
        raise DescargaSATError(
            f"[Portal SAT] El campo '{campo}' no se encontró con ningún selector conocido "
            f"(el layout del portal probablemente cambió) y no hay API key de Anthropic "
            f"configurada para intentar auto-reparación. Detalle en {LOG_SELECTORES_PATH}. "
            f"Requiere arreglo manual en SELECTORES (sat_descarga_portal.py)."
        )

    try:
        import anthropic

        html_resumido = page.content()[:20000]
        cliente = anthropic.Anthropic(api_key=api_key)
        resp = cliente.messages.create(
            model="claude-sonnet-5",
            max_tokens=200,
            messages=[{
                "role": "user",
                "content": (
                    "Esta es una página HTML del portal del SAT (México, gobierno). "
                    f"Necesito un selector CSS válido para Playwright que ubique: "
                    f"{_DESCRIPCIONES_CAMPO.get(campo, campo)}.\n"
                    "Responde ÚNICAMENTE el selector CSS en una línea, sin explicación, "
                    "sin backticks, sin comillas.\n\n"
                    f"HTML:\n{html_resumido}"
                ),
            }],
        )
        selector_sugerido = resp.content[0].text.strip().strip("`").strip()
        loc = page.locator(selector_sugerido).first
        if loc.count() == 0:
            raise ValueError(f"selector sugerido no encontró nada en la página: {selector_sugerido!r}")

        _log_selector_roto(
            campo, contexto,
            resultado=f"IA sugirió {selector_sugerido!r} y funcionó — agregar a SELECTORES manualmente",
        )
        return loc

    except Exception as exc:
        _log_selector_roto(campo, contexto, resultado=f"auto-reparación con IA falló: {exc}")
        raise DescargaSATError(
            f"[Portal SAT] El campo '{campo}' cambió de layout y la auto-reparación con IA "
            f"también falló ({exc}). Detalle completo en {LOG_SELECTORES_PATH}. "
            f"Requiere arreglo manual en SELECTORES (sat_descarga_portal.py)."
        ) from exc


def _resolver(page: Page, campo: str):
    """Primer Locator visible para `campo`; si ninguno de SELECTORES existe, intenta auto-reparación."""
    for sel in SELECTORES.get(campo, []):
        try:
            loc = page.locator(sel).first
            if loc.count() > 0:
                return loc
        except Exception:
            continue
    return _reparar_con_ia(page, campo)


def _existe(page: Page, campo: str) -> bool:
    """Como _resolver pero sin lanzar error ni auto-reparar — para checks opcionales."""
    for sel in SELECTORES.get(campo, []):
        try:
            if page.locator(sel).first.count() > 0:
                return True
        except Exception:
            continue
    return False


# ─── Login FIEL ────────────────────────────────────────────────────────────────
def _login_exitoso(url: str) -> bool:
    return PORTAL_URL in url and "cfdiau" not in url


def _login_fiel(page: Page, cer_path: Path, key_path: Path, password: str, callback: Optional[Callable]) -> None:
    _log(callback, "[Portal] Abriendo login FIEL del SAT...")
    page.goto(LOGIN_URL_FIEL, wait_until="domcontentloaded", timeout=30000)
    try:
        page.wait_for_load_state("networkidle", timeout=15000)
    except PlaywrightTimeoutError:
        pass

    # La URL de entrada normalmente aterriza directo en la pestaña e.firma;
    # por si el portal cambia esto, intento best-effort el botón que cambia
    # de "Acceso por contraseña" a e.firma.
    if not _existe(page, "login_cer_input"):
        try:
            page.click("#buttonFiel", timeout=5000)
        except Exception:
            pass

    _resolver(page, "login_cer_input").set_input_files(str(cer_path))
    _resolver(page, "login_key_input").set_input_files(str(key_path))
    _resolver(page, "login_password_input").fill(password)
    _resolver(page, "login_submit").click()

    try:
        page.wait_for_url(_login_exitoso, timeout=LOGIN_TIMEOUT_MS)
    except PlaywrightTimeoutError:
        if _existe(page, "login_error"):
            msg = _resolver(page, "login_error").text_content() or ""
            raise DescargaSATError(
                f"[Portal] El SAT rechazó el login FIEL: «{msg.strip()}». Verifica que el "
                f".cer, el .key y la contraseña correspondan al mismo RFC y estén vigentes."
            )
        raise DescargaSATError(
            f"[Portal] El login FIEL no aterrizó en el portal a tiempo. URL actual: {page.url}."
        )
    try:
        page.wait_for_load_state("networkidle", timeout=15000)
    except PlaywrightTimeoutError:
        pass

    _log(callback, f"[Portal] Login FIEL exitoso: {page.url}")


# ─── Descarga por rango (punto de entrada principal del canal Portal) ────────
def descargar_rango_portal(
    cer_path: str | Path,
    key_path: str | Path,
    password: str,
    fecha_inicial: date,
    fecha_final: date,
    carpeta_destino: str | Path,
    callback: Optional[Callable[[str], None]] = None,
    tipos: Optional[set[str]] = None,
    headless: bool = True,
) -> DescargaSATResult:
    """
    Descarga CFDIs vía el portal web del SAT (canal de respaldo), automatizado
    con Playwright + FIEL. Mismo resultado (DescargaSATResult) que
    sat_descarga.descargar_rango() — el llamador no necesita saber por cuál
    canal llegaron los XMLs.

    Emitidos se busca por rango de fechas (una sola solicitud); Recibidos
    solo admite un día a la vez en el portal, así que se itera día por día
    dentro del rango pedido — puede tardar bastante más que el Web Service
    para rangos largos.

    ⚠ Cuota del portal: máx. ~2,000 XML/día, separada de la cuota del Web
    Service SOAP (sat_descarga.py). Usar como respaldo, no como reemplazo.
    """
    tipos = tipos if tipos else set(_TIPOS_VALIDOS)
    invalidos = tipos - _TIPOS_VALIDOS
    if invalidos:
        raise ValueError(f"tipos inválido(s): {invalidos}. Usa 'emitidos'/'recibidos'.")

    cer_path = Path(cer_path)
    key_path = Path(key_path)
    carpeta = Path(carpeta_destino)
    carpeta.mkdir(parents=True, exist_ok=True)

    resultado = DescargaSATResult()
    _log(callback, f"[Portal] Período: {fecha_inicial} → {fecha_final} | Tipos: {', '.join(sorted(tipos))}")
    _log(callback, f"[Portal] ⚠ Cuota del portal: ~{LIMITE_XML_DIA_PORTAL} XML/día (independiente del Web Service).")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        try:
            context = browser.new_context(accept_downloads=True)
            page = context.new_page()
            _login_fiel(page, cer_path, key_path, password, callback)

            for tipo in sorted(tipos):
                out_dir = carpeta / tipo
                out_dir.mkdir(parents=True, exist_ok=True)
                _log(callback, f"[Portal] === {tipo.capitalize()} ===")
                try:
                    if tipo == "emitidos":
                        xmls = _descargar_emitidos_portal(page, fecha_inicial, fecha_final, out_dir, callback)
                    else:
                        xmls = _descargar_recibidos_portal(page, fecha_inicial, fecha_final, out_dir, callback)
                    setattr(resultado, tipo, xmls)
                except DescargaSATError as exc:
                    resultado.errores.append(f"{tipo.capitalize()}: {exc}")
                    _log(callback, f"[Portal] ⚠ Error en {tipo}: {exc}")
        finally:
            browser.close()

    _log(callback, f"[Portal] {resultado.resumen()}")
    return resultado


# ─── Búsqueda: Emitidos (rango de fechas) ─────────────────────────────────────
def _descargar_emitidos_portal(page: Page, fecha_inicial: date, fecha_final: date,
                                out_dir: Path, callback: Optional[Callable]) -> list[Path]:
    page.goto(f"{PORTAL_URL}/ConsultaEmisor.aspx", wait_until="domcontentloaded", timeout=30000)
    try:
        page.wait_for_load_state("networkidle", timeout=20000)
    except PlaywrightTimeoutError:
        pass

    _resolver(page, "radio_fechas").click()
    try:
        # Los inputs de fecha son disabled-por-diseño (date-picker); solo el
        # SELECT de hora se habilita de verdad tras el click del radio — es
        # la señal real de que el postback terminó (networkidle no basta).
        page.wait_for_function(
            "() => { const e = document.getElementById("
            "'ctl00_MainContent_CldFechaInicial2_DdlHora'); return e && !e.disabled; }",
            timeout=20000,
        )
    except PlaywrightTimeoutError:
        raise DescargaSATError(
            "[Portal] Los controles de fecha de Emitidos no se habilitaron tras elegir "
            "'Por fecha' — el portal pudo haber cambiado su formulario."
        )

    fi_s, ff_s = fecha_inicial.strftime("%d/%m/%Y"), fecha_final.strftime("%d/%m/%Y")
    page.evaluate(
        """([fi, ff]) => {
            const set = (id, v) => {
                const el = document.getElementById(id);
                if (!el) return;
                el.removeAttribute('disabled'); el.disabled = false;
                el.value = v;
                el.dispatchEvent(new Event('change', {bubbles: true}));
            };
            set('ctl00_MainContent_CldFechaInicial2_Calendario_text', fi);
            set('ctl00_MainContent_CldFechaFinal2_Calendario_text', ff);
        }""",
        [fi_s, ff_s],
    )
    _select_opcional(page, "#ctl00_MainContent_CldFechaInicial2_DdlHora", "0")
    _select_opcional(page, "#ctl00_MainContent_CldFechaInicial2_DdlMinuto", "0")
    _select_opcional(page, "#ctl00_MainContent_CldFechaInicial2_DdlSegundo", "0")
    _select_opcional(page, "#ctl00_MainContent_CldFechaFinal2_DdlHora", "23")
    _select_opcional(page, "#ctl00_MainContent_CldFechaFinal2_DdlMinuto", "59")
    _select_opcional(page, "#ctl00_MainContent_CldFechaFinal2_DdlSegundo", "59")

    _click_buscar_portal(page)

    filas = _extraer_filas_resultado(page)
    _log(callback, f"[Portal] Emitidos {fecha_inicial}→{fecha_final}: {len(filas)} CFDIs en resultados.")
    return _descargar_filas(page, filas, out_dir, "emitidos", callback)


# ─── Búsqueda: Recibidos (un día a la vez) ────────────────────────────────────
def _descargar_recibidos_portal(page: Page, fecha_inicial: date, fecha_final: date,
                                 out_dir: Path, callback: Optional[Callable]) -> list[Path]:
    descargados: list[Path] = []
    vistos: set[str] = set()
    dia = fecha_inicial
    while dia <= fecha_final:
        page.goto(f"{PORTAL_URL}/ConsultaReceptor.aspx", wait_until="domcontentloaded", timeout=30000)
        try:
            page.wait_for_load_state("networkidle", timeout=20000)
        except PlaywrightTimeoutError:
            pass

        _resolver(page, "radio_fechas").click()
        try:
            page.wait_for_load_state("networkidle", timeout=20000)
        except PlaywrightTimeoutError:
            pass

        _select_opcional(page, "select#DdlAnio", str(dia.year))
        _select_opcional(page, "select#ctl00_MainContent_CldFecha_DdlMes", str(dia.month))
        _select_opcional(page, "select#ctl00_MainContent_CldFecha_DdlDia", f"{dia.day:02d}")
        _select_opcional(page, "select#ctl00_MainContent_CldFecha_DdlHora", "0")
        _select_opcional(page, "select#ctl00_MainContent_CldFecha_DdlMinuto", "0")
        _select_opcional(page, "select#ctl00_MainContent_CldFecha_DdlSegundo", "0")
        _select_opcional(page, "select#ctl00_MainContent_CldFecha_DdlHoraFin", "23")
        _select_opcional(page, "select#ctl00_MainContent_CldFecha_DdlMinutoFin", "59")
        _select_opcional(page, "select#ctl00_MainContent_CldFecha_DdlSegundoFin", "59")

        _click_buscar_portal(page)

        filas = [f for f in _extraer_filas_resultado(page) if f["uuid"] not in vistos]
        if filas:
            _log(callback, f"[Portal] Recibidos {dia}: {len(filas)} CFDIs en resultados.")
            nuevos = _descargar_filas(page, filas, out_dir, "recibidos", callback)
            descargados.extend(nuevos)
            vistos.update(f["uuid"] for f in filas)
        dia += timedelta(days=1)

    return descargados


# ─── Helpers de búsqueda/descarga compartidos ─────────────────────────────────
def _click_buscar_portal(page: Page) -> None:
    """
    Clic en 'Buscar CFDI' y espera a que el UpdatePanel (AJAX) RENDERICE de
    verdad. El botón dispara un postback parcial; `networkidle` regresa antes
    de que las filas aparezcan (falso 0) — hay que esperar a que el spinner
    se oculte y aparezcan filas o el panel "sin resultados".
    """
    _resolver(page, "boton_buscar_portal").click()
    page.wait_for_timeout(800)  # tiempo a que arranque el postback
    try:
        page.wait_for_function(
            """() => {
                const filas = document.querySelectorAll('input.ListaFolios').length;
                const noRes = document.getElementById('ctl00_MainContent_PnlNoResultados');
                const sinResultados = !!(noRes && noRes.offsetParent !== null);
                const prog = document.getElementById('ctl00_MainContent_UpdateProgress1');
                const cargando = !!(prog && prog.offsetParent !== null);
                return !cargando && (filas > 0 || sinResultados);
            }""",
            timeout=60000,
        )
    except PlaywrightTimeoutError:
        logger.warning("[Portal] La búsqueda no terminó de renderizar a tiempo.")
    try:
        page.wait_for_load_state("networkidle", timeout=10000)
    except PlaywrightTimeoutError:
        pass


def _extraer_filas_resultado(page: Page) -> list[dict]:
    """
    Lee todas las filas de resultados (paginación client-side, todas están
    en el DOM). De cada una saca el UUID (input.ListaFolios) y la URL
    relativa RecuperaCfdi.aspx?Datos=<token> del ícono de descarga.
    """
    filas = page.evaluate(
        r"""() => {
            const cbs = Array.from(document.querySelectorAll('input.ListaFolios'));
            const spans = Array.from(document.querySelectorAll('span[name="BtnDescarga"]'));
            const out = [];
            for (let i = 0; i < cbs.length; i++) {
                let url = null;
                if (spans[i]) {
                    const oc = spans[i].getAttribute('onclick') || '';
                    const m = oc.match(/RecuperaCfdi\.aspx\?Datos=([^']+)/);
                    if (m) url = 'RecuperaCfdi.aspx?Datos=' + m[1];
                }
                out.push({ uuid: (cbs[i].value || '').toUpperCase(), url });
            }
            return out;
        }"""
    )
    for f in filas:
        if f["url"]:
            f["url"] = f"{PORTAL_URL}/{f['url']}"
    return filas


def _descargar_filas(page: Page, filas: list[dict], out_dir: Path, tipo: str,
                      callback: Optional[Callable]) -> list[Path]:
    """Descarga cada fila (UUID + URL RecuperaCfdi.aspx) a out_dir/{uuid}.xml."""
    descargados: list[Path] = []
    fallos_seguidos = 0

    for fila in filas:
        uuid_cfdi, url = fila["uuid"], fila["url"]
        if not uuid_cfdi:
            continue
        dest = out_dir / f"{uuid_cfdi}.xml"
        if dest.exists() and dest.stat().st_size > 0:
            descargados.append(dest)
            continue
        if not url:
            _log(callback, f"[Portal] ⚠ {uuid_cfdi} sin URL de descarga.")
            continue

        if _descargar_xml_portal(page, url, dest):
            descargados.append(dest)
            fallos_seguidos = 0
        else:
            fallos_seguidos += 1
            _log(callback, f"[Portal] ⚠ No se pudo descargar {uuid_cfdi}.")
            if fallos_seguidos >= FALLOS_SEGUIDOS_MAX:
                _log(callback,
                     f"[Portal] ⛔ {FALLOS_SEGUIDOS_MAX} descargas seguidas fallaron "
                     f"(¿cuota diaria del portal agotada?). Deteniendo {tipo}; "
                     f"reintenta más tarde o mañana.")
                break

    return descargados


def _descargar_xml_portal(page: Page, url: str, dest: Path) -> bool:
    """
    Descarga un XML navegando a RecuperaCfdi.aspx?Datos=<token>. El servidor
    responde con el adjunto; `goto` aborta la navegación pero
    `expect_download` captura el archivo (replica el window.open del portal).
    """
    try:
        with page.expect_download(timeout=DOWNLOAD_TIMEOUT_MS) as info:
            try:
                page.goto(url, wait_until="commit")
            except Exception:
                pass  # ERR_ABORTED esperado: la navegación se vuelve descarga
        descarga = info.value
        descarga.save_as(str(dest))
        return dest.exists() and dest.stat().st_size > 0
    except PlaywrightTimeoutError:
        return False
    except Exception as exc:
        logger.debug("[Portal] error descargando %s: %s", dest.name, exc)
        return False


def _select_opcional(page: Page, selector: str, valor: str) -> None:
    """Selecciona `valor` en un <select> si existe; no falla si no está (dropdown secundario)."""
    if page.query_selector(selector):
        try:
            page.select_option(selector, valor)
        except Exception as exc:
            logger.debug("[Portal] no se pudo seleccionar %s=%s: %s", selector, valor, exc)


def _log(callback: Optional[Callable], msg: str) -> None:
    logger.info(msg)
    if callback:
        callback(msg)
