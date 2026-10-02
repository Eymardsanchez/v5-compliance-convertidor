# -*- coding: utf-8 -*-
"""
ConvertidorBancos v5.0 — Compliance + Convertidor
==================================================
• PDF → Excel para múltiples bancos (detección automática)
• Gestión de Obligaciones Fiscales/Laborales por empresa
• Envío de correo multi-proveedor (Gmail, Outlook/Hotmail, Yahoo, SMTP Custom)
• Login con usuarios + catálogo de empresas por RFC
• Bitácora de conversiones
"""

import os
import re as _re
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext, simpledialog
try:
    import customtkinter as ctk           # UI moderna (tarjetas redondeadas)
    ctk.set_appearance_mode("light")      # se sincroniza con ui_theme en el toggle
    _HAS_CTK = True
except Exception:                         # respaldo si no está instalado
    ctk = None
    _HAS_CTK = False

# Catálogo c_FormaPago del SAT (para filtrar qué comprobantes conciliar)
FORMAS_PAGO = [
    ("01", "Efectivo"), ("02", "Cheque nominativo"),
    ("03", "Transferencia electrónica de fondos"), ("04", "Tarjeta de crédito"),
    ("05", "Monedero electrónico"), ("06", "Dinero electrónico"),
    ("08", "Vales de despensa"), ("12", "Dación en pago"),
    ("13", "Pago por subrogación"), ("14", "Pago por consignación"),
    ("15", "Condonación"), ("17", "Compensación"), ("23", "Novación"),
    ("24", "Confusión"), ("25", "Remisión de deuda"),
    ("26", "Prescripción o caducidad"), ("27", "A satisfacción del acreedor"),
    ("28", "Tarjeta de débito"), ("29", "Tarjeta de servicios"),
    ("30", "Aplicación de anticipos"), ("31", "Intermediario pagos"),
    ("99", "Por definir"),
]
from pathlib  import Path
from datetime import date, datetime
from dataclasses import dataclass

# ── Tema UI: DPI + dark mode — debe ir antes de crear cualquier ventana Tk ──
import ui_theme
ui_theme.enable_dpi()

from bancos.detector import detectar_banco, detectar_o_generico, PDF_LOCK
from exportar_excel  import export_to_excel
from usuarios import (
    hay_usuarios, crear_usuario, verificar_usuario,
    listar_usuarios, eliminar_usuario, cambiar_password,
    recuperar_password, tomar_nuevo_codigo, regenerar_codigo,
)
from empresas import (
    agregar_empresa, eliminar_empresa, listar_empresas,
    buscar_por_rfc, validar_rfc, actualizar_empresa,
    registrar_desde_cif, obtener_empresa,
)
from bitacora import (
    registrar as bitacora_registrar,
    listar    as bitacora_listar,
    total     as bitacora_total,
)
import compliance_engine as CE
from email_sender import (
    enviar_correo, probar_conexion,
    detectar_proveedor, PROVIDERS, config_proveedor,
)

# ── Persistencia de config de email ──────────────────────────────────────────
import json, hashlib, uuid

# ── Módulo de conciliación CFDI (opcional — no rompe el app si falta) ─────────
try:
    from cfdi_parser           import leer_cfdis_carpetas, clasificar_cfdis
    from conciliacion          import (mov_a_banco, conciliar_multibanco,
                                       FASES_CONCILIACION, FASES_DEFAULT)
    from exportar_conciliacion import exportar_conciliacion_multi as _exp_concil_multi
    _CONCIL_OK = True
except ImportError:
    _CONCIL_OK = False

# ── Módulo de Pólizas COI (opcional) ─────────────────────────────────────────
try:
    from ventana_polizas    import VentanaPolizas
    from catalogo_cuentas  import cargar_catalogo, importar_desde_excel as _imp_cuentas
    from catalogo_terceros import cargar_terceros, poblar_desde_cfdis as _poblar_terceros
    _COI_OK = True
except ImportError:
    _COI_OK = False

# ── Módulo de Grupos RFC (opcional) ──────────────────────────────────────────
try:
    from grupos_conciliacion import cargar_grupos as _cargar_grupos
    from ventana_grupos_rfc  import VentanaGruposRFC
    _GRUPOS_OK = True
except ImportError:
    _GRUPOS_OK = False

# ── Módulo de Descarga Masiva SAT (opcional — requiere satcfdi) ──────────────
try:
    import sat_descarga as _SAT
    _SAT_OK = True
except ImportError:
    _SAT_OK = False

# ── Canal de respaldo: robot sobre el portal web del SAT (opcional — requiere
# playwright con Chromium instalado; `anthropic` solo se usa si se configura
# una API key para auto-reparación de selectores) ────────────────────────────
try:
    import sat_descarga_portal as _SAT_PORTAL
    _SAT_PORTAL_OK = True
except ImportError:
    _SAT_PORTAL_OK = False

# ── Módulo de Saldos a favor ISR/IVA (opcional — motor sat_saldos) ───────────
try:
    from sat_saldos import pipeline as _SF_PIPE
    from sat_saldos.analyzer_iva_acumulado import etiqueta_mes as _sf_etiqueta_mes
    _SALDOS_OK = True
except Exception:
    _SALDOS_OK = False

def _app_dir():
    p = Path(os.environ.get("APPDATA", Path.home())) / "ConvertidorBancos"
    p.mkdir(parents=True, exist_ok=True)
    return p

# ── Carpeta destino de descargas SAT (configurable, fuera del perfil de
# usuario por default) ────────────────────────────────────────────────────
_SAT_CARPETA_DEFAULT = "C:/ConciliaRFC/SAT"

def _sat_config_path() -> Path:
    return _app_dir() / "sat_config.json"

def _sat_carpeta_base() -> Path:
    """Carpeta raíz donde se guardan las descargas SAT. Configurable desde
    el diálogo de descarga; si nunca se configuró, usa un default fuera de
    AppData/perfil de usuario para no acumular CFDIs ahí."""
    try:
        cfg = json.loads(_sat_config_path().read_text(encoding="utf-8"))
        ruta = cfg.get("carpeta_base", "").strip()
        if ruta:
            return Path(ruta)
    except Exception:
        pass
    return Path(_SAT_CARPETA_DEFAULT)

def _guardar_sat_carpeta_base(ruta: str) -> None:
    _sat_config_path().write_text(
        json.dumps({"carpeta_base": ruta}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

def _xor(data: bytes, key: bytes) -> bytes:
    return bytes(b ^ key[i % len(key)] for i, b in enumerate(data))

def _key() -> bytes:
    return hashlib.sha256(str(uuid.getnode()).encode()).digest()

def _cargar_email_cfg() -> dict:
    p = _app_dir() / "email_cfg.bin"
    try:
        if p.exists():
            return json.loads(_xor(p.read_bytes(), _key()).decode("utf-8"))
    except Exception:
        pass
    return {"proveedor": "gmail", "remitente": "", "password": "",
            "host_smtp": "smtp.gmail.com", "puerto": 587, "ssl": False}

def _guardar_email_cfg(cfg: dict) -> None:
    p = _app_dir() / "email_cfg.bin"
    p.write_bytes(_xor(json.dumps(cfg).encode("utf-8"), _key()))

# ─── Paleta y tipografía — centralizado en ui_theme ─────────────────────────
from ui_theme import (
    T,  F,
    C_DARK, C_ACCENT, C_BG, C_WHITE, C_GRAY, C_GRAY2, C_TEXT,
    C_GREEN, C_RED, C_RED_ERR, C_BLUE, C_WARN, C_HDR_FG,
    FONT_TITLE, FONT_HEAD, FONT_NORMAL, FONT_SMALL, FONT_MONO, FONT_BTN,
)

# ── FECHA DE CORTE helper ─────────────────────────────────────────────────────
_FECHA_RE = _re.compile(r'FECHA\s+DE\s+CORTE\s+(\d{1,2})/(\d{2})/(\d{4})', _re.IGNORECASE)
_MESES_NUM = {"ENE":"01","FEB":"02","MAR":"03","ABR":"04","MAY":"05","JUN":"06",
              "JUL":"07","AGO":"08","SEP":"09","OCT":"10","NOV":"11","DIC":"12"}

def _buscar_fecha_corte(pdf_path: str) -> str:
    import pdfplumber
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages[:4]:
                m = _FECHA_RE.search(page.extract_text() or "")
                if m:
                    return f"{m.group(3)}-{m.group(2)}"
    except Exception:
        pass
    return ""

def _año_mes_fallback(periodo: str) -> str:
    m = _re.findall(r'(\d{1,2})/(\d{2})/(\d{4})', periodo)
    if m:
        _, mes, año = m[-1]; return f"{año}-{mes}"
    m = _re.findall(r'(\d{1,2})/([A-Za-z]{3})/(\d{4})', periodo)
    if m:
        _, mt, año = m[-1]; return f"{año}-{_MESES_NUM.get(mt.upper(),'01')}"
    m = _re.search(r'(\d{4})-(\d{2})', periodo)
    if m: return f"{m.group(1)}-{m.group(2)}"
    return date.today().strftime("%Y-%m")

def _safe_folder(name: str) -> str:
    s = _re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', name).strip('. ')
    return s or "Empresa"

# ═══════════════════════════════════════════════════════════════════════════════
#  LOGIN / PRIMER USUARIO
# ═══════════════════════════════════════════════════════════════════════════════
class DialogLogin(tk.Tk):
    def __init__(self):
        super().__init__()
        # Aplicar tema inmediatamente (detecta light/dark de Windows)
        ui_theme.apply_theme(self)
        self.title("Acceso — ConvertidorBancos v5.0")
        self.resizable(False, False)
        self.configure(bg=T.BG)
        self.usuario = ""
        self._primer = not hay_usuarios()
        self._build()
        self._center()
        self.protocol("WM_DELETE_WINDOW", self._cancel)

    def _center(self):
        self.update_idletasks()
        mx, my, sw, sh = ui_theme.work_area(self)
        w = min(ui_theme.sx(400), int(sw * 0.95))
        h = min(self.winfo_reqheight() + 20, int(sh * 0.90))
        x = mx + (sw - w) // 2
        y = my + max(0, (sh - h) // 2 - 30)
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _build(self):
        # ── Barra de acento superior (3 px) ───────────────────────────────────
        tk.Frame(self, bg=T.ACCENT, height=3).pack(fill="x")

        # ── Encabezado ─────────────────────────────────────────────────────────
        hdr = tk.Frame(self, bg=T.ACCENT_SOFT, pady=22)
        hdr.pack(fill="x")
        tk.Label(hdr, text="ConvertidorBancos",
                 bg=T.ACCENT_SOFT, fg=T.ACCENT, font=("Segoe UI", 15, "bold")).pack()
        tk.Label(hdr, text="v5.0  ·  Obligaciones + Convertidor + Email",
                 bg=T.ACCENT_SOFT, fg=T.TEXT_MUTED, font=F.SMALL).pack(pady=(2, 0))
        tk.Label(
            hdr,
            text="Crear primer usuario" if self._primer else "Iniciar sesión",
            bg=T.ACCENT_SOFT, fg=T.TEXT_MUTED, font=F.BODY,
        ).pack(pady=(6, 0))

        # ── Cuerpo del formulario ──────────────────────────────────────────────
        body = tk.Frame(self, bg=T.SURFACE, padx=32, pady=24)
        body.pack(fill="both", expand=True)

        def campo(lbl_text, show=""):
            tk.Label(body, text=lbl_text, bg=T.SURFACE, fg=T.TEXT_MUTED,
                     font=F.SMALL, anchor="w").pack(fill="x")
            var = tk.StringVar()
            # Marco con borde inferior de acento al enfocar
            ef = tk.Frame(body, bg=T.BORDER, height=36)
            ef.pack(fill="x", pady=(2, 14))
            ef.pack_propagate(False)
            ent = tk.Entry(
                ef, textvariable=var, show=show,
                font=("Segoe UI", 10),
                bg=T.SURFACE, fg=T.TEXT,
                relief="flat", bd=0,
                insertbackground=T.ACCENT,
            )
            ent.pack(fill="both", expand=True, padx=10, pady=2)

            def _focus_in(_):  ef.config(bg=T.ACCENT)
            def _focus_out(_): ef.config(bg=T.BORDER)
            ent.bind("<FocusIn>",  _focus_in)
            ent.bind("<FocusOut>", _focus_out)
            return var

        self._vuser = campo("Usuario")
        self._vpwd  = campo("Contraseña", show="*")
        if self._primer:
            self._vpwd2 = campo("Confirmar contraseña", show="*")

        self._lbl_msg = tk.Label(
            body, text="", bg=T.SURFACE, fg=T.DANGER,
            font=(F.SMALL[0], F.SMALL[1], "bold"), wraplength=320,
        )
        self._lbl_msg.pack(pady=(0, 6))

        # ── Botones ────────────────────────────────────────────────────────────
        bf = tk.Frame(body, bg=T.SURFACE)
        bf.pack(fill="x", pady=(4, 0))

        tk.Button(
            bf, text="Cancelar", command=self._cancel,
            bg=T.BG, fg=T.TEXT_MUTED, font=F.BODY,
            relief="flat", padx=14, pady=8, cursor="hand2",
            bd=0, activebackground=T.BORDER, activeforeground=T.TEXT,
        ).pack(side="right", padx=(8, 0))

        lbl = "Crear usuario" if self._primer else "Entrar"
        tk.Button(
            bf, text=lbl, command=self._ok,
            bg=T.ACCENT, fg=T.ACCENT_FG, font=F.BTN,
            relief="flat", padx=22, pady=8, cursor="hand2",
            bd=0, activebackground=T.BLUE, activeforeground=T.ACCENT_FG,
        ).pack(side="right")

        if not self._primer:
            lnk = tk.Label(
                body, text="¿Olvidaste tu contraseña?",
                bg=T.SURFACE, fg=T.TEXT_MUTED,
                font=(F.SMALL[0], F.SMALL[1], "underline"), cursor="hand2",
            )
            lnk.pack(pady=(12, 0))
            lnk.bind("<Button-1>", lambda _: self._recuperar())
            lnk.bind("<Enter>",    lambda _: lnk.config(fg=T.ACCENT))
            lnk.bind("<Leave>",    lambda _: lnk.config(fg=T.TEXT_MUTED))

        # ── Borde inferior ─────────────────────────────────────────────────────
        tk.Frame(self, bg=T.BORDER, height=1).pack(fill="x")

        self._vuser.trace_add("write", lambda *_: self._lbl_msg.config(text=""))
        self._vpwd.trace_add("write",  lambda *_: self._lbl_msg.config(text=""))
        self.bind("<Return>", lambda _: self._ok())

    def _ok(self):
        user, pwd = self._vuser.get().strip(), self._vpwd.get()
        if self._primer:
            if pwd != self._vpwd2.get():
                self._lbl_msg.config(text="Las contraseñas no coinciden."); return
            ok, msg, code = crear_usuario(user, pwd)
            if not ok:
                self._lbl_msg.config(text=msg); return
            self.usuario = user.lower()
            self.withdraw()
            _DialogMostrarCodigo(self, self.usuario, code)
            self.destroy()
        else:
            if verificar_usuario(user, pwd):
                self.usuario = user.lower(); self.destroy()
            else:
                self._lbl_msg.config(text="Usuario o contraseña incorrectos.")

    def _recuperar(self):
        _DialogRecuperar(self)

    def _cancel(self):
        self.usuario = ""; self.destroy()


class _DialogMostrarCodigo(tk.Toplevel):
    def __init__(self, parent, usuario, code):
        super().__init__(parent)
        self.title("Código de recuperación")
        self.resizable(False, False)
        self.configure(bg=T.SURFACE)
        self.grab_set()
        # Barra superior de advertencia
        tk.Frame(self, bg=T.WARN, height=3).pack(fill="x")
        wrap = tk.Frame(self, bg=T.SURFACE, padx=28, pady=22)
        wrap.pack(fill="both", expand=True)
        tk.Label(wrap, text="Guarda este código de recuperación",
                 bg=T.SURFACE, fg=T.TEXT, font=F.HEAD).pack()
        tk.Label(wrap, text=f"Usuario: {usuario}",
                 bg=T.SURFACE, fg=T.TEXT_MUTED, font=F.BODY).pack(pady=(4, 0))
        # Código en callout amarillo
        code_frame = tk.Frame(wrap, bg=T.WARN_BG,
                              highlightthickness=1, highlightbackground=T.WARN)
        code_frame.pack(fill="x", pady=16)
        tk.Label(code_frame, text=code,
                 bg=T.WARN_BG, fg=T.WARN,
                 font=("Consolas", 20, "bold"), pady=12).pack()
        tk.Label(wrap,
                 text="Úsalo en Login → '¿Olvidaste tu contraseña?' para restablecer tu contraseña.\n"
                      "Este código solo se muestra UNA vez — no se puede recuperar.",
                 bg=T.SURFACE, fg=T.TEXT_MUTED, font=F.SMALL, justify="center",
                 wraplength=300).pack(pady=(0, 16))
        tk.Button(wrap, text="Entendido, ya lo guardé", command=self.destroy,
                  bg=T.SUCCESS, fg=T.ACCENT_FG, font=F.BTN,
                  relief="flat", padx=18, pady=8, cursor="hand2",
                  bd=0, activebackground=T.BLUE, activeforeground=T.ACCENT_FG,
                  ).pack()
        self.wait_window(self)


class _DialogRecuperar(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Recuperar contraseña")
        self.resizable(False, False)
        self.configure(bg=T.SURFACE)
        self.grab_set()
        tk.Frame(self, bg=T.ACCENT, height=3).pack(fill="x")
        tk.Label(self, text="Recuperar contraseña",
                 bg=T.SURFACE, fg=T.TEXT, font=F.TITLE).pack(pady=(18, 4))
        body = tk.Frame(self, bg=T.SURFACE, padx=30)
        body.pack(fill="x")

        def campo(lbl_text, show=""):
            tk.Label(body, text=lbl_text, bg=T.SURFACE, fg=T.TEXT_MUTED,
                     font=F.SMALL, anchor="w").pack(fill="x")
            var = tk.StringVar()
            ef = tk.Frame(body, bg=T.BORDER, height=34)
            ef.pack(fill="x", pady=(2, 10))
            ef.pack_propagate(False)
            ent = tk.Entry(ef, textvariable=var, show=show, font=("Segoe UI", 10),
                           bg=T.SURFACE, fg=T.TEXT, relief="flat", bd=0,
                           insertbackground=T.ACCENT)
            ent.pack(fill="both", expand=True, padx=10, pady=2)
            ent.bind("<FocusIn>",  lambda _: ef.config(bg=T.ACCENT))
            ent.bind("<FocusOut>", lambda _: ef.config(bg=T.BORDER))
            return var

        self._vu = campo("Usuario")
        self._vc = campo("Código de recuperación  (XXXX-XXXX-XXXX)")
        self._vp = campo("Nueva contraseña", show="*")
        self._vm = tk.Label(body, text="", bg=T.SURFACE, fg=T.DANGER,
                            font=F.SMALL, wraplength=320)
        self._vm.pack()
        bf = tk.Frame(body, bg=T.SURFACE); bf.pack(fill="x", pady=(8, 22))
        tk.Button(bf, text="Cancelar", command=self.destroy,
                  bg=T.BG, fg=T.TEXT_MUTED, font=F.BODY,
                  relief="flat", padx=12, pady=7, cursor="hand2", bd=0,
                  activebackground=T.BORDER, activeforeground=T.TEXT,
                  ).pack(side="right", padx=(8, 0))
        tk.Button(bf, text="Restablecer", command=self._ok,
                  bg=T.ACCENT, fg=T.ACCENT_FG, font=F.BTN,
                  relief="flat", padx=16, pady=7, cursor="hand2", bd=0,
                  activebackground=T.BLUE, activeforeground=T.ACCENT_FG,
                  ).pack(side="right")

    def _ok(self):
        ok, msg = recuperar_password(self._vu.get(), self._vc.get(), self._vp.get())
        if ok:
            nuevo = tomar_nuevo_codigo(self._vu.get().strip().lower())
            messagebox.showinfo("OK",
                f"Contraseña restablecida.\nNuevo código de recuperación:\n\n{nuevo}\n\n"
                "Guárdalo — no se vuelve a mostrar.", parent=self)
            self.destroy()
        else:
            self._vm.config(text=msg)


# ═══════════════════════════════════════════════════════════════════════════════
#  DESCARGA MASIVA SAT
# ═══════════════════════════════════════════════════════════════════════════════
def _filtrar_por_uuid_local(carpeta: Path, uuids_extra: list, log) -> Path:
    """
    Dado un lote ya descargado en `carpeta` (subcarpetas emitidos/, recibidos/),
    mueve a `carpeta/filtrado/` solo los XMLs cuyo UUID esté en `uuids_extra`.
    No borra nada del lote completo — solo separa el subconjunto pedido para
    cargarlo al conciliador. Retorna la carpeta filtrada.
    """
    from cfdi_parser import parsear_cfdi
    wanted  = {u.strip().upper() for u in uuids_extra if u.strip()}
    destino = carpeta / "filtrado"
    destino.mkdir(parents=True, exist_ok=True)
    encontrados = set()
    for xml_path in carpeta.rglob("*.xml"):
        if destino in xml_path.parents:
            continue
        try:
            cfdi = parsear_cfdi(xml_path)
        except Exception:
            continue
        if cfdi and cfdi.uuid.upper() in wanted:
            xml_path.rename(destino / xml_path.name)
            encontrados.add(cfdi.uuid.upper())
    faltantes = wanted - encontrados
    if faltantes:
        log(f"⚠ {len(faltantes)} UUID(s) no encontrados en el lote descargado: "
            + ", ".join(sorted(faltantes)))
    log(f"Filtro UUID: {len(encontrados)}/{len(wanted)} encontrados -> {destino}")
    return destino


class _DialogDescargaSAT(tk.Toplevel):
    """
    Descarga XMLs de CFDIs directamente del SAT (Descarga Masiva oficial)
    usando la e.firma del contribuyente, y los deja listos para Conciliación CFDI.

    Dos modos:
      • Por rango de fechas — una sola solicitud SAT cubre todo el período;
        opcionalmente se puede filtrar localmente por UUID después (gratis).
      • Por UUID específico  — cada UUID genera una solicitud SAT independiente
        y consume una unidad del límite diario (MAX_SOLICITUDES_DIA por RFC).

    La contraseña de la e.firma nunca se escribe a disco — solo vive en
    memoria mientras el diálogo está abierto.
    """

    def __init__(self, parent, rfc_sugerido: str, on_completado):
        super().__init__(parent)
        self.title("Descargar CFDIs del SAT")
        self.update_idletasks()
        mx, my, sw, sh = ui_theme.work_area(parent)
        w = min(ui_theme.sx(980), int(sw * 0.92))
        h = min(ui_theme.sx(760), int(sh * 0.90))
        x = mx + (sw - w) // 2
        y = my + max(0, (sh - h) // 2 - 20)
        self.geometry(f"{w}x{h}+{x}+{y}")
        self.minsize(min(ui_theme.sx(820), int(sw * 0.75)),
                     min(ui_theme.sx(560), int(sh * 0.75)))
        self.configure(bg=T.BG)
        self.resizable(True, True)
        self.grab_set()
        self._on_completado = on_completado

        tk.Frame(self, bg=T.ACCENT, height=3).pack(fill="x")
        hdr = tk.Frame(self, bg=T.ACCENT_SOFT, pady=14)
        hdr.pack(fill="x")
        tk.Label(hdr, text="📥  Descarga Masiva SAT", bg=T.ACCENT_SOFT, fg=T.ACCENT,
                 font=("Segoe UI", 13, "bold")).pack()
        tk.Label(hdr, text="Usa tu e.firma (FIEL) — la contraseña no se guarda en disco",
                 bg=T.ACCENT_SOFT, fg=T.TEXT_MUTED, font=F.SMALL).pack(pady=(2, 0))

        # ── Layout: barra lateral izquierda (login + filtros + ejecutar) ──────
        # y panel principal derecho (progreso + log) ──────────────────────────
        content = tk.Frame(self, bg=T.BG)
        content.pack(fill="both", expand=True)

        SIDEBAR_W = 320
        sidebar_wrap = tk.Frame(content, bg=T.BG, width=SIDEBAR_W)
        sidebar_wrap.pack(side="left", fill="y")
        sidebar_wrap.pack_propagate(False)

        sidebar_canvas = tk.Canvas(sidebar_wrap, bg=T.BG, highlightthickness=0,
                                    width=SIDEBAR_W)
        sidebar_scroll = ttk.Scrollbar(sidebar_wrap, orient="vertical", command=sidebar_canvas.yview)
        body = tk.Frame(sidebar_canvas, bg=T.BG, padx=16, pady=14)
        body_win = sidebar_canvas.create_window((0, 0), window=body, anchor="nw", width=SIDEBAR_W)
        body.bind("<Configure>", lambda e: sidebar_canvas.configure(scrollregion=sidebar_canvas.bbox("all")))
        sidebar_canvas.configure(yscrollcommand=sidebar_scroll.set)
        sidebar_canvas.pack(side="left", fill="both", expand=True)
        sidebar_scroll.pack(side="right", fill="y")

        def _on_mousewheel(event):
            sidebar_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        sidebar_canvas.bind("<Enter>", lambda e: sidebar_canvas.bind_all("<MouseWheel>", _on_mousewheel))
        sidebar_canvas.bind("<Leave>", lambda e: sidebar_canvas.unbind_all("<MouseWheel>"))

        tk.Frame(content, bg=T.BORDER, width=1).pack(side="left", fill="y")

        main_area = tk.Frame(content, bg=T.BG, padx=20, pady=14)
        main_area.pack(side="left", fill="both", expand=True)

        WRAP = SIDEBAR_W - 40

        def campo(parent_, lbl_text, show=""):
            tk.Label(parent_, text=lbl_text, bg=T.BG, fg=T.TEXT_MUTED,
                     font=F.SMALL, anchor="w").pack(fill="x", pady=(8, 0))
            var = tk.StringVar()
            ef = tk.Frame(parent_, bg=T.BORDER, height=32)
            ef.pack(fill="x", pady=(2, 0))
            ef.pack_propagate(False)
            ent = tk.Entry(ef, textvariable=var, show=show, font=("Segoe UI", 10),
                            bg=T.SURFACE, fg=T.TEXT, relief="flat", bd=0,
                            insertbackground=T.ACCENT)
            ent.pack(fill="both", expand=True, padx=8, pady=2)
            return var

        # ── RFC ──────────────────────────────────────────────────────────────
        tk.Label(body, text="RFC:", bg=T.BG, fg=T.TEXT, font=F.HEAD).pack(anchor="w")
        self._var_rfc = tk.StringVar(value=rfc_sugerido)
        rfcs = [rfc for rfc, _ in listar_empresas()]
        ttk.Combobox(body, textvariable=self._var_rfc, values=rfcs,
                     width=24, font=("Segoe UI", 10)).pack(fill="x", pady=(2, 0))

        # ── Carpeta destino ──────────────────────────────────────────────────
        tk.Frame(body, bg=T.BORDER, height=1).pack(fill="x", pady=10)
        tk.Label(body, text="Carpeta destino", bg=T.BG, fg=T.TEXT, font=F.HEAD).pack(anchor="w")
        self._var_carpeta = tk.StringVar(value=str(_sat_carpeta_base()))
        carp_f = tk.Frame(body, bg=T.BG); carp_f.pack(fill="x", pady=(6, 0))
        tk.Entry(carp_f, textvariable=self._var_carpeta, font=F.SMALL,
                 bg=T.SURFACE, fg=T.TEXT, relief="flat",
                 state="readonly").pack(side="left", fill="x", expand=True)
        tk.Button(carp_f, text="📂", command=self._elegir_carpeta_destino,
                  bg=T.SURFACE, fg=T.TEXT, relief="flat", padx=8, cursor="hand2",
                  highlightthickness=1, highlightbackground=T.BORDER).pack(side="left", padx=(6, 0))
        tk.Label(body,
                 text="Cada descarga crea {carpeta}/{RFC}/{fecha}/emitidos y /recibidos.",
                 bg=T.BG, fg=T.TEXT_MUTED, font=F.SMALL, wraplength=WRAP, justify="left"
                 ).pack(anchor="w", pady=(4, 0))

        # ── e.firma ──────────────────────────────────────────────────────────
        tk.Frame(body, bg=T.BORDER, height=1).pack(fill="x", pady=10)
        tk.Label(body, text="e.firma (FIEL)", bg=T.BG, fg=T.TEXT, font=F.HEAD).pack(anchor="w")

        self._var_cer = tk.StringVar()
        cer_f = tk.Frame(body, bg=T.BG); cer_f.pack(fill="x", pady=(6, 0))
        tk.Label(cer_f, text=".cer:", bg=T.BG, fg=T.TEXT_MUTED,
                 font=F.SMALL, width=5, anchor="w").pack(side="left")
        tk.Entry(cer_f, textvariable=self._var_cer, font=F.SMALL, bg=T.SURFACE, fg=T.TEXT,
                 relief="flat", state="readonly").pack(side="left", fill="x", expand=True, padx=6)
        tk.Button(cer_f, text="…", command=lambda: self._elegir_archivo(self._var_cer, ".cer"),
                  bg=T.SURFACE, fg=T.TEXT, relief="flat", padx=8, cursor="hand2",
                  highlightthickness=1, highlightbackground=T.BORDER).pack(side="left")

        self._var_key = tk.StringVar()
        key_f = tk.Frame(body, bg=T.BG); key_f.pack(fill="x", pady=(6, 0))
        tk.Label(key_f, text=".key:", bg=T.BG, fg=T.TEXT_MUTED,
                 font=F.SMALL, width=5, anchor="w").pack(side="left")
        tk.Entry(key_f, textvariable=self._var_key, font=F.SMALL, bg=T.SURFACE, fg=T.TEXT,
                 relief="flat", state="readonly").pack(side="left", fill="x", expand=True, padx=6)
        tk.Button(key_f, text="…", command=lambda: self._elegir_archivo(self._var_key, ".key"),
                  bg=T.SURFACE, fg=T.TEXT, relief="flat", padx=8, cursor="hand2",
                  highlightthickness=1, highlightbackground=T.BORDER).pack(side="left")

        self._var_pwd = campo(body, "Contraseña (.key):", show="*")

        # ── Modo de búsqueda ─────────────────────────────────────────────────
        tk.Frame(body, bg=T.BORDER, height=1).pack(fill="x", pady=10)
        tk.Label(body, text="Modo de búsqueda", bg=T.BG, fg=T.TEXT, font=F.HEAD).pack(anchor="w")
        self._var_modo = tk.StringVar(value="fecha")
        modo_f = tk.Frame(body, bg=T.BG); modo_f.pack(fill="x", pady=(2, 0))
        tk.Radiobutton(modo_f, text="Por rango de fechas", variable=self._var_modo, value="fecha",
                       command=self._actualizar_modo, bg=T.BG, fg=T.TEXT, font=F.BODY,
                       activebackground=T.BG, selectcolor=T.SURFACE, cursor="hand2", anchor="w"
                       ).pack(fill="x")
        tk.Radiobutton(modo_f, text="Por UUID específico", variable=self._var_modo, value="uuid",
                       command=self._actualizar_modo, bg=T.BG, fg=T.TEXT, font=F.BODY,
                       activebackground=T.BG, selectcolor=T.SURFACE, cursor="hand2", anchor="w"
                       ).pack(fill="x")

        # -- Panel: por fecha --
        self._panel_fecha = tk.Frame(body, bg=T.BG)

        # -- Canal: Web Service (principal) vs Portal (respaldo, robot) --
        self._var_canal = tk.StringVar(value="webservice")
        self._var_portal_visible = tk.BooleanVar(value=True)
        if _SAT_PORTAL_OK:
            tk.Label(self._panel_fecha, text="Canal:", bg=T.BG, fg=T.TEXT_MUTED,
                     font=F.SMALL).pack(anchor="w", pady=(0, 2))
            canal_f = tk.Frame(self._panel_fecha, bg=T.BG)
            canal_f.pack(fill="x", pady=(0, 4))
            tk.Radiobutton(canal_f, text="Web Service (recomendado)", variable=self._var_canal,
                           value="webservice", command=self._actualizar_canal, bg=T.BG, fg=T.TEXT,
                           font=F.BODY, activebackground=T.BG, selectcolor=T.SURFACE, cursor="hand2",
                           anchor="w").pack(fill="x")
            tk.Radiobutton(canal_f, text="Portal (respaldo, más lento)", variable=self._var_canal,
                           value="portal", command=self._actualizar_canal, bg=T.BG, fg=T.TEXT,
                           font=F.BODY, activebackground=T.BG, selectcolor=T.SURFACE, cursor="hand2",
                           anchor="w").pack(fill="x")
            tk.Label(self._panel_fecha,
                     text="El Portal tiene su propia cuota (~2,000 XML/día), separada de la "
                          "del Web Service — úsalo solo si el Web Service falla.",
                     bg=T.BG, fg=T.TEXT_MUTED, font=F.SMALL, wraplength=WRAP, justify="left"
                     ).pack(anchor="w", pady=(0, 8))

            self._chk_portal_visible = tk.Checkbutton(
                self._panel_fecha, text="Mostrar navegador al usar Portal "
                                         "(recomendado en la primera corrida)",
                variable=self._var_portal_visible, bg=T.BG, fg=T.TEXT, font=F.SMALL,
                activebackground=T.BG, selectcolor=T.SURFACE, cursor="hand2",
                wraplength=WRAP, justify="left", anchor="w")
            # Solo tiene sentido si el canal Portal está elegido — se muestra/oculta
            # dinámicamente desde _actualizar_canal().

        hoy = date.today()
        self._var_f_ini = campo(self._panel_fecha, "Fecha inicial (AAAA-MM-DD):")
        self._var_f_ini.set(hoy.replace(day=1).isoformat())
        self._var_f_fin = campo(self._panel_fecha, "Fecha final (AAAA-MM-DD):")
        self._var_f_fin.set(hoy.isoformat())

        tk.Label(self._panel_fecha, text="Tipo a descargar:", bg=T.BG, fg=T.TEXT_MUTED,
                 font=F.SMALL, anchor="w").pack(fill="x", pady=(8, 2))
        tipos_f = tk.Frame(self._panel_fecha, bg=T.BG); tipos_f.pack(fill="x")
        self._var_tipo_emitidos  = tk.BooleanVar(value=True)
        self._var_tipo_recibidos = tk.BooleanVar(value=True)
        tk.Checkbutton(tipos_f, text="Emitidos", variable=self._var_tipo_emitidos,
                       bg=T.BG, fg=T.TEXT, font=F.BODY, activebackground=T.BG,
                       selectcolor=T.SURFACE, cursor="hand2").pack(side="left")
        tk.Checkbutton(tipos_f, text="Recibidos", variable=self._var_tipo_recibidos,
                       bg=T.BG, fg=T.TEXT, font=F.BODY, activebackground=T.BG,
                       selectcolor=T.SURFACE, cursor="hand2").pack(side="left", padx=(8, 0))
        tk.Label(self._panel_fecha,
                 text="Cada tipo marcado es una solicitud SAT independiente — desmarca el "
                      "que no necesites para ahorrar cuota diaria.",
                 bg=T.BG, fg=T.TEXT_MUTED, font=F.SMALL, wraplength=WRAP, justify="left"
                 ).pack(anchor="w", pady=(4, 0))
        tk.Label(self._panel_fecha, text="Filtrar por UUID(s) tras descargar (opcional, uno por línea):",
                 bg=T.BG, fg=T.TEXT_MUTED, font=F.SMALL, anchor="w", wraplength=WRAP, justify="left"
                 ).pack(fill="x", pady=(8, 2))
        self._txt_filtro = tk.Text(self._panel_fecha, height=3, font=("Consolas", 9),
                                    bg=T.SURFACE, fg=T.TEXT, relief="flat",
                                    highlightthickness=1, highlightbackground=T.BORDER)
        self._txt_filtro.pack(fill="x")

        # -- Panel: por UUID --
        self._panel_uuid = tk.Frame(body, bg=T.BG)
        tk.Label(self._panel_uuid, text="UUID(s) — uno por línea:", bg=T.BG, fg=T.TEXT_MUTED,
                 font=F.SMALL, anchor="w").pack(fill="x", pady=(8, 2))
        self._txt_uuids = tk.Text(self._panel_uuid, height=4, font=("Consolas", 9),
                                   bg=T.SURFACE, fg=T.TEXT, relief="flat",
                                   highlightthickness=1, highlightbackground=T.BORDER)
        self._txt_uuids.pack(fill="x")
        _max = _SAT.MAX_SOLICITUDES_DIA if _SAT_OK else 10
        tk.Label(self._panel_uuid,
                 text=f"⚠ Cada UUID consume 1 de las {_max} solicitudes/día del RFC ante el SAT.",
                 bg=T.BG, fg=T.WARN, font=F.SMALL, wraplength=WRAP, justify="left"
                 ).pack(anchor="w", pady=(4, 0))

        self._panel_fecha.pack(fill="x", pady=(8, 0))
        self._actualizar_modo()
        if _SAT_PORTAL_OK:
            self._actualizar_canal()

        # ── Botón de ejecutar (al fondo de la barra lateral) ───────────────────
        tk.Frame(body, bg=T.BORDER, height=1).pack(fill="x", pady=10)
        self._btn_descargar = tk.Button(
            body, text="📥  Descargar", command=self._iniciar_descarga,
            bg=T.SUCCESS, fg=T.ACCENT_FG, font=F.BTN, relief="flat",
            padx=18, pady=8, cursor="hand2", bd=0)
        self._btn_descargar.pack(fill="x")
        tk.Button(body, text="Cerrar", command=self.destroy,
                  bg=T.SURFACE, fg=T.TEXT, font=F.BODY, relief="flat",
                  padx=14, pady=8, cursor="hand2", bd=0,
                  highlightthickness=1, highlightbackground=T.BORDER
                  ).pack(fill="x", pady=(6, 0))

        # ── Panel principal: progreso + log ─────────────────────────────────
        tk.Label(main_area, text="Progreso", bg=T.BG, fg=T.TEXT, font=F.HEAD).pack(anchor="w")
        self._prog = ttk.Progressbar(main_area, mode="indeterminate")
        self._prog.pack(fill="x", pady=(6, 0))
        self._log = tk.Text(main_area, font=F.MONO, bg="#1A1A19", fg="#C3C2B7",
                             relief="flat", state="disabled", wrap="word",
                             highlightthickness=1, highlightbackground=T.BORDER)
        self._log.pack(fill="both", expand=True, pady=(8, 0))

        self.protocol("WM_DELETE_WINDOW", self.destroy)

    def _actualizar_modo(self):
        if self._var_modo.get() == "fecha":
            self._panel_uuid.pack_forget()
            self._panel_fecha.pack(fill="x", pady=(8, 0))
        else:
            self._panel_fecha.pack_forget()
            self._panel_uuid.pack(fill="x", pady=(8, 0))

    def _actualizar_canal(self):
        """Muestra el checkbox 'Mostrar navegador' solo si el canal elegido es Portal."""
        if self._var_canal.get() == "portal":
            self._chk_portal_visible.pack(anchor="w", pady=(0, 8))
        else:
            self._chk_portal_visible.pack_forget()

    def _elegir_archivo(self, var: tk.StringVar, ext: str):
        path = filedialog.askopenfilename(
            title=f"Seleccionar archivo {ext}",
            filetypes=[(ext, f"*{ext}"), ("Todos", "*.*")],
            parent=self,
        )
        if path:
            var.set(path)

    def _elegir_carpeta_destino(self):
        carpeta = filedialog.askdirectory(
            title="Carpeta donde guardar las descargas SAT",
            initialdir=self._var_carpeta.get(),
            parent=self,
        )
        if carpeta:
            self._var_carpeta.set(carpeta)
            _guardar_sat_carpeta_base(carpeta)

    def _log_write(self, msg: str):
        def _do():
            self._log.config(state="normal")
            self._log.insert("end", msg + "\n")
            self._log.see("end")
            self._log.config(state="disabled")
        self.after(0, _do)

    def _iniciar_descarga(self):
        if not _SAT_OK:
            messagebox.showerror(
                "Módulo no disponible",
                "Falta el paquete 'satcfdi'. Instala con: pip install satcfdi",
                parent=self)
            return

        rfc      = self._var_rfc.get().strip().upper()
        cer      = self._var_cer.get().strip()
        key      = self._var_key.get().strip()
        pwd      = self._var_pwd.get()
        carpeta_base = self._var_carpeta.get().strip() or _SAT_CARPETA_DEFAULT

        if not rfc:
            messagebox.showwarning("RFC requerido", "Selecciona o escribe el RFC.", parent=self)
            return
        if not cer or not key:
            messagebox.showwarning("e.firma requerida",
                                   "Selecciona los archivos .cer y .key.", parent=self)
            return
        if not pwd:
            messagebox.showwarning("Contraseña requerida",
                                   "Ingresa la contraseña de la clave privada.", parent=self)
            return

        modo = self._var_modo.get()
        tipos = None
        if modo == "fecha":
            try:
                f_ini = date.fromisoformat(self._var_f_ini.get().strip())
                f_fin = date.fromisoformat(self._var_f_fin.get().strip())
            except ValueError:
                messagebox.showwarning("Fecha inválida", "Usa el formato AAAA-MM-DD.", parent=self)
                return
            if f_ini > f_fin:
                messagebox.showwarning("Rango inválido",
                                       "La fecha inicial debe ser anterior a la final.", parent=self)
                return
            tipos = set()
            if self._var_tipo_emitidos.get():  tipos.add("emitidos")
            if self._var_tipo_recibidos.get(): tipos.add("recibidos")
            if not tipos:
                messagebox.showwarning("Tipo requerido",
                                       "Marca al menos Emitidos o Recibidos.", parent=self)
                return
            extra = [u.strip() for u in self._txt_filtro.get("1.0", "end").splitlines() if u.strip()]
        else:
            f_ini = f_fin = None
            extra = [u.strip() for u in self._txt_uuids.get("1.0", "end").splitlines() if u.strip()]
            if not extra:
                messagebox.showwarning("UUID requerido", "Ingresa al menos un UUID.", parent=self)
                return

        canal = self._var_canal.get() if (modo == "fecha" and _SAT_PORTAL_OK) else "webservice"
        visible = self._var_portal_visible.get() if _SAT_PORTAL_OK else False

        self._btn_descargar.config(state="disabled", text="Descargando…")
        self._prog.start(10)
        self._log.config(state="normal"); self._log.delete("1.0", "end"); self._log.config(state="disabled")

        threading.Thread(
            target=self._ejecutar_descarga,
            args=(rfc, cer, key, pwd, modo, f_ini, f_fin, extra, carpeta_base, tipos, canal, visible),
            daemon=True,
        ).start()

    def _ejecutar_descarga(self, rfc, cer, key, pwd, modo, f_ini, f_fin, extra, carpeta_base, tipos, canal="webservice", visible=False):
        try:
            carpeta = Path(carpeta_base) / rfc / datetime.now().strftime("%Y%m%d_%H%M%S")

            if modo == "fecha" and canal == "portal":
                self._log_write("[Portal] Usando canal de respaldo (robot sobre el portal web del SAT)...")
                resultado = _SAT_PORTAL.descargar_rango_portal(
                    cer_path=cer, key_path=key, password=pwd,
                    fecha_inicial=f_ini, fecha_final=f_fin,
                    carpeta_destino=carpeta, callback=self._log_write,
                    tipos=tipos, headless=not visible,
                )
                carpeta_final = carpeta
                if extra:
                    self._log_write(f"Filtrando localmente por {len(extra)} UUID(s)...")
                    carpeta_final = _filtrar_por_uuid_local(carpeta, extra, self._log_write)
                self.after(0, lambda: self._finalizar(carpeta_final, resultado))
                return

            signer = _SAT.cargar_fiel(cer, key, pwd)
            if signer.rfc.upper() != rfc:
                self._log_write(
                    f"⚠ Aviso: la e.firma pertenece a {signer.rfc}, "
                    f"distinto del RFC seleccionado ({rfc})."
                )

            if modo == "fecha":
                resultado = _SAT.descargar_rango(
                    signer=signer, fecha_inicial=f_ini, fecha_final=f_fin,
                    carpeta_destino=carpeta, callback=self._log_write,
                    tipos=tipos,
                )
                carpeta_final = carpeta
                if extra:
                    self._log_write(f"Filtrando localmente por {len(extra)} UUID(s)...")
                    carpeta_final = _filtrar_por_uuid_local(carpeta, extra, self._log_write)
            else:
                resultado = _SAT.descargar_por_uuid(
                    signer=signer, uuids=extra, carpeta_destino=carpeta,
                    callback=self._log_write,
                )
                carpeta_final = carpeta / "por_uuid"

            self.after(0, lambda: self._finalizar(carpeta_final, resultado))

        except _SAT.DescargaSATError as exc:
            self.after(0, lambda: self._error(str(exc)))
        except Exception as exc:
            self.after(0, lambda: self._error(f"Error inesperado: {exc}"))

    def _finalizar(self, carpeta_final: Path, resultado):
        self._prog.stop()
        self._btn_descargar.config(state="normal", text="📥  Descargar")

        if resultado.total_cfdi == 0:
            if resultado.errores:
                messagebox.showerror("Errores en la descarga",
                                     "\n".join(resultado.errores), parent=self)
            else:
                messagebox.showinfo(
                    "Sin resultados",
                    "El SAT no devolvió CFDIs para los criterios indicados.",
                    parent=self)
            return

        self._on_completado(carpeta_final, resultado)
        messagebox.showinfo(
            "Descarga completa",
            f"{resultado.resumen()}\n\nSe agregaron a Conciliación CFDI (Paso 1).",
            parent=self,
        )

    def _error(self, msg: str):
        self._prog.stop()
        self._btn_descargar.config(state="normal", text="📥  Descargar")
        self._log_write(f"❌ {msg}")
        messagebox.showerror("Error de descarga SAT", msg, parent=self)


# ═══════════════════════════════════════════════════════════════════════════════
#  DIÁLOGO: GESTIÓN DE EMPRESAS (RFC)
# ═══════════════════════════════════════════════════════════════════════════════
class DialogEmpresas(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Catálogo de Empresas / RFC")
        self.resizable(True, False)
        self.configure(bg=C_BG)
        self.grab_set()
        self._build()
        self._refresh()
        self._center(parent)

    def _center(self, parent):
        self.update_idletasks()
        w, h = 640, self.winfo_reqheight()
        sw, sh = ui_theme.screen_size(parent)
        w = min(ui_theme.sx(640), int(sw * 0.90))
        h = min(self.winfo_reqheight(), int(sh * 0.88))
        x = max(0, parent.winfo_x() + (parent.winfo_width()  - w) // 2)
        y = max(0, parent.winfo_y() + (parent.winfo_height() - h) // 2)
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _build(self):
        hdr = tk.Frame(self, bg=T.ACCENT, pady=12); hdr.pack(fill="x")
        tk.Label(hdr, text="Catálogo de Empresas",
                 bg=T.ACCENT, fg=T.ACCENT_FG, font=("Segoe UI", 11, "bold")).pack(side="left", padx=16)
        tk.Label(hdr, text="RFC · Nombre · Régimen fiscal (CSF)",
                 bg=T.ACCENT, fg=T.ACCENT_FG_2, font=FONT_SMALL).pack(side="left")

        tf = tk.Frame(self, bg=C_BG); tf.pack(fill="both", expand=True, padx=14, pady=10)
        cols = ("RFC", "Nombre de Empresa", "Régimen(es) SAT")
        self._tree = ttk.Treeview(tf, columns=cols, show="headings", height=10, selectmode="browse")
        for c, w in zip(cols, [140, 240, 220]):
            self._tree.heading(c, text=c, anchor="w")
            self._tree.column(c, width=w, minwidth=80)
        self._tree.tag_configure("sin_csf", foreground="#9E9E9E")
        sb = ttk.Scrollbar(tf, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=sb.set)
        self._tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self._tree.bind("<<TreeviewSelect>>", self._on_select)

        form = tk.Frame(self, bg=C_WHITE, padx=14, pady=10); form.pack(fill="x")
        lbs = ["RFC:", "Nombre:"]
        self._var_rfc    = tk.StringVar()
        self._var_nombre = tk.StringVar()
        self._var_rfc.trace_add("write", self._on_rfc_change)
        for i, (lb, var) in enumerate(zip(lbs, [self._var_rfc, self._var_nombre])):
            tk.Label(form, text=lb, bg=C_WHITE, fg=C_TEXT, font=FONT_HEAD).grid(
                row=i, column=0, sticky="w", pady=3)
            ef = tk.Frame(form, bg=C_GRAY)
            ef.grid(row=i, column=1, sticky="ew", padx=(8, 20), pady=3)
            kw = {"font": ("Consolas", 10, "bold"), "width": 16} if i == 0 else {"font": ("Segoe UI", 10), "width": 40}
            tk.Entry(ef, textvariable=var, bg=C_WHITE, fg=C_TEXT, relief="flat", bd=5,
                     insertbackground=C_ACCENT, **kw).pack(fill="x")
        self._lbl_rfc_msg = tk.Label(form, text="", bg=C_WHITE, fg=C_RED, font=FONT_SMALL)
        self._lbl_rfc_msg.grid(row=2, column=1, sticky="w", padx=8)

        # Fila de régimen (solo lectura — se llena al seleccionar)
        tk.Label(form, text="Régimen(es):", bg=C_WHITE, fg=C_TEXT,
                 font=FONT_HEAD).grid(row=3, column=0, sticky="nw", pady=(6, 3))
        self._lbl_reg_info = tk.Label(form, text="", bg=C_WHITE, fg=C_GRAY2,
                                      font=FONT_SMALL, anchor="w", justify="left", wraplength=420)
        self._lbl_reg_info.grid(row=3, column=1, sticky="w", padx=8, pady=(6, 3))
        form.columnconfigure(1, weight=1)

        bf = tk.Frame(self, bg=C_BG, padx=14, pady=8); bf.pack(fill="x")
        for txt, cmd, bg, fg in [
            ("➕ Nuevo",      self._nuevo,    C_BLUE,    C_WHITE),
            ("💾 Guardar",    self._guardar,  C_GREEN,   C_WHITE),
            ("📄 Subir CSF",  self._subir_cif,"#5C6BC0", C_WHITE),
            ("🗑 Eliminar",   self._eliminar, C_RED,     C_WHITE),
            ("✖ Cerrar",     self.destroy,   C_WHITE,   C_TEXT),
        ]:
            _hl = dict(highlightthickness=1, highlightbackground=T.BORDER) if bg == C_WHITE else {}
            tk.Button(bf, text=txt, command=cmd, bg=bg, fg=fg,
                      font=FONT_NORMAL, relief="flat", padx=12, pady=6,
                      cursor="hand2", **_hl).pack(side="left", padx=(0, 8))

    def _on_rfc_change(self, *_):
        rfc = self._var_rfc.get().upper()
        self._var_rfc.set(rfc)
        ok  = validar_rfc(rfc)
        msg = "" if (ok or not rfc) else "⚠ Formato: AAA000000AA0 (moral) o AAAA000000AA0 (física)"
        self._lbl_rfc_msg.config(text=msg)

    def _nuevo(self):
        """Limpia el formulario para capturar una empresa nueva."""
        self._tree.selection_remove(self._tree.selection())
        self._var_rfc.set("")
        self._var_nombre.set("")
        self._lbl_rfc_msg.config(text="")

    def _refresh(self):
        for row in self._tree.get_children():
            self._tree.delete(row)
        for rfc, nombre in listar_empresas():
            emp     = obtener_empresa(rfc) or {}
            codigos = emp.get("regimenes", [])
            nombres = emp.get("regimenes_nombres", [])
            if codigos and nombres:
                reg_txt = ",  ".join(f"{n} ({c})" for n, c in zip(nombres, codigos))
                tag = ""
            elif codigos:
                reg_txt = ", ".join(codigos)
                tag = ""
            else:
                reg_txt = "—  sin CSF"
                tag = "sin_csf"
            self._tree.insert("", "end", values=(rfc, nombre, reg_txt), tags=(tag,))

    def _on_select(self, _=None):
        sel = self._tree.selection()
        if not sel:
            return
        vals   = self._tree.item(sel[0])["values"]
        rfc    = str(vals[0])
        nombre = str(vals[1])
        self._var_rfc.set(rfc)
        self._var_nombre.set(nombre)
        emp     = obtener_empresa(rfc) or {}
        codigos = emp.get("regimenes", [])
        nombres = emp.get("regimenes_nombres", [])
        if codigos and nombres:
            partes = [f"{n} ({c})" for n, c in zip(nombres, codigos)]
            self._lbl_reg_info.config(
                text="\n".join(f"  • {p}" for p in partes), fg="#2E7D32")
        elif codigos:
            self._lbl_reg_info.config(text="  " + ",  ".join(codigos), fg="#2E7D32")
        else:
            self._lbl_reg_info.config(
                text="  Sin CSF registrada — sube la Constancia de Situación Fiscal",
                fg="#E65100")

    def _subir_cif(self):
        """Parsea una CSF PDF y registra/actualiza la empresa con su régimen fiscal."""
        pdf = filedialog.askopenfilename(
            title="Seleccionar Constancia de Situación Fiscal (PDF)",
            filetypes=[("PDF", "*.pdf"), ("Todos", "*.*")],
            parent=self,
        )
        if not pdf:
            return
        try:
            ok, msg, datos = registrar_desde_cif(pdf)
            if ok and datos:
                self._refresh()
                # Seleccionar la empresa recién registrada en el árbol
                for item in self._tree.get_children():
                    if str(self._tree.item(item)["values"][0]) == datos.rfc:
                        self._tree.selection_set(item)
                        self._tree.see(item)
                        self._on_select()
                        break
                messagebox.showinfo("CSF cargada", msg, parent=self)
            else:
                messagebox.showerror("Error al leer CSF", msg, parent=self)
        except Exception as e:
            messagebox.showerror("Error", str(e), parent=self)

    def _guardar(self):
        rfc    = self._var_rfc.get().strip().upper()
        nombre = self._var_nombre.get().strip()
        if not rfc or not nombre:
            messagebox.showwarning("Campos vacíos", "RFC y nombre son obligatorios.", parent=self); return
        if not validar_rfc(rfc):
            messagebox.showwarning("RFC inválido",
                f"El RFC '{rfc}' no tiene el formato correcto.\n"
                "Persona Moral: AAA000000AA0 (12 chars)\n"
                "Persona Física: AAAA000000AA0 (13 chars)", parent=self); return
        existente = buscar_por_rfc(rfc)
        if existente:
            if messagebox.askyesno("Actualizar",
                    f"RFC {rfc} ya existe ({existente}).\n¿Actualizar nombre?", parent=self):
                actualizar_empresa(rfc, nombre)
                messagebox.showinfo("Actualizado", f"✅ Empresa '{nombre}' actualizada.", parent=self)
        else:
            agregar_empresa(rfc, nombre)
            messagebox.showinfo("Guardado", f"✅ Empresa '{nombre}' registrada con RFC {rfc}.", parent=self)
        self._refresh()
        self._nuevo()   # limpiar formulario tras guardar

    def _eliminar(self):
        sel = self._tree.selection()
        if not sel:
            messagebox.showinfo("Selecciona", "Elige una empresa de la lista.", parent=self); return
        rfc, nombre = self._tree.item(sel[0])["values"]
        if messagebox.askyesno("Eliminar", f"¿Eliminar '{nombre}' ({rfc})?", parent=self):
            eliminar_empresa(rfc); self._refresh()
            self._nuevo()


# ═══════════════════════════════════════════════════════════════════════════════
#  DIÁLOGO: BITÁCORA
# ═══════════════════════════════════════════════════════════════════════════════
class _DialogBitacora(tk.Toplevel):
    _COLS = [
        ("Fecha / Hora",  "fecha_hora",    160),
        ("RFC",           "rfc",           120),
        ("Empresa",       "empresa",       200),
        ("Banco",         "banco",          80),
        ("Periodo",       "periodo",        70),
        ("Movimientos",   "movimientos",    90),
        ("Usuario",       "usuario",        90),
        ("Archivo PDF",   "archivo_origen",200),
    ]
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Bitácora de Conversiones")
        self.configure(bg=C_BG); self.resizable(True, True); self.grab_set()
        self._build(); self._cargar()
        sw, sh = ui_theme.screen_size(parent)
        w = min(ui_theme.sx(1060), int(sw * 0.90))
        h = min(ui_theme.sx(520),  int(sh * 0.80))
        x = max(0, parent.winfo_x() + (parent.winfo_width()  - w) // 2)
        y = max(0, parent.winfo_y() + (parent.winfo_height() - h) // 2)
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _build(self):
        hdr = tk.Frame(self, bg=T.ACCENT, padx=16, pady=10); hdr.pack(fill="x")
        tk.Label(hdr, text="Bitácora de Conversiones",
                 bg=T.ACCENT, fg=T.ACCENT_FG, font=FONT_TITLE).pack(side="left")
        self._lbl_total = tk.Label(hdr, text="", bg=T.ACCENT, fg=T.ACCENT_FG_2, font=FONT_NORMAL)
        self._lbl_total.pack(side="left", padx=16)

        flt = tk.Frame(self, bg=C_BG, padx=12, pady=8); flt.pack(fill="x")
        tk.Label(flt, text="RFC:", bg=C_BG, fg=C_TEXT, font=FONT_HEAD).pack(side="left")
        self._var_rfc = tk.StringVar()
        tk.Entry(flt, textvariable=self._var_rfc, width=16,
                 font=FONT_MONO, bg=C_WHITE, fg=C_TEXT, relief="flat", bd=4).pack(side="left", padx=(4, 16))
        tk.Label(flt, text="Banco:", bg=C_BG, fg=C_TEXT, font=FONT_HEAD).pack(side="left")
        self._var_banco = tk.StringVar()
        cb = ttk.Combobox(flt, textvariable=self._var_banco, state="readonly", width=12, font=FONT_NORMAL,
                          values=["Todos","BBVA","Santander","Banorte","Banamex","Inbursa","Banbajío","Banregio","HSBC"])
        cb.current(0); cb.pack(side="left", padx=(4, 16))
        tk.Label(flt, text="Periodo (YYYY-MM):", bg=C_BG, fg=C_TEXT, font=FONT_HEAD).pack(side="left")
        self._var_per = tk.StringVar()
        tk.Entry(flt, textvariable=self._var_per, width=9,
                 font=FONT_MONO, bg=C_WHITE, fg=C_TEXT, relief="flat", bd=4).pack(side="left", padx=(4, 16))
        tk.Button(flt, text="🔍 Filtrar", command=self._cargar,
                  bg=C_ACCENT, fg=C_TEXT, font=FONT_NORMAL,
                  relief="flat", padx=10, pady=4, cursor="hand2").pack(side="left", padx=(0, 8))
        tk.Button(flt, text="Limpiar", command=self._limpiar,
                  bg=C_WHITE, fg=C_TEXT, highlightthickness=1, highlightbackground=T.BORDER, font=FONT_NORMAL,
                  relief="flat", padx=10, pady=4, cursor="hand2").pack(side="left")

        tvf = tk.Frame(self, bg=C_BG); tvf.pack(fill="both", expand=True, padx=12, pady=(0, 8))
        cols = [c[0] for c in self._COLS]
        self._tv = ttk.Treeview(tvf, columns=cols, show="headings", selectmode="browse")
        # ttk.Style ya configurado globalmente por apply_theme()
        for label, _, width in self._COLS:
            self._tv.heading(label, text=label)
            self._tv.column(label, width=width, minwidth=60, anchor="w")
        self._tv.tag_configure("even", background=T.SURFACE_ALT)
        self._tv.tag_configure("odd",  background=T.SURFACE)
        sb_y = ttk.Scrollbar(tvf, orient="vertical",   command=self._tv.yview)
        sb_x = ttk.Scrollbar(tvf, orient="horizontal", command=self._tv.xview)
        self._tv.configure(yscrollcommand=sb_y.set, xscrollcommand=sb_x.set)
        self._tv.grid(row=0, column=0, sticky="nsew")
        sb_y.grid(row=0, column=1, sticky="ns")
        sb_x.grid(row=1, column=0, sticky="ew")
        tvf.grid_rowconfigure(0, weight=1); tvf.grid_columnconfigure(0, weight=1)

        bot = tk.Frame(self, bg=C_BG, padx=12, pady=8); bot.pack(fill="x")
        tk.Button(bot, text="Cerrar", command=self.destroy,
                  bg=C_WHITE, fg=C_TEXT, highlightthickness=1, highlightbackground=T.BORDER, font=FONT_NORMAL,
                  relief="flat", padx=14, pady=5, cursor="hand2").pack(side="right")

    def _cargar(self):
        rfc    = self._var_rfc.get().strip() or None
        banco  = None if self._var_banco.get() in ("", "Todos") else self._var_banco.get()
        per    = self._var_per.get().strip() or None
        filas  = bitacora_listar(rfc=rfc, banco=banco, periodo=per)
        for r in self._tv.get_children(): self._tv.delete(r)
        for i, e in enumerate(filas):
            vals = tuple(str(e.get(k, "")) for _, k, _ in self._COLS)
            self._tv.insert("", "end", values=vals, tags=(("even" if i%2==0 else "odd"),))
        self._lbl_total.config(text=f"{len(filas)} mostrado(s)  |  {bitacora_total()} en total")

    def _limpiar(self):
        self._var_rfc.set(""); self._var_banco.set("Todos"); self._var_per.set("")
        self._cargar()


# ═══════════════════════════════════════════════════════════════════════════════
#  DIÁLOGO: PREPARAR ENVÍO COMPLIANCE
# ═══════════════════════════════════════════════════════════════════════════════
class _DialogEnvioCompliance(tk.Toplevel):
    def __init__(self, parent, empresa_nombre, rfc, cfg_rfc, email_cfg, mes, anio, log_cb):
        super().__init__(parent)
        self.title(f"Preparar envío — {empresa_nombre} — {CE.MESES_ES[mes]} {anio}")
        self.update_idletasks()
        mx, my, sw, sh = ui_theme.work_area(parent)
        w = min(ui_theme.sx(900), int(sw * 0.88))
        h = min(ui_theme.sx(680), int(sh * 0.85))
        x = mx + (sw - w) // 2
        y = my + max(0, (sh - h) // 2 - 20)
        self.geometry(f"{w}x{h}+{x}+{y}")
        self.minsize(min(ui_theme.sx(780), int(sw * 0.70)), min(ui_theme.sx(520), int(sh * 0.70)))
        self.configure(bg=C_BG); self.grab_set()
        self._empresa = empresa_nombre; self._rfc = rfc
        self._cfg = cfg_rfc; self._email_cfg = email_cfg
        self._mes = mes; self._anio = anio; self._log = log_cb
        self._check_vars: dict[str, tk.BooleanVar] = {}
        self._archivos = CE.listar_archivos_mes(cfg_rfc, anio, mes)
        self._build()

    def _build(self):
        hdr = tk.Frame(self, bg=C_ACCENT); hdr.pack(fill="x")
        tk.Label(hdr, text=f"📤  {self._empresa}  —  {CE.MESES_ES[self._mes]} {self._anio}",
                 bg=C_ACCENT, fg=C_TEXT, font=("Segoe UI", 11, "bold")).pack(side="left", padx=16, pady=10)

        paned = tk.PanedWindow(self, orient="horizontal", bg=C_BG, sashwidth=6)
        paned.pack(fill="both", expand=True, padx=10, pady=8)

        # ── Izquierda: archivos ──
        izq = tk.Frame(paned, bg=C_BG); paned.add(izq, minsize=380)
        ruta = CE.carpeta_mes(self._cfg, self._anio, self._mes)
        info = tk.Frame(izq, bg=C_WHITE); info.pack(fill="x", pady=(0, 6))
        tk.Label(info, text="📁  Carpeta del mes:",
                 bg=C_WHITE, fg=C_TEXT, font=FONT_HEAD).pack(anchor="w", padx=10, pady=(8, 0))
        tk.Label(info, text=ruta if ruta else "⚠ Carpeta no encontrada",
                 bg=C_WHITE, fg=C_GRAY2 if ruta else C_RED,
                 font=FONT_SMALL, wraplength=360, justify="left").pack(anchor="w", padx=10, pady=(0, 8))

        hdr2 = tk.Frame(izq, bg=C_BG); hdr2.pack(fill="x")
        tk.Label(hdr2, text="Selecciona archivos a adjuntar:",
                 bg=C_BG, fg=C_TEXT, font=FONT_HEAD).pack(side="left")
        for txt, val in [("✓ Todos", True), ("✗ Ninguno", False)]:
            tk.Button(hdr2, text=txt, bg=C_GREEN if val else C_RED, fg=C_WHITE,
                      font=FONT_SMALL, relief="flat", padx=6, pady=2, cursor="hand2",
                      command=lambda v=val: self._sel_all(v)).pack(side="right", padx=(4, 0))

        lista_f = tk.Frame(izq, bg=C_BG); lista_f.pack(fill="both", expand=True)
        canvas = tk.Canvas(lista_f, bg=C_BG, bd=0, highlightthickness=0)
        sb = ttk.Scrollbar(lista_f, orient="vertical", command=canvas.yview)
        self._lista = tk.Frame(canvas, bg=C_BG)
        self._lista.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        _wid = canvas.create_window((0, 0), window=self._lista, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(_wid, width=e.width))
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        if not self._archivos:
            tk.Label(self._lista, text="Sin archivos en la carpeta del mes.\n"
                     "Usa '➕ Agregar' para añadir manualmente.",
                     bg=C_BG, fg=C_GRAY2, font=FONT_NORMAL, justify="center").pack(pady=30)
        else:
            for a in self._archivos: self._fila(a)

        tk.Button(izq, text="➕ Agregar archivos adicionales…",
                  bg="#EFF6FF", fg=C_BLUE, font=FONT_NORMAL, relief="flat",
                  padx=10, pady=6, cursor="hand2",
                  command=self._agregar).pack(fill="x", pady=(6, 0))

        # ── Derecha: preview ──
        der = tk.Frame(paned, bg=C_BG); paned.add(der, minsize=340)
        tk.Label(der, text="Vista previa del correo:",
                 bg=C_BG, fg=C_TEXT, font=FONT_HEAD).pack(anchor="w", pady=(0, 4))
        self._preview = scrolledtext.ScrolledText(
            der, wrap="word", font=("Consolas", 8), bg="#F8FAFF",
            fg=C_TEXT, relief="flat", bd=0)
        self._preview.pack(fill="both", expand=True)
        self._refrescar_preview()

        # ── Barra inferior ──
        barra = tk.Frame(self, bg=C_WHITE); barra.pack(fill="x", padx=10, pady=(0, 10))
        self._lbl_cuenta = tk.Label(barra, text="", bg=C_WHITE, fg=C_GRAY2, font=FONT_NORMAL)
        self._lbl_cuenta.pack(side="left", padx=16, pady=10)
        self._actualizar_cuenta()
        tk.Button(barra, text="✖ Cancelar", command=self.destroy,
                  bg=C_WHITE, fg=C_TEXT, highlightthickness=1, highlightbackground=T.BORDER, font=("Segoe UI", 10),
                  relief="flat", padx=16, pady=8, cursor="hand2").pack(side="right", padx=(8, 16), pady=8)
        self._btn_env = tk.Button(barra, text="📨  Enviar Correo",
                                   bg=C_GREEN, fg=C_WHITE, font=("Segoe UI", 10, "bold"),
                                   relief="flat", padx=20, pady=8, cursor="hand2",
                                   command=self._enviar)
        self._btn_env.pack(side="right", pady=8)

    def _fila(self, arch, extra=False):
        var = tk.BooleanVar(value=True)
        self._check_vars[arch["ruta"]] = var
        bg = "#F0FDF4" if not extra else "#FFF7ED"
        fila = tk.Frame(self._lista, bg=bg); fila.pack(fill="x", pady=1, padx=2)
        tk.Checkbutton(fila, variable=var, bg=bg, cursor="hand2",
                       command=lambda: (self._refrescar_preview(), self._actualizar_cuenta())
                       ).pack(side="left", padx=(6, 0), pady=4)
        ext = os.path.splitext(arch["nombre"])[1].lower()
        ico = {".pdf":"📄",".xml":"📋",".xlsx":"📊",".xls":"📊",
               ".zip":"🗜",".rar":"🗜",".sua":"💾",".txt":"📝"}.get(ext, "📎")
        inner = tk.Frame(fila, bg=bg); inner.pack(side="left", fill="x", expand=True, padx=6, pady=4)
        tk.Label(inner, text=f"{ico}  {arch['nombre']}",
                 bg=bg, fg=C_TEXT, font=FONT_NORMAL, anchor="w").pack(anchor="w")
        tk.Label(inner, text=CE.formato_tamano(arch["tamano"]),
                 bg=bg, fg=C_GRAY2, font=FONT_SMALL).pack(anchor="w")

    def _sel_all(self, v):
        for var in self._check_vars.values(): var.set(v)
        self._refrescar_preview(); self._actualizar_cuenta()

    def _agregar(self):
        rutas = filedialog.askopenfilenames(parent=self, title="Agregar archivos",
                    filetypes=[("Todos","*.*"),("PDF","*.pdf"),("XML","*.xml"),
                               ("Excel","*.xlsx *.xls"),("ZIP","*.zip")])
        for r in rutas:
            if r not in self._check_vars:
                self._fila({"nombre": os.path.basename(r),
                             "ruta": r, "tamano": os.path.getsize(r)}, extra=True)
        self._refrescar_preview(); self._actualizar_cuenta()

    def _seleccionados(self): return [r for r, v in self._check_vars.items() if v.get()]

    def _actualizar_cuenta(self):
        sel = self._seleccionados()
        tam = sum(os.path.getsize(r) for r in sel if os.path.exists(r))
        self._lbl_cuenta.config(
            text=f"📎 {len(sel)} archivo(s)  |  {CE.formato_tamano(tam)} total")

    def _refrescar_preview(self):
        sel = self._seleccionados()
        txt = CE.generar_cuerpo(self._empresa, self._rfc, self._cfg,
                                self._mes, self._anio, sel if sel else None)
        self._preview.config(state="normal")
        self._preview.delete("1.0", "end")
        self._preview.insert("1.0", txt)
        self._preview.config(state="disabled")

    def _enviar(self):
        if not self._email_cfg.get("remitente") or not self._email_cfg.get("password"):
            messagebox.showerror("Email no configurado",
                "Configura el correo en  ⚙ Configuración → 📧 Email.", parent=self); return
        dest_raw = self._cfg.get("destinatario", "").strip()
        if not dest_raw:
            messagebox.showerror("Sin destinatario",
                "Esta empresa no tiene correo de destinatario configurado.\n"
                "Ve a la pestaña ⚖ Obligaciones → Configuración empresa.", parent=self); return

        sel    = self._seleccionados()
        asunto = CE.generar_asunto(self._cfg, self._empresa, self._mes, self._anio)
        cuerpo = CE.generar_cuerpo(self._empresa, self._rfc, self._cfg,
                                   self._mes, self._anio, sel if sel else None)
        dests  = [d.strip() for d in dest_raw.split(",") if d.strip()]
        cc_raw = self._cfg.get("cc", "")
        ccs    = [c.strip() for c in cc_raw.split(",") if c.strip()]
        prov   = PROVIDERS.get(self._email_cfg.get("proveedor", "smtp"), {})

        if not messagebox.askyesno("Confirmar envío",
            f"¿Enviar correo?\n\nPara: {dest_raw}\nCC: {cc_raw or '(ninguno)'}\n"
            f"Asunto: {asunto}\nAdjuntos: {len(sel)}\nProveedor: {prov.get('nombre','SMTP')}",
            parent=self): return

        self._btn_env.config(text="Enviando…", state="disabled", bg=C_GRAY2)
        self.update()

        def tarea():
            try:
                enviar_correo(self._email_cfg, asunto, cuerpo, dests, cc=ccs, adjuntos=sel)
                self.after(0, self._ok, len(sel), dest_raw)
            except Exception as e:
                self.after(0, self._error, str(e))

        threading.Thread(target=tarea, daemon=True).start()

    def _ok(self, n, dest):
        self._log(f"✅ Compliance enviado: {self._empresa} – {CE.MESES_ES[self._mes]} {self._anio} ({n} adj.)")
        messagebox.showinfo("Enviado", f"✅ Correo enviado a {dest}\nAdjuntos: {n}", parent=self)
        self.destroy()

    def _error(self, err):
        self._btn_env.config(text="📨  Enviar Correo", state="normal", bg=C_GREEN)
        self._log(f"❌ Error envío compliance: {self._empresa} — {err}")
        messagebox.showerror("Error al enviar",
            f"No se pudo enviar:\n\n{err}\n\nVerifica la configuración de email.", parent=self)


# ═══════════════════════════════════════════════════════════════════════════════
#  DIÁLOGOS INTERNOS CONVERTIDOR
# ═══════════════════════════════════════════════════════════════════════════════
class _DialogCambiarPwd(tk.Toplevel):
    def __init__(self, parent, usuario):
        super().__init__(parent); self.title("Cambiar contraseña")
        self.resizable(False, False); self.configure(bg=C_WHITE); self.grab_set()
        tk.Label(self, text="Cambiar contraseña", bg=C_WHITE, fg=C_TEXT,
                 font=FONT_TITLE).pack(pady=(20, 8))
        body = tk.Frame(self, bg=C_WHITE, padx=30); body.pack(fill="x")
        def campo(lb, show=""):
            tk.Label(body, text=lb, bg=C_WHITE, fg=C_TEXT, font=FONT_HEAD, anchor="w").pack(fill="x")
            v = tk.StringVar()
            f = tk.Frame(body, bg=C_GRAY); f.pack(fill="x", pady=(4, 10))
            tk.Entry(f, textvariable=v, show=show, font=("Segoe UI", 10),
                     bg=C_GRAY, fg=C_TEXT, relief="flat", bd=5,
                     insertbackground=C_ACCENT).pack(fill="x")
            return v
        self._va = campo("Contraseña actual:", show="*")
        self._vn = campo("Nueva contraseña:", show="*")
        self._vc = campo("Confirmar nueva:", show="*")
        self._vm = tk.Label(body, text="", bg=C_WHITE, fg=C_RED, font=FONT_SMALL, wraplength=320)
        self._vm.pack()
        self._usuario = usuario
        bf = tk.Frame(body, bg=C_WHITE); bf.pack(fill="x", pady=(8, 20))
        tk.Button(bf, text="Cancelar", command=self.destroy,
                  bg=C_GRAY, fg=C_GRAY2, font=FONT_NORMAL,
                  relief="flat", padx=12, pady=6, cursor="hand2").pack(side="right", padx=(8, 0))
        tk.Button(bf, text="Cambiar", command=self._ok,
                  bg=C_ACCENT, fg=C_TEXT, font=FONT_HEAD,
                  relief="flat", padx=16, pady=6, cursor="hand2").pack(side="right")

    def _ok(self):
        if self._vn.get() != self._vc.get():
            self._vm.config(text="Las contraseñas no coinciden."); return
        ok, msg = cambiar_password(self._usuario, self._va.get(), self._vn.get())
        if ok: messagebox.showinfo("OK", "✅ Contraseña actualizada.", parent=self); self.destroy()
        else: self._vm.config(text=msg)


# ═══════════════════════════════════════════════════════════════════════════════
#  DATACLASS PDF ITEM
# ═══════════════════════════════════════════════════════════════════════════════
@dataclass
class ArchivoItem:
    path:        str
    nombre:      str
    banco:       str = "Detectando…"
    estado:      str = "pendiente"
    movimientos: int = 0
    mensaje:     str = ""


# ═══════════════════════════════════════════════════════════════════════════════
#  APP PRINCIPAL — COMBINADA
# ═══════════════════════════════════════════════════════════════════════════════
class App(tk.Tk):
    def __init__(self, usuario: str = ""):
        super().__init__()
        # Aplicar tema (detecta light/dark de Windows y configura ttk.Style)
        ui_theme.apply_theme(self)
        self.title("ConvertidorBancos v5.0 — Compliance + Convertidor")
        self.update_idletasks()
        # Detecta el área útil de la PANTALLA ACTUAL (DPI-aware)
        mx, my, sw, sh = ui_theme.work_area(self)
        w  = int(sw * 0.94)
        h  = int(sh * 0.90)
        x  = mx + (sw - w) // 2
        y  = my + max(0, (sh - h) // 2 - 10)
        self.geometry(f"{w}x{h}+{x}+{y}")   # tamaño "restaurado" razonable
        self.minsize(min(ui_theme.sx(920), int(sw * 0.72)), min(ui_theme.sx(560), int(sh * 0.72)))
        self.configure(bg=T.BG)
        # Ajuste automático a la pantalla: iniciar MAXIMIZADO (respeta la barra
        # de tareas y el DPI del monitor). Si falla, queda la geometría de arriba.
        try:
            self.state("zoomed")
        except Exception:
            pass

        self._usuario    = usuario
        self._archivos:  list[ArchivoItem] = []
        self._output_dir: str = ""
        self._running    = False
        self._stop_detect = threading.Event()
        self._email_cfg  = _cargar_email_cfg()

        self._build_ui()

    # ── Layout ───────────────────────────────────────────────────────────────
    def _build_ui(self):
        # Layout nuevo: SIDEBAR izquierdo + zona derecha (header arriba + contenido).
        # Reemplaza el Notebook de pestañas. El CONTENIDO de cada módulo es el mismo
        # de siempre (sus botones ya están cableados a sus funciones); solo cambia
        # la navegación: ahora se elige el módulo desde el menú lateral.
        root = tk.Frame(self, bg=T.BG); root.pack(fill="both", expand=True)
        root.grid_columnconfigure(1, weight=1); root.grid_rowconfigure(0, weight=1)

        self._build_sidebar(root)

        right = tk.Frame(root, bg=T.BG); right.grid(row=0, column=1, sticky="nsew")
        right.grid_columnconfigure(0, weight=1); right.grid_rowconfigure(1, weight=1)
        self._build_header(right)

        self._content = tk.Frame(right, bg=T.BG); self._content.grid(row=1, column=0, sticky="nsew")
        self._content.grid_columnconfigure(0, weight=1); self._content.grid_rowconfigure(0, weight=1)

        self._tab_conv       = tk.Frame(self._content, bg=T.BG)
        self._tab_concil     = tk.Frame(self._content, bg=T.BG)
        self._tab_compliance = tk.Frame(self._content, bg=T.BG)
        self._tab_saldos     = tk.Frame(self._content, bg=T.BG)
        self._tab_config     = tk.Frame(self._content, bg=T.BG)
        for _fr in (self._tab_conv, self._tab_concil, self._tab_compliance,
                    self._tab_saldos, self._tab_config):
            _fr.grid(row=0, column=0, sticky="nsew")

        # estado compartido para la conciliación
        self._concil_pdfs:     list[dict] = []   # {path, nombre, banco, movs, info}
        self._concil_carpetas: list[str]  = []

        self._build_convertidor_tab()
        self._build_conciliacion_tab()
        self._build_compliance_tab()
        self._build_config_tab()

        self._mostrar_modulo("concil")

    # ── Sidebar (navegación entre módulos) ────────────────────────────────────
    def _build_sidebar(self, root):
        SB, SB2, INK, STRONG, ACC, ACTV = ("#101826", "#0B1220", "#C7D2E0",
                                           "#FFFFFF", "#4E9E79", "#1E2E3E")
        self._side_cols = (SB, INK, ACTV, STRONG)
        side = tk.Frame(root, bg=SB, width=234); side.grid(row=0, column=0, sticky="ns")
        side.grid_propagate(False)
        br = tk.Frame(side, bg=SB); br.pack(fill="x", padx=14, pady=(16, 14))
        tk.Label(br, text="$", width=3, font=("Segoe UI", 14, "bold"), bg=ACC, fg=STRONG).pack(side="left")
        tk.Label(br, text="ConvertidorBancos", font=("Segoe UI", 12, "bold"), bg=SB, fg=STRONG).pack(side="left", padx=8)
        tk.Label(side, text="MÓDULOS", font=("Segoe UI", 8, "bold"), bg=SB, fg=INK, anchor="w").pack(fill="x", padx=16, pady=(6, 4))
        self._side_btns = {}
        for key, txt in (("conv", "🏦  Convertidor PDF → Excel"),
                         ("concil", "📊  Conciliación CFDI"),
                         ("compliance", "⚖  Obligaciones Fiscales"),
                         ("config", "⚙  Configuración")):
            b = tk.Button(side, text=txt, anchor="w", relief="flat", cursor="hand2",
                          font=("Segoe UI", 11), bd=0, padx=12, pady=9, bg=SB, fg=INK,
                          activebackground=SB2, activeforeground=STRONG,
                          command=lambda k=key: self._mostrar_modulo(k))
            b.pack(fill="x", padx=10, pady=1); self._side_btns[key] = b

    def _mostrar_modulo(self, key):
        frames = {"conv": self._tab_conv, "concil": self._tab_concil,
                  "compliance": self._tab_compliance, "saldos": self._tab_saldos,
                  "config": self._tab_config}
        frames.get(key, self._tab_concil).tkraise()
        SB, INK, ACTV, STRONG = self._side_cols
        for k, b in self._side_btns.items():
            act = (k == key)
            b.config(bg=(ACTV if act else SB), fg=(STRONG if act else INK))
        # Al cambiar de módulo, refresca los combos de RFC desde el catálogo central
        # (empresas FDB) para que TODOS carguen las empresas registradas en Config.
        self._refrescar_rfcs()

    def _card_badge(self, parent, num, titulo, grid=None, pack=None):
        """Tarjeta estilo rediseño: badge numerado + título; devuelve el FRAME
        de contenido. Redondeada si hay CustomTkinter."""
        if _HAS_CTK:
            card = ctk.CTkFrame(parent, corner_radius=14, fg_color=C_WHITE,
                                border_width=1, border_color=T.BORDER)
        else:
            card = tk.Frame(parent, bg=C_WHITE, highlightbackground=T.BORDER,
                            highlightcolor=T.BORDER, highlightthickness=1)
        if grid is not None:
            card.grid(**grid)
        else:
            card.pack(**(pack or {"fill": "x", "padx": 12, "pady": (0, 6)}))
        card.rowconfigure(1, weight=1); card.columnconfigure(0, weight=1)
        head = tk.Frame(card, bg=C_WHITE)
        head.grid(row=0, column=0, sticky="ew", padx=14, pady=(12, 2))
        if _HAS_CTK:
            ctk.CTkLabel(head, text=str(num), fg_color=T.ACCENT, text_color=T.ACCENT_FG,
                         corner_radius=8, width=28, height=28,
                         font=("Segoe UI", 12, "bold")).pack(side="left")
        else:
            tk.Label(head, text=str(num), bg=T.ACCENT, fg=T.ACCENT_FG,
                     font=("Segoe UI", 11, "bold"), width=2).pack(side="left")
        tk.Label(head, text=titulo, bg=C_WHITE, fg=T.TEXT,
                 font=("Segoe UI", 11, "bold")).pack(side="left", padx=9)
        body = tk.Frame(card, bg=C_WHITE)
        body.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 10))
        return body

    # ══════════════════════════════════════════════════════════════════════════
    #  TAB: SALDOS A FAVOR ISR / IVA  (motor sat_saldos)
    # ══════════════════════════════════════════════════════════════════════════
    def _build_saldos_tab(self, parent=None):
        f = parent if parent is not None else self._tab_saldos
        self._module_header(
            f, "💰", "Contribuciones, Pagos y Saldos a Favor",
            "Lee acuses, declaraciones detalladas y acuses de pago del SAT y arma el "
            "papel de trabajo del remanente acumulado (2015–2026).")

        if not _SALDOS_OK:
            tk.Label(f, text="Módulo no disponible: no se pudo cargar el motor «sat_saldos».",
                     bg=T.BG, fg=T.DANGER, font=F.BODY).pack(padx=16, pady=16, anchor="w")
            return

        self._saldos_rutas: list[Path] = []
        self._saldos_cargados = None
        self._saldos_res = None

        # ── Barra de acción ──────────────────────────────────────────────────
        # La CARGA se hace en la pestaña «Archivos» (punto único); aquí solo se
        # analiza y se genera el Excel del remanente.
        action = tk.Frame(f, bg=T.BG, pady=8); action.pack(fill="x", padx=12)
        tk.Label(action, text="Carga los PDF en la pestaña «Archivos». Aquí analiza "
                 "y genera el papel de trabajo.",
                 bg=T.BG, fg=T.TEXT_MUTED, font=F.SMALL).pack(side="left", padx=(0, 6))
        self._rbtn(action, "Limpiar", self._saldos_limpiar, "ghost").pack(side="left")
        self._saldos_btn_excel = self._rbtn(action, "⬇  Generar Excel",
                                            self._saldos_generar_excel, "accent")
        self._saldos_btn_excel.pack(side="right")
        self._rbtn(action, "📊  Generar reporte", self._reporte_saldos,
                   "primary").pack(side="right", padx=(0, 6))
        self._saldos_btn_analizar = self._rbtn(action, "▶  ANALIZAR",
                                               self._saldos_analizar, "accent")
        self._saldos_btn_analizar.pack(side="right", padx=(0, 6))

        # ── Opciones (saldo inicial IVA + RFC) ───────────────────────────────
        opt = tk.Frame(f, bg=T.BG); opt.pack(fill="x", padx=12, pady=(0, 2))
        tk.Label(opt, text="Saldo a favor de IVA inicial ($):", bg=T.BG,
                 fg=T.TEXT_MUTED, font=F.SMALL).pack(side="left")
        self._saldos_var_ini = tk.StringVar(value="0")
        tk.Entry(opt, textvariable=self._saldos_var_ini, width=12).pack(side="left", padx=(4, 16))
        tk.Label(opt, text="RFC:", bg=T.BG, fg=T.TEXT_MUTED, font=F.SMALL).pack(side="left")
        self._saldos_var_rfc = tk.StringVar()
        self._saldos_rfc_combo = ttk.Combobox(opt, textvariable=self._saldos_var_rfc,
                                               width=18, state="readonly")
        self._saldos_rfc_combo.pack(side="left", padx=(4, 16))
        self._saldos_rfc_combo.bind("<<ComboboxSelected>>", lambda e: self._saldos_reanalizar())
        # Periodo (para el reporte de esta pestaña; los años salen de Config Empresa)
        tk.Label(opt, text="Mes:", bg=T.BG, fg=T.TEXT_MUTED, font=F.SMALL).pack(side="left")
        self._sa_mes = ttk.Combobox(opt, values=CE.MESES_ES[1:], width=11, state="readonly")
        self._sa_mes.set(CE.MESES_ES[date.today().month]); self._sa_mes.pack(side="left", padx=(4, 8))
        tk.Label(opt, text="Año:", bg=T.BG, fg=T.TEXT_MUTED, font=F.SMALL).pack(side="left")
        self._sa_anio = ttk.Combobox(opt, values=[str(y) for y in CE.anios_disponibles()],
                                     width=7, state="readonly")
        self._sa_anio.set(self._anio_default()); self._sa_anio.pack(side="left", padx=(4, 16))
        self._saldos_estado = tk.Label(opt, text="", bg=T.BG, fg=T.ACCENT, font=F.SMALL)
        self._saldos_estado.pack(side="left")

        # ── Lista de rutas seleccionadas ─────────────────────────────────────
        listf = tk.Frame(f, bg=T.BG); listf.pack(fill="x", padx=12, pady=(2, 0))
        self._saldos_lista = tk.Listbox(listf, height=3,
                                        bg=T.SURFACE, fg=T.TEXT, relief="flat",
                                        highlightthickness=1, highlightbackground=T.BORDER)
        self._saldos_lista.pack(fill="x")

        # ── Resumen ──────────────────────────────────────────────────────────
        self._saldos_resumen = tk.Label(f, text="", bg=T.BG, fg=T.TEXT,
                                        font=F.BODY, justify="left", anchor="w")
        self._saldos_resumen.pack(fill="x", padx=12, pady=(6, 2))

        # ── Pestañas: IVA acumulado / Pagos ──────────────────────────────────
        body = tk.Frame(f, bg=T.BG); body.pack(fill="both", expand=True, padx=12, pady=(2, 4))
        nb = ttk.Notebook(body); nb.pack(fill="both", expand=True); self._saldos_nb = nb

        tab_iva = tk.Frame(nb, bg=T.BG); nb.add(tab_iva, text="IVA acumulado")
        cols = ("periodo", "operacion", "fecha", "cargo", "generado", "acreditado", "remanente")
        self._saldos_tv_iva = ttk.Treeview(tab_iva, columns=cols, show="headings", height=11)
        for c, txt, w in zip(cols,
                             ("Periodo", "No. operación", "Fecha present.", "IVA a cargo",
                              "Saldo a favor gen.", "Acreditado (ant.)", "Remanente acumulado"),
                             (90, 120, 130, 110, 140, 130, 150)):
            self._saldos_tv_iva.heading(c, text=txt)
            self._saldos_tv_iva.column(
                c, width=w, anchor="center" if c in ("periodo", "operacion", "fecha") else "e")
        sbi = ttk.Scrollbar(tab_iva, orient="vertical", command=self._saldos_tv_iva.yview)
        self._saldos_tv_iva.configure(yscrollcommand=sbi.set)
        self._saldos_tv_iva.pack(side="left", fill="both", expand=True); sbi.pack(side="right", fill="y")

        tab_pag = tk.Frame(nb, bg=T.BG); nb.add(tab_pag, text="Pagos")
        colp = ("periodo", "tipo", "operacion", "venc", "contrib", "actualiz",
                "recargos", "a_pagar", "pagado", "fecha_pago", "estatus")
        self._saldos_tv_pag = ttk.Treeview(tab_pag, columns=colp, show="headings", height=11)
        for c, txt, w in zip(
                colp,
                ("Periodo", "Tipo", "No. operación", "Vencim.", "Contribución",
                 "Actualización", "Recargos", "Importe a pagar", "Importe pagado",
                 "Fecha pago", "Estatus"),
                (75, 55, 110, 85, 100, 100, 85, 110, 110, 90, 120)):
            self._saldos_tv_pag.heading(c, text=txt)
            self._saldos_tv_pag.column(
                c, width=w, anchor="center" if c in ("periodo", "tipo", "operacion", "venc",
                                                     "fecha_pago", "estatus") else "e")
        self._saldos_tv_pag.tag_configure("pagado", background=T.SUCCESS_BG, foreground=T.SUCCESS)
        self._saldos_tv_pag.tag_configure("nopagado", background=T.DANGER_BG, foreground=T.DANGER)
        sbp = ttk.Scrollbar(tab_pag, orient="vertical", command=self._saldos_tv_pag.yview)
        self._saldos_tv_pag.configure(yscrollcommand=sbp.set)
        self._saldos_tv_pag.pack(side="left", fill="both", expand=True); sbp.pack(side="right", fill="y")

        # ── Nota fiscal ──────────────────────────────────────────────────────
        self._saldos_nota = tk.Label(f, text="", bg=T.BG, fg=T.DANGER, font=F.SMALL,
                                     justify="left", anchor="w", wraplength=1100)
        self._saldos_nota.pack(fill="x", padx=12, pady=(0, 6))

    # ── Saldos: selección de archivos ────────────────────────────────────────
    def _saldos_refrescar_lista(self):
        self._saldos_lista.delete(0, "end")
        for r in self._saldos_rutas:
            self._saldos_lista.insert("end", str(r))

    def _saldos_sumar(self, nuevos):
        for n in nuevos:
            if n not in self._saldos_rutas:
                self._saldos_rutas.append(n)
        self._saldos_refrescar_lista()

    def _saldos_limpiar(self):
        self._saldos_rutas.clear()
        self._saldos_refrescar_lista()

    def _saldos_saldo_inicial(self):
        try:
            v = float(str(self._saldos_var_ini.get()).replace(",", "").replace("$", "").strip() or 0)
            return v or None
        except ValueError:
            return None

    # ── Saldos: análisis (lectura en hilo) ───────────────────────────────────
    def _saldos_analizar(self):
        if not self._saldos_rutas:
            messagebox.showwarning("Sin archivos", "Agrega al menos una carpeta, PDF o ZIP.")
            return
        self._saldos_btn_analizar.configure(state="disabled")
        self._saldos_estado.configure(text="Leyendo PDFs…", fg=T.ACCENT)
        threading.Thread(target=self._saldos_leer_worker, daemon=True).start()

    def _saldos_leer_worker(self):
        try:
            pdfs = _SF_PIPE.expandir_entradas(self._saldos_rutas)

            def _prog(i, n, nom):
                try:
                    self.after(0, lambda: self._saldos_estado.configure(
                        text=f"Leyendo {nom} ({i + 1}/{n})"))
                except Exception:
                    pass

            cargados = _SF_PIPE.cargar(pdfs, progreso=_prog)
            self.after(0, self._saldos_leido, cargados)
        except Exception as e:  # noqa
            import traceback
            self.after(0, self._saldos_error, f"{e}\n\n{traceback.format_exc()}")

    def _saldos_error(self, msg):
        self._saldos_btn_analizar.configure(state="normal")
        self._saldos_estado.configure(text="Error", fg=T.DANGER)
        messagebox.showerror("Saldos a favor — error", msg)

    def _saldos_leido(self, cargados):
        self._saldos_cargados = cargados
        self._saldos_btn_analizar.configure(state="normal")
        if not cargados.declaraciones and not cargados.detalles and not cargados.pagos:
            self._saldos_error("No se leyó ningún acuse/declaración/pago legible "
                               "(¿PDF escaneados sin texto o vacíos?).")
            return
        rfcs = cargados.rfcs()
        self._saldos_rfc_combo["values"] = rfcs
        # Conservar el RFC ya elegido si sigue presente (antes se pisaba con rfcs[0]).
        if rfcs and self._saldos_var_rfc.get() not in rfcs:
            self._saldos_var_rfc.set(rfcs[0])
        self._saldos_estado.configure(
            text=f"{cargados.n_pdfs} PDF · {len(cargados.declaraciones)} acuses · "
                 f"{len(cargados.detalles)} detalles IVA · {len(cargados.pagos)} pagos · "
                 f"{len(rfcs)} RFC", fg=T.ACCENT)
        self._saldos_reanalizar()

    def _saldos_reanalizar(self):
        if not self._saldos_cargados:
            return
        try:
            # NUNCA analizar con rfc=None: si hay varios RFC cargados, el motor
            # mezcla documentos de empresas distintas (RFC de una con la razón
            # social de otra). Si no hay selección, se fuerza el primer RFC.
            rfc = (self._saldos_var_rfc.get() or "").strip() or None
            if rfc is None:
                _rfcs = self._saldos_cargados.rfcs()
                if _rfcs:
                    rfc = _rfcs[0]
                    self._saldos_var_rfc.set(rfc)
            self._saldos_res = _SF_PIPE.analizar(
                self._saldos_cargados, rfc=rfc, saldo_inicial_iva=self._saldos_saldo_inicial())
            self._saldos_pintar()
            self._saldos_btn_excel.configure(state="normal")
            self._saldos_persistir()   # deja lo procesado en la FDB
        except Exception as e:  # noqa
            import traceback
            self._saldos_error(f"No se pudo analizar:\n{e}\n\n{traceback.format_exc()}")

    def _saldos_pintar(self):
        res = self._saldos_res
        acum = res.iva_acumulado
        lineas = [f"{res.rfc} — {res.razon_social}", res.regimen,
                  f"Periodos vigentes: {res.n_vigentes}   ·   Remanente ISR: "
                  f"${res.remanente_isr_total:,.0f}"]
        base = (self._saldos_saldo_inicial() or 0) or (acum.saldo_inicial_requerido if acum else 0)
        for i in self._saldos_tv_iva.get_children():
            self._saldos_tv_iva.delete(i)
        if acum and acum.renglones:
            running = base
            for r in acum.renglones:
                running += r.a_favor_generado - r.acreditamiento
                fp = getattr(r, "fecha_presentacion", None)
                fecha = fp.strftime("%d/%m/%Y %H:%M") if fp else "—"
                op = getattr(r, "operacion", None) or "—"
                self._saldos_tv_iva.insert("", "end", values=(
                    _sf_etiqueta_mes(r.ejercicio, r.periodo), op, fecha,
                    f"${r.a_cargo_periodo:,.0f}", f"${r.a_favor_generado:,.0f}",
                    f"${r.acreditamiento:,.0f}", f"${running:,.0f}"))
            cierre = base + acum.neto
            lineas.append(f"IVA — Generado: ${acum.total_generado:,.0f}   Acreditado: "
                          f"${acum.total_acreditado:,.0f}   Neto: ${acum.neto:,.0f}   "
                          f"→ Remanente al cierre: ${cierre:,.0f}")
            self._saldos_nota.configure(text=acum.nota, fg=T.DANGER)
        else:
            self._saldos_nota.configure(text=res.iva.texto_alcance, fg=T.DANGER)

        # Pagos
        for i in self._saldos_tv_pag.get_children():
            self._saldos_tv_pag.delete(i)
        if res.pagos:
            pagados = sum(1 for p in res.pagos if p.pagado)
            rec = sum(p.recargos + p.actualizacion for p in res.pagos)
            lineas.append(f"Pagos: {pagados}/{len(res.pagos)} verificadas · "
                          f"Total pagado: ${res.total_pagado:,.0f} · "
                          f"Recargos+actualización: ${rec:,.0f}")
            for p in res.pagos:
                per = _sf_etiqueta_mes(p.ejercicio, p.periodo) if p.periodo else str(p.ejercicio)
                estatus = "✓ Pagado" if p.pagado else "⚠ Sin comprobante"
                self._saldos_tv_pag.insert(
                    "", "end", tags=("pagado" if p.pagado else "nopagado",),
                    values=(per, p.tipo[:4], p.operacion or "—", p.vencimiento or "—",
                            f"${p.contribucion:,.0f}", f"${p.actualizacion:,.0f}",
                            f"${p.recargos:,.0f}", f"${p.importe_a_pagar:,.0f}",
                            f"${(p.importe_pagado or 0):,.0f}", p.fecha_pago or "—", estatus))
            self._saldos_nb.tab(1, text=f"Pagos ({pagados}/{len(res.pagos)})")
        else:
            self._saldos_nb.tab(1, text="Pagos")
        self._saldos_resumen.configure(text="\n".join(lineas))

    def _saldos_generar_excel(self):
        if not self._saldos_res:
            messagebox.showinfo("Saldos a favor", "Primero pulsa ANALIZAR.")
            return
        ruta = filedialog.asksaveasfilename(
            title="Guardar papel de trabajo", defaultextension=".xlsx",
            initialfile=f"Analisis_{self._saldos_res.rfc or 'SAT'}.xlsx",
            filetypes=[("Excel", "*.xlsx")])
        if not ruta:
            return
        try:
            _SF_PIPE.generar_xlsx(self._saldos_res, Path(ruta),
                                  saldo_inicial_iva=self._saldos_saldo_inicial())
            self._saldos_estado.configure(text=f"✓ Excel generado", fg=T.ACCENT)
            messagebox.showinfo("Listo", f"Papel de trabajo generado:\n{ruta}")
        except Exception as e:  # noqa
            self._saldos_error(str(e))

    def _rbtn(self, parent, text, command, kind="primary", small=False):
        """Botón redondeado (CTkButton) con respaldo tk.
        kind: primary(azul) · accent(verde) · danger(rojo) · warn(naranja) · ghost."""
        pal = {"primary": (T.BLUE, "#FFFFFF"),
               "accent":  (T.ACCENT, T.ACCENT_FG),
               "danger":  (T.DANGER, "#FFFFFF"),
               "warn":    (getattr(T, "WARN", "#B7791F"), "#FFFFFF"),
               "ghost":   (T.SURFACE, T.TEXT)}
        bg, fg = pal.get(kind, pal["primary"])
        if _HAS_CTK:
            kw = dict(corner_radius=8, fg_color=bg, text_color=fg,
                      font=("Segoe UI", 10, "bold"), height=30 if small else 34)
            if kind == "ghost":
                kw.update(border_width=1, border_color=T.BORDER,
                          hover_color=T.ACCENT_SOFT)
            return ctk.CTkButton(parent, text=text, command=command, **kw)
        return tk.Button(parent, text=text, command=command, bg=bg, fg=fg,
                         relief="flat", padx=10, pady=4, cursor="hand2",
                         font=("Segoe UI", 9, "bold"))

    def _module_header(self, parent, icon, titulo, subtitulo):
        """Cabecera de módulo estilo prototipo: barra blanca con chip de icono
        (redondo si hay CTk) + título + subtítulo."""
        hdr = tk.Frame(parent, bg=T.SURFACE, highlightbackground=T.BORDER,
                       highlightthickness=1)
        hdr.pack(fill="x")
        if _HAS_CTK:
            ctk.CTkLabel(hdr, text=icon, fg_color=T.ACCENT_SOFT, text_color=T.ACCENT,
                         corner_radius=10, width=38, height=38,
                         font=("Segoe UI", 14, "bold")).pack(side="left", padx=(16, 12), pady=10)
        else:
            tk.Label(hdr, text=icon, bg=T.ACCENT_SOFT, fg=T.ACCENT, width=3,
                     font=("Segoe UI", 14, "bold")).pack(side="left", padx=(16, 12), pady=10)
        _tt = tk.Frame(hdr, bg=T.SURFACE); _tt.pack(side="left", pady=8)
        tk.Label(_tt, text=titulo, bg=T.SURFACE, fg=T.TEXT,
                 font=("Segoe UI", 13, "bold")).pack(anchor="w")
        tk.Label(_tt, text=subtitulo, bg=T.SURFACE, fg=T.TEXT_MUTED,
                 font=F.SMALL).pack(anchor="w")
        return hdr

    def _build_header(self, parent):
        # Barra superior estilo prototipo: fondo BLANCO, chip de icono + título,
        # pills (usuario / licencia) a la derecha.
        h = tk.Frame(parent, bg=T.SURFACE, height=58,
                     highlightbackground=T.BORDER, highlightthickness=1)
        h.grid(row=0, column=0, sticky="ew")
        h.grid_propagate(False)
        if _HAS_CTK:
            ctk.CTkLabel(h, text="⇄", fg_color=T.ACCENT_SOFT, text_color=T.ACCENT,
                         corner_radius=10, width=40, height=40,
                         font=("Segoe UI", 15, "bold")).pack(side="left", padx=(16, 12), pady=8)
        else:
            tk.Label(h, text="⇄", bg=T.ACCENT_SOFT, fg=T.ACCENT,
                     font=("Segoe UI", 15, "bold"), width=3).pack(side="left", padx=(16, 12), pady=9)
        tt = tk.Frame(h, bg=T.SURFACE); tt.pack(side="left", pady=9)
        tk.Label(tt, text="ConvertidorBancos", bg=T.SURFACE, fg=T.TEXT,
                 font=("Segoe UI", 13, "bold")).pack(anchor="w")
        tk.Label(tt, text="Compliance · Convertidor · Email  ·  v5.0",
                 bg=T.SURFACE, fg=T.TEXT_MUTED, font=F.SMALL).pack(anchor="w")
        if self._usuario:
            tk.Button(h, text=f"👤 {self._usuario}  ·  Cambiar contraseña",
                      command=lambda: _DialogCambiarPwd(self, self._usuario),
                      bg=T.SURFACE, fg=T.TEXT_MUTED, font=F.SMALL, relief="flat",
                      cursor="hand2", padx=8, pady=4, bd=0,
                      activebackground=T.SURFACE, activeforeground=T.ACCENT,
                      ).pack(side="right", padx=(0, 12), pady=13)

    # ══════════════════════════════════════════════════════════════════════════
    #  TAB: CONVERTIDOR PDF→EXCEL
    # ══════════════════════════════════════════════════════════════════════════
    def _build_convertidor_tab(self):
        f = self._tab_conv

        # ── Cabecera de módulo (blanca, estilo prototipo) ────────────────────
        self._module_header(
            f, "🏦", "Convertidor de Estados de Cuenta",
            "Convierte los PDF del banco a Excel con saldos y clasificación de movimientos")

        # ── Barra de acción ──────────────────────────────────────────────────
        action = tk.Frame(f, bg=T.BG, pady=8)
        action.pack(fill="x", padx=12)

        self._rbtn(action, "📂  Agregar PDFs", self._agregar_pdf, "primary"
                   ).pack(side="left", padx=(0, 6))
        self._rbtn(action, "Limpiar lista", self._limpiar_lista, "ghost"
                   ).pack(side="left")
        self._btn_convert = self._rbtn(action, "▶  CONVERTIR A EXCEL",
                                       self._start_conversion, "accent")
        self._btn_convert.pack(side="right")
        tk.Label(action, text=f"👤 {self._usuario}",
                 bg=T.BG, fg=T.TEXT_MUTED, font=F.SMALL).pack(side="right", padx=12)

        # ── Cuerpo ───────────────────────────────────────────────────────────
        body = tk.Frame(f, bg=T.BG)
        body.pack(fill="both", expand=True, padx=12, pady=(8, 4))

        # Tabla de archivos
        table_f = tk.Frame(body, bg=T.BG)
        table_f.pack(side="left", fill="both", expand=True)
        cols = ("Archivo", "Banco", "Estado", "Movs")
        self._tree = ttk.Treeview(table_f, columns=cols, show="headings", height=9, selectmode="extended")
        for c, w in zip(cols, [320, 130, 110, 70]):
            self._tree.heading(c, text=c, anchor="w")
            self._tree.column(c, width=w, minwidth=60, anchor="w")
        # Tags de estado con colores semánticos del tema
        self._tree.tag_configure("ok",         background=T.SUCCESS_BG, foreground=T.SUCCESS)
        self._tree.tag_configure("error",       background=T.DANGER_BG,  foreground=T.DANGER)
        self._tree.tag_configure("procesando",  background=T.WARN_BG,    foreground=T.WARN)
        self._tree.tag_configure("pendiente",   background=T.SURFACE)
        sb = ttk.Scrollbar(table_f, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=sb.set)
        self._tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self._tree.bind("<Delete>", lambda _: self._eliminar_seleccionados())

        # ── Panel derecho ─────────────────────────────────────────────────────
        right = tk.Frame(body, bg=T.BG, width=200)
        right.pack(side="right", fill="y", padx=(12, 0))
        right.pack_propagate(False)

        def _btn_r(txt, cmd, bg=None, fg=None):
            _bg = bg if bg else T.SURFACE
            _fg = fg if fg else T.TEXT
            b = tk.Button(right, text=txt, command=cmd, bg=_bg, fg=_fg,
                          font=F.BODY, relief="flat", padx=8, pady=6, cursor="hand2", bd=0,
                          highlightthickness=1, highlightbackground=T.BORDER,
                          activebackground=T.ACCENT_SOFT, activeforeground=T.ACCENT)
            b.pack(fill="x", pady=(0, 4))

        _btn_r("🏢 Empresas / RFC",  lambda: DialogEmpresas(self), bg=T.ACCENT_SOFT, fg=T.ACCENT)
        _btn_r("📋 Bitácora",        self._open_bitacora,          bg=T.ACCENT_SOFT, fg=T.ACCENT)
        _btn_r("📥 Reimportar Excel (actualizar)", self._reimportar_movimientos,
               bg=T.ACCENT_SOFT, fg=T.ACCENT)
        tk.Frame(right, bg=T.BORDER, height=1).pack(fill="x", pady=6)

        tk.Label(right, text="Carpeta de salida:", bg=T.BG, fg=T.TEXT_MUTED,
                 font=F.SMALL, anchor="w").pack(fill="x")
        self._var_outdir = tk.StringVar(value="(misma carpeta del PDF)")
        tk.Entry(right, textvariable=self._var_outdir, font=F.SMALL,
                 bg=T.SURFACE_ALT, fg=T.TEXT_MUTED, relief="flat", bd=0,
                 highlightthickness=1, highlightbackground=T.BORDER,
                 state="readonly").pack(fill="x", pady=(4, 4), ipady=3)
        _btn_r("📂 Elegir carpeta salida", self._elegir_salida)
        _btn_r("↺ Restablecer carpeta",   self._reset_salida)

        tk.Frame(right, bg=T.BORDER, height=1).pack(fill="x", pady=6)
        self._var_open = tk.BooleanVar(value=True)
        tk.Checkbutton(right, text="Abrir Excel al terminar",
                       variable=self._var_open, bg=T.BG, fg=T.TEXT,
                       font=F.SMALL, cursor="hand2",
                       activebackground=T.BG, selectcolor=T.SURFACE).pack(anchor="w")

        # ── Barra de progreso ─────────────────────────────────────────────────
        prog_f = tk.Frame(f, bg=T.BG)
        prog_f.pack(fill="x", padx=12, pady=(0, 4))
        self._progress = ttk.Progressbar(prog_f, length=400, mode="determinate")
        self._progress.pack(side="left", fill="x", expand=True)
        self._lbl_status = tk.Label(prog_f, text="Listo",
                                     bg=T.BG, fg=T.TEXT_MUTED, font=F.SMALL)
        self._lbl_status.pack(side="left", padx=10)

        # ── Log tipo terminal ─────────────────────────────────────────────────
        log_f = tk.Frame(f, bg=T.BG)
        log_f.pack(fill="x", padx=12, pady=(0, 8))
        tk.Label(log_f, text="Registro de conversión:", bg=T.BG, fg=T.TEXT_MUTED,
                 font=F.SMALL).pack(anchor="w", pady=(0, 3))
        self._log_text = scrolledtext.ScrolledText(
            log_f, height=5, wrap="word", font=F.MONO,
            bg="#1A1A19", fg="#C3C2B7", relief="flat", bd=0,
            highlightthickness=1, highlightbackground=T.BORDER)
        self._log_text.pack(fill="x")
        self._log_text.tag_configure("ok",   foreground="#4ADE80")
        self._log_text.tag_configure("err",  foreground="#F87171")
        self._log_text.tag_configure("info", foreground="#60A5FA")
        self._log_text.tag_configure("bank", foreground="#FACC15")

    # ── Convertidor helpers ───────────────────────────────────────────────────
    def _log_write(self, msg: str, tag: str = ""):
        ts = datetime.now().strftime("%H:%M:%S")
        self._log_text.config(state="normal")
        self._log_text.insert("end", f"[{ts}] {msg}\n", tag)
        self._log_text.see("end")
        self._log_text.config(state="disabled")

    def _agregar_pdf(self):
        paths = filedialog.askopenfilenames(
            title="Seleccionar estados de cuenta PDF",
            filetypes=[("PDF", "*.pdf"), ("Todos", "*.*")])
        existing = {a.path for a in self._archivos}
        for p in paths:
            if p not in existing:
                self._archivos.append(ArchivoItem(path=p, nombre=Path(p).name))
        self._refresh_list()
        self._stop_detect.clear()
        threading.Thread(target=self._auto_detect, daemon=True).start()

    def _auto_detect(self):
        for item in self._archivos:
            if self._stop_detect.is_set(): break
            if item.banco in ("Detectando…", ""):
                try:
                    with PDF_LOCK:
                        ext = detectar_o_generico(item.path)
                    item.banco = ext.NOMBRE if ext else "Desconocido"
                except Exception:
                    item.banco = "Error"
                self.after(0, self._refresh_list)

    def _refresh_list(self):
        for r in self._tree.get_children(): self._tree.delete(r)
        for item in self._archivos:
            estado_txt = {"pendiente": "⏳ Pendiente", "procesando": "⚙ Procesando…",
                          "ok": f"✔ OK ({item.movimientos})", "error": f"✘ {item.mensaje[:30]}"
                          }.get(item.estado, item.estado)
            self._tree.insert("", "end",
                values=(item.nombre, item.banco, estado_txt, item.movimientos),
                tags=(item.estado,))

    def _limpiar_lista(self):
        if self._running: return
        self._archivos.clear(); self._refresh_list()

    def _eliminar_seleccionados(self):
        if self._running: return
        sel = self._tree.selection()
        idxs = {self._tree.index(s) for s in sel}
        self._archivos = [a for i, a in enumerate(self._archivos) if i not in idxs]
        self._refresh_list()

    def _elegir_salida(self):
        d = filedialog.askdirectory(title="Carpeta de salida de archivos Excel")
        if d:
            self._output_dir = d
            self._var_outdir.set(d)

    def _reset_salida(self):
        self._output_dir = ""
        self._var_outdir.set("(misma carpeta del PDF)")

    def _open_bitacora(self):
        _DialogBitacora(self)

    def _reimportar_movimientos(self):
        """Reimporta un Excel de movimientos ya modificado y ACTUALIZA en la BD
        SOLO lo que cambió (match por ID en la columna A)."""
        ruta = filedialog.askopenfilename(
            title="Selecciona el Excel de movimientos modificado",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("Todos", "*.*")])
        if not ruta:
            return
        try:
            import reimportar
            st = reimportar.reimportar_movimientos(
                ruta, resolver=self._resolver_conflicto, usuario=self._usuario)
        except Exception as e:
            messagebox.showerror("Error al reimportar", str(e), parent=self)
            return
        _extra = ""
        if st.get("conflictos"):
            _extra = (f"\nConflictos detectados: {st.get('conflictos', 0)}"
                      f"  (descartados: {st.get('descartados', 0)})"
                      "\nCada conflicto quedó en la bitácora de auditoría.")
        messagebox.showinfo(
            "Reimportación de movimientos",
            f"Filas leídas: {st.get('leidos', 0)}\n"
            f"Actualizadas (cambiaron): {st.get('actualizados', 0)}\n"
            f"Nuevas (insertadas): {st.get('nuevos', 0)}\n"
            f"Sin cambios: {st.get('sin_cambio', 0)}{_extra}\n\n"
            "Solo se actualizó en la base lo que modificaste.", parent=self)

    # ── Conversión ────────────────────────────────────────────────────────────
    def _start_conversion(self):
        if self._running: return
        if not self._archivos:
            messagebox.showwarning("Sin archivos", "Agrega al menos un PDF."); return
        if all(a.estado == "ok" for a in self._archivos):
            if not messagebox.askyesno("Reconvertir", "Todos ya convertidos. ¿Reconvertir?"):
                return
            for a in self._archivos:
                a.estado = "pendiente"; a.movimientos = 0
        self._running = True
        if _HAS_CTK:
            self._btn_convert.configure(state="disabled", fg_color="#999999", text="⏳  Procesando…")
        else:
            self._btn_convert.config(state="disabled", bg="#999999", text="⏳  Procesando…")
        self._progress["value"] = 0
        self._stop_detect.set()
        threading.Thread(target=self._run_conversion, daemon=True).start()

    def _run_conversion(self):
        outputs = []
        self._bancos_nuevos = []   # rutas de PDFs de bancos NO reconocidos
        total   = len(self._archivos)
        for fi, item in enumerate(self._archivos):
            item.estado = "procesando"
            self.after(0, self._refresh_list)
            ruta_pdf = item.path        # ruta a leer (puede cambiar a la OCR)
            _ocr_tmp = None             # PDF OCR temporal a limpiar al final
            try:
                self._log_write(f"  Detectando: {item.nombre}", "info")
                # ── OCR: si el PDF es imagen (sin texto), generar un PDF con capa
                #    de texto y leer ESE (los parsers lo leen igual: respeta su forma).
                import ocr_pdf as _ocr
                with PDF_LOCK:
                    _sin_texto = not _ocr.pdf_tiene_texto(item.path)
                if _sin_texto:
                    if _ocr.ocr_disponible():
                        self._log_write("  PDF escaneado (imagen) — aplicando OCR…", "info")
                        with PDF_LOCK:
                            _res = _ocr.ocr_a_pdf_buscable(item.path, log=self._log_write)
                        if _res:
                            ruta_pdf = _ocr_tmp = _res
                            self._log_write("  ✔ OCR aplicado (PDF convertido a texto).", "ok")
                        else:
                            raise ValueError("PDF escaneado y el OCR falló. Revisa que "
                                             "el archivo no esté dañado o usa el PDF digital.")
                    else:
                        raise ValueError(
                            "PDF sin texto (escaneo). Instala Tesseract OCR "
                            "(idioma español) o usa el PDF digital del banco.")
                with PDF_LOCK:
                    extractor = detectar_o_generico(ruta_pdf)
                item.banco = extractor.NOMBRE
                self.after(0, self._refresh_list)
                if getattr(extractor, "ES_GENERICO", False):
                    self._log_write(
                        f"[{item.banco}] {item.nombre}  —  banco NUEVO (genérico)", "warn")
                    self._bancos_nuevos.append(item.path)
                else:
                    self._log_write(f"[{item.banco}] {item.nombre}", "bank")

                def progress(cur, tot, fi=fi):
                    base  = (fi / total) * 100
                    chunk = (1 / total) * 100
                    pct   = base + chunk * (cur / tot)
                    self.after(0, lambda p=pct, c=cur, t=tot:
                        (self._progress.__setitem__("value", p),
                         self._lbl_status.config(text=f"Pág {c}/{t}  [{item.banco}] {fi+1}/{total}")))

                with PDF_LOCK:
                    movs, info = extractor.extraer(ruta_pdf, progress_callback=progress)
                if not movs:
                    raise ValueError("Sin movimientos en el PDF.")

                # Candado RFC
                if not info.rfc:
                    raise ValueError("PDF sin RFC. Registra la empresa en 🏢 Empresas / RFC.")
                empresa_nombre = buscar_por_rfc(info.rfc)
                if empresa_nombre is None:
                    raise ValueError(
                        f"RFC {info.rfc} no registrado.\n"
                        "Ve a 🏢 Empresas / RFC en la pestaña Convertidor.")
                self._log_write(f"  🏢  {empresa_nombre}  ({info.rfc})", "info")

                # Ruta de salida
                with PDF_LOCK:
                    año_mes = _buscar_fecha_corte(ruta_pdf)
                if not año_mes:
                    año_mes = _año_mes_fallback(info.periodo)
                base_dir = self._output_dir or str(Path(item.path).parent)
                out_dir  = os.path.join(base_dir, _safe_folder(empresa_nombre), año_mes)
                os.makedirs(out_dir, exist_ok=True)
                out_name = Path(item.path).stem + "_movimientos.xlsx"
                out_path = os.path.join(out_dir, out_name)
                self._log_write(f"  📁  {_safe_folder(empresa_nombre)}/{año_mes}/", "info")

                # Registrar los movimientos en la BD con ID estable → sale en la
                # columna A del Excel (para reimportar y actualizar solo lo que
                # cambie). Defensivo: si la BD falla, el Excel se genera igual.
                try:
                    import persistencia as _pers
                    _mb = [_pers.MovimientoBanco(
                        banco=extractor.NOMBRE, fecha=str(m.fecha or ""),
                        folio=str(m.folio or ""), descripcion=m.descripcion or "",
                        deposito=float(m.deposito or 0), retiro=float(m.retiro or 0),
                        saldo=float(m.saldo or 0)) for m in movs]
                    _p = _pers.FirebirdPersistencia(); _p.inicializar()
                    try:
                        _p.guardar_conversion(item.path, extractor.NOMBRE,
                                              info.rfc or "", info.periodo or "", _mb)
                        for _m, _row in zip(movs, _mb):
                            _m.id_mov = _row.id
                    finally:
                        _p.cerrar()
                    self._log_write(f"  BD: {len(_mb)} movimientos registrados.", "info")
                except Exception as _e_bd:
                    self._log_write(
                        f"  ⚠ BD no disponible ({_e_bd}); Excel sin IDs.", "info")

                export_to_excel(movs, out_path, info,
                                empresa_nombre=empresa_nombre, usuario=self._usuario)
                bitacora_registrar(
                    rfc=info.rfc, empresa=empresa_nombre, banco=extractor.NOMBRE,
                    periodo=año_mes, fecha_corte=info.periodo, movimientos=len(movs),
                    usuario=self._usuario, archivo_origen=item.path, archivo_excel=out_path)
                outputs.append(out_path)
                item.estado      = "ok"
                item.movimientos = len(movs)
                self._log_write(f"  ✔  {len(movs)} movimientos → {out_name}", "ok")

            except Exception as exc:
                item.estado  = "error"
                item.mensaje = str(exc)
                self._log_write(f"  ✘  ERROR: {exc}", "err")
            finally:
                if _ocr_tmp:
                    try: os.remove(_ocr_tmp)
                    except Exception: pass
            self.after(0, self._refresh_list)

        self.after(0, lambda: self._finish(outputs))

    def _finish(self, outputs):
        self._running = False
        if _HAS_CTK:
            self._btn_convert.configure(state="normal", fg_color=C_GREEN, text="▶  CONVERTIR A EXCEL")
        else:
            self._btn_convert.config(state="normal", bg=C_GREEN, text="▶  CONVERTIR A EXCEL")
        self._progress["value"] = 100
        self._lbl_status.config(text="Completado", fg=C_GREEN)
        ok  = sum(1 for a in self._archivos if a.estado == "ok")
        err = sum(1 for a in self._archivos if a.estado == "error")
        msg = f"Conversión finalizada.\n{ok} archivo(s) generados."
        if err: msg += f"\n{err} con error (ver registro)."
        messagebox.showinfo("Completado", msg)
        if outputs and self._var_open.get():
            os.startfile(outputs[-1])
        # ── Bancos NUEVOS: ofrecer generar su parser específico ──────────────
        nuevos = getattr(self, "_bancos_nuevos", [])
        if nuevos:
            vistos, unicos = set(), []
            for p in nuevos:
                if p not in vistos:
                    vistos.add(p); unicos.append(p)
            if messagebox.askyesno(
                    "Banco(s) no reconocido(s)",
                    f"Se convirtieron {len(unicos)} estado(s) de cuenta con el "
                    "extractor GENÉRICO (best-effort).\n\n¿Quieres GENERAR el parser "
                    "específico para ese banco (para mejorar la extracción)?"):
                for p in unicos:
                    self._generar_parser_banco(p)

    def _generar_parser_banco(self, pdf_path: str):
        """Analiza el PDF de un banco nuevo, deja confirmar/editar el nombre,
        keyword y formato de fecha, y ESCRIBE bancos/<slug>.py. Se recoge solo en
        el siguiente arranque (cargador dinámico)."""
        try:
            from bancos import generador
            info = generador.analizar(pdf_path)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo analizar el PDF:\n{e}")
            return

        dlg = tk.Toplevel(self); dlg.title("Generar parser de banco nuevo")
        dlg.configure(bg=T.SURFACE); dlg.transient(self); dlg.resizable(False, False)
        tk.Label(dlg, text="Confirma los datos del banco nuevo",
                 bg=T.SURFACE, fg=T.TEXT, font=("Segoe UI", 12, "bold")
                 ).pack(anchor="w", padx=16, pady=(14, 2))
        tk.Label(dlg, text=f"PDF: {os.path.basename(pdf_path)}",
                 bg=T.SURFACE, fg=T.TEXT_MUTED, font=FONT_SMALL).pack(anchor="w", padx=16, pady=(0, 8))

        card = tk.Frame(dlg, bg=T.SURFACE); card.pack(fill="x", padx=12); card.columnconfigure(1, weight=1)
        _v = {}
        def _f(r, etq, clave, hint=""):
            tk.Label(card, text=etq, bg=T.SURFACE, fg=C_TEXT, font=FONT_HEAD
                     ).grid(row=r, column=0, sticky="w", padx=(12, 8), pady=6)
            var = tk.StringVar(value=str(info.get(clave, ""))); _v[clave] = var
            tk.Entry(card, textvariable=var, font=FONT_NORMAL, width=42, bg=C_GRAY,
                     fg=C_TEXT, relief="flat", bd=4).grid(row=r, column=1, sticky="ew", pady=6)
            if hint:
                tk.Label(card, text=hint, bg=T.SURFACE, fg=C_GRAY2, font=("Segoe UI", 8)
                         ).grid(row=r, column=2, sticky="w", padx=(8, 12))
        _f(0, "Nombre del banco:", "nombre")
        _f(1, "Palabra clave (detectar):", "keyword", "aparece en el encabezado")
        _f(2, "Formato de fecha (regex):", "date_src", "auto-detectado")

        bf = tk.Frame(dlg, bg=T.SURFACE); bf.pack(fill="x", padx=16, pady=12)
        def _crear():
            info["nombre"]   = _v["nombre"].get().strip() or info["nombre"]
            info["keyword"]  = _v["keyword"].get().strip().upper() or info["keyword"]
            info["date_src"] = _v["date_src"].get().strip() or info["date_src"]
            info["slug"]     = generador._slug(info["nombre"])
            info["clase"]    = generador._clase(info["slug"])
            try:
                from bancos import generador as _g
                # Reusar la plantilla con los valores confirmados. Se escribe en la
                # carpeta ESCRIBIBLE (%PROGRAMDATA%\...\bancos_custom): dentro del
                # .exe el paquete bancos/ vive en el PYZ y no admite archivos nuevos.
                _carpeta = _g.carpeta_bancos_custom()
                ruta = os.path.join(_carpeta, f"{info['slug']}.py")
                n = 2
                while os.path.exists(ruta):
                    ruta = os.path.join(_carpeta, f"{info['slug']}_{n}.py"); n += 1
                codigo = _g._PLANTILLA.format(
                    nombre=info["nombre"].replace('"', "'"),
                    pdf_base=os.path.basename(pdf_path).replace('"', "'"),
                    date_src=info["date_src"], clase=info["clase"], keyword=info["keyword"])
                with open(ruta, "w", encoding="utf-8") as fh:
                    fh.write(codigo)
            except Exception as e:
                messagebox.showerror("Error al generar", str(e), parent=dlg); return
            dlg.destroy()
            messagebox.showinfo(
                "Parser generado",
                f"Se creó:\n{ruta}\n\nEl banco '{info['nombre']}' quedará reconocido "
                "al reiniciar la app. Puedes afinar el archivo si algún movimiento "
                "no sale bien (ver bancos/banbajio.py como ejemplo).")
        self._rbtn(bf, "✔  Generar parser", _crear, "accent").pack(side="left")
        self._rbtn(bf, "Cancelar", dlg.destroy, "ghost").pack(side="left", padx=(8, 0))
        dlg.update_idletasks()
        try:
            px, py = self.winfo_rootx(), self.winfo_rooty()
            dlg.geometry(f"+{px + 120}+{py + 120}")
        except Exception:
            pass

    # ══════════════════════════════════════════════════════════════════════════
    #  TAB: COMPLIANCE FISCAL
    # ══════════════════════════════════════════════════════════════════════════
    def _build_compliance_tab(self):
        f = self._tab_compliance

        # ── Cabecera de módulo (blanca, estilo prototipo) ────────────────────
        self._module_header(
            f, "⚖", "Obligaciones Fiscales",
            "Panel de cumplimiento, obligaciones por empresa, archivos y configuración")

        # Sub-notebook dentro del tab
        subnb = ttk.Notebook(f); subnb.pack(fill="both", expand=True)
        style = ttk.Style()
        style.configure("Sub.TNotebook",     background=C_BG, borderwidth=0)
        style.configure("Sub.TNotebook.Tab", background="#DDE4EF", foreground=C_TEXT,
                        padding=[14, 6], font=("Segoe UI", 9))
        style.map("Sub.TNotebook.Tab",
                  background=[("selected", C_BLUE)],
                  foreground=[("selected", C_WHITE)])

        self._cpanel   = tk.Frame(subnb, bg=C_BG)
        self._cobligs  = tk.Frame(subnb, bg=C_BG)
        self._csaldos  = tk.Frame(subnb, bg=C_BG)
        self._carch    = tk.Frame(subnb, bg=C_BG)
        self._cconfig  = tk.Frame(subnb, bg=C_BG)

        subnb.add(self._cpanel,  text="📊  Panel General")
        subnb.add(self._cobligs, text="📋  Obligaciones")
        subnb.add(self._csaldos, text="💰  Contribuciones, Pagos y Saldos a Favor")
        subnb.add(self._carch,   text="📁  Archivos")
        subnb.add(self._cconfig, text="⚙  Config Empresa")

        self._build_compliance_panel()
        self._build_compliance_obligs()
        self._build_saldos_tab(self._csaldos)     # integrado como pestaña
        self._build_compliance_archivos()
        self._build_compliance_config()

        # Al cambiar de sub-pestaña dentro de Obligaciones, refresca los combos de
        # RFC (para que Obligaciones/Archivos vean lo registrado en Config Empresa).
        subnb.bind("<<NotebookTabChanged>>", lambda _e: self._refrescar_rfcs())

    # Mapa categoría (sat_saldos) → obligación del catálogo compliance
    _CAT_A_OBLIG = {
        "ISR_PM":             "isr_declaracion",
        "ISR_RET_SUELDOS":    "isr_ret_declaracion",
        "ISR_RET_ASIMILADOS": "isr_ret_declaracion",
        "IVA":                "iva_declaracion",
        "IVA_RET":            "iva_ret_declaracion",
    }

    def _abrir_anios_trabajo(self):
        """Sub-ventana para capturar MANUALMENTE los años de trabajo (lista global).
        Acepta años sueltos y rangos: "2015, 2018, 2024-2026". Se reflejan en los
        combos de Año de Panel General, Obligaciones, Contribuciones/Saldos y Archivos."""
        dlg = tk.Toplevel(self); dlg.title("Años de trabajo"); dlg.transient(self)
        dlg.configure(bg=T.SURFACE); dlg.resizable(False, False); dlg.grab_set()
        tk.Label(dlg, text="Años de trabajo", bg=T.SURFACE, fg=T.TEXT,
                 font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=16, pady=(14, 2))
        tk.Label(dlg, text=f"Captura los años que necesites (permitido "
                 f"{CE.ANIO_MIN}–{CE.ANIO_MAX}). Ejemplos: 2015, 2018, 2024-2026",
                 bg=T.SURFACE, fg=T.TEXT_MUTED, font=FONT_SMALL, justify="left",
                 wraplength=460).pack(anchor="w", padx=16, pady=(0, 8))

        v_txt = tk.StringVar(value=CE.formatear_anios(CE.cargar_anios()))
        ent = tk.Entry(dlg, textvariable=v_txt, font=FONT_NORMAL, width=52,
                       bg=C_GRAY, fg=C_TEXT, relief="flat", bd=6)
        ent.pack(anchor="w", padx=16)
        lbl_prev = tk.Label(dlg, text="", bg=T.SURFACE, fg=T.ACCENT, font=FONT_SMALL,
                            justify="left", wraplength=460)
        lbl_prev.pack(anchor="w", padx=16, pady=(6, 0))

        def _prev(*_a):
            ys = CE.parsear_anios(v_txt.get())
            if ys:
                lbl_prev.config(text=f"✔ {len(ys)} año(s): {CE.formatear_anios(ys)}",
                                fg=T.ACCENT)
            else:
                lbl_prev.config(text="⚠ Sin años válidos (se usará 2015–2040 completo).",
                                fg=C_WARN)
        v_txt.trace_add("write", _prev); _prev()

        # Atajos rápidos
        qf = tk.Frame(dlg, bg=T.SURFACE); qf.pack(anchor="w", padx=16, pady=(8, 0))
        def _quick(txt):
            v_txt.set(txt)
        _hoy = date.today().year
        self._rbtn(qf, "Todo 2015–2040", lambda: _quick(f"{CE.ANIO_MIN}-{CE.ANIO_MAX}"),
                   "ghost", small=True).pack(side="left", padx=(0, 6))
        self._rbtn(qf, f"Últimos 5 ({_hoy-4}-{_hoy})",
                   lambda: _quick(f"{_hoy-4}-{_hoy}"), "ghost", small=True).pack(side="left")

        bf2 = tk.Frame(dlg, bg=T.SURFACE); bf2.pack(fill="x", padx=16, pady=12)
        def _guardar():
            ys = CE.parsear_anios(v_txt.get())
            CE.guardar_anios(ys)
            self._refrescar_anios()
            dlg.destroy()
            messagebox.showinfo(
                "Años de trabajo",
                ("Años guardados: " + CE.formatear_anios(ys)) if ys else
                "Sin años válidos: se usará el rango completo 2015–2040.",
                parent=self)
        self._rbtn(bf2, "💾 Guardar", _guardar, "accent").pack(side="left")
        self._rbtn(bf2, "Cancelar", dlg.destroy, "ghost").pack(side="left", padx=(8, 0))
        dlg.update_idletasks()
        try:
            px, py = self.winfo_rootx(), self.winfo_rooty()
            dlg.geometry(f"+{px + 130}+{py + 120}")
        except Exception:
            pass

    # ══════════════════════════════════════════════════════════════════════
    #  REPORTES POR PESTAÑA (Excel / PDF / CSV / TXT) del mes-año seleccionado
    # ══════════════════════════════════════════════════════════════════════
    def _anio_default(self) -> str:
        """Año preseleccionado: el actual si está en la lista configurada; si no, el
        más reciente de la lista."""
        try:
            ys = CE.anios_disponibles()
            hoy = date.today().year
            if hoy in ys:
                return str(hoy)
            return str(max(ys)) if ys else str(hoy)
        except Exception:
            return str(date.today().year)

    def _refrescar_anios(self):
        """Repuebla los combos de Año de TODAS las pestañas con la lista configurada
        en Config Empresa → «Años de trabajo». Conserva la selección si sigue válida."""
        try:
            vals = [str(y) for y in CE.anios_disponibles()]
        except Exception:
            return
        for nombre in ("_cp_anio", "_co_anio", "_ca_anio", "_sa_anio", "_concil_anio"):
            combo = getattr(self, nombre, None)
            if combo is None:
                continue
            try:
                actual = combo.get()
                combo["values"] = vals
                if actual not in vals:
                    combo.set(self._anio_default())
            except Exception:
                pass

    def _refrescar_rfcs(self):
        """Repobla los combos de RFC de todos los módulos/pestañas desde el catálogo
        central de empresas (FDB). Conserva la selección actual si sigue existiendo."""
        try:
            rfcs = [rfc for rfc, _ in listar_empresas()]
        except Exception:
            return
        for nombre in ("_cb_sel_rfc", "_ca_combo", "_cc_combo"):
            combo = getattr(self, nombre, None)
            if combo is None:
                continue
            try:
                actual = combo.get()
                combo["values"] = rfcs
                if actual not in rfcs and rfcs and not combo.get():
                    combo.set(rfcs[0])
            except Exception:
                pass
        # Conciliación usa su propio formato "nombre — RFC"
        try:
            self._concil_refresh_empresas()
        except Exception:
            pass

    def _cargar_declaracion(self):
        """Carga una o varias declaraciones/acuses del SAT (multi-selección) y las
        procesa AUTOMÁTICAMENTE: lee cada PDF, clasifica el impuesto (ISR/IVA/
        retenciones), toma RFC, periodo, número de operación y fecha DEL PROPIO PDF,
        marca la obligación (✅) y alimenta Contribuciones/Saldos si trae ISR/IVA.
        Sin lista desplegable ni captura manual."""
        rutas = filedialog.askopenfilenames(
            title="Selecciona uno o varios PDF/ZIP de declaraciones o acuses del SAT",
            filetypes=[("PDF/ZIP", "*.pdf *.zip"), ("Todos", "*.*")])
        if not rutas:
            return
        self._cargar_y_procesar(list(rutas))

    def _cargar_y_procesar(self, entradas):
        """Expande (archivos/zip/carpeta) + lee con sat_saldos y procesa. Compartido
        por «Cargar declaración» (archivos) y «Cargar carpeta»."""
        if not _SALDOS_OK:
            messagebox.showwarning("No disponible",
                "El motor de lectura de PDFs (sat_saldos) no está disponible.",
                parent=self)
            return
        from pathlib import Path as _P
        try:
            pdfs = _SF_PIPE.expandir_entradas([_P(x) for x in entradas])
            cargados = _SF_PIPE.cargar(pdfs)
        except Exception as e:
            messagebox.showerror("Error al leer", str(e), parent=self)
            return
        # ── PRIMERO: verificar a qué RFC pertenecen los archivos ────────────────
        # Sin esto, mezclar empresas produce análisis cruzados (RFC de una con la
        # razón social de otra) y marcas en la empresa equivocada.
        mapa: dict = {}

        def _reg(rfc, archivo):
            rfc = (rfc or "").strip().upper()
            if not rfc:
                return
            d = mapa.setdefault(rfc, {"docs": 0, "archivos": set(), "nombre": None})
            d["docs"] += 1
            if archivo:
                d["archivos"].add(str(archivo))

        for dec in getattr(cargados, "declaraciones", []) or []:
            _reg(dec.rfc, dec.archivo)
        for pg in getattr(cargados, "pagos", []) or []:
            _reg(getattr(pg, "rfc", None), getattr(pg, "archivo", None))
        for det in getattr(cargados, "detalles", []) or []:
            _reg(getattr(det, "rfc", None), getattr(det, "archivo", None))
        for _r in mapa:
            mapa[_r]["nombre"] = buscar_por_rfc(_r)

        if not mapa:
            messagebox.showwarning(
                "Sin RFC identificable",
                "No se pudo identificar el RFC en los archivos seleccionados "
                "(¿PDF escaneados sin texto o no son acuses/declaraciones del SAT?).",
                parent=self)
            return

        seleccion = self._confirmar_rfcs(mapa)
        if not seleccion:
            return                      # cancelado o nada seleccionado

        # Alimentar Contribuciones/Saldos SOLO con los archivos del/los RFC elegidos
        archivos_ok = set()
        for _r in seleccion:
            archivos_ok |= mapa[_r]["archivos"]
        if hasattr(self, "_saldos_rutas") and archivos_ok:
            try:
                self._saldos_sumar([_P(x) for x in sorted(archivos_ok)])
            except Exception:
                pass
        # Cargar el campo RFC de Contribuciones/Saldos YA (antes se quedaba vacío
        # hasta pulsar ANALIZAR; ahora la carga se hace desde «Archivos»).
        if hasattr(self, "_saldos_rfc_combo"):
            try:
                _vals = sorted(seleccion)
                self._saldos_rfc_combo["values"] = _vals
                if _vals and self._saldos_var_rfc.get() not in _vals:
                    self._saldos_var_rfc.set(_vals[0])
            except Exception:
                pass
        # Marcar obligaciones + copiar a las carpetas por periodo (solo esos RFC)
        self._procesar_declaraciones(cargados, rfcs_permitidos=seleccion)
        # Refrescar el explorador de Archivos (se acaban de copiar PDFs)
        try:
            self._refrescar_archivos()
        except Exception:
            pass

    def _confirmar_rfcs(self, mapa: dict):
        """Muestra los RFC detectados en los archivos y deja elegir cuáles cargar.
        Devuelve el set de RFC seleccionados, o None si se cancela."""
        dlg = tk.Toplevel(self); dlg.title("Verificar RFC de los archivos")
        dlg.transient(self); dlg.configure(bg=T.SURFACE)
        dlg.resizable(False, False); dlg.grab_set()
        tk.Label(dlg, text="Verifica a qué empresa pertenecen los archivos",
                 bg=T.SURFACE, fg=T.TEXT, font=("Segoe UI", 12, "bold")
                 ).pack(anchor="w", padx=16, pady=(14, 2))
        tk.Label(dlg, text="Solo se cargarán los RFC que dejes marcados.",
                 bg=T.SURFACE, fg=T.TEXT_MUTED, font=FONT_SMALL
                 ).pack(anchor="w", padx=16, pady=(0, 8))
        cont = tk.Frame(dlg, bg=T.SURFACE); cont.pack(fill="x", padx=16)
        _vars = {}
        for rfc, info in sorted(mapa.items()):
            registrado = info.get("nombre")
            v = tk.BooleanVar(value=bool(registrado))   # no registrados: desmarcados
            _vars[rfc] = v
            row = tk.Frame(cont, bg=T.SURFACE); row.pack(fill="x", pady=2)
            tk.Checkbutton(row, variable=v, bg=T.SURFACE, activebackground=T.SURFACE
                           ).pack(side="left")
            tk.Label(row, text=f"{rfc}   ·   {info['docs']} doc(s)", bg=T.SURFACE,
                     fg=T.TEXT, font=("Consolas", 9)).pack(side="left")
            if registrado:
                tk.Label(row, text=f"   {registrado}", bg=T.SURFACE, fg=T.ACCENT,
                         font=FONT_SMALL).pack(side="left")
            else:
                tk.Label(row, text="   ⚠ RFC no registrado (regístralo en Config Empresa)",
                         bg=T.SURFACE, fg=C_WARN, font=FONT_SMALL).pack(side="left")
        if len(mapa) > 1:
            tk.Label(dlg, text="⚠ Hay archivos de varias empresas. Cárgalas por "
                     "separado si vas a analizar saldos, para no mezclar.",
                     bg=T.SURFACE, fg=C_WARN, font=FONT_SMALL, justify="left",
                     wraplength=520).pack(anchor="w", padx=16, pady=(8, 0))
        res = {"ok": False}
        bf = tk.Frame(dlg, bg=T.SURFACE); bf.pack(fill="x", padx=16, pady=12)
        def _ok():
            res["ok"] = True; dlg.destroy()
        self._rbtn(bf, "✔  Cargar seleccionados", _ok, "accent").pack(side="left")
        self._rbtn(bf, "Cancelar", dlg.destroy, "ghost").pack(side="left", padx=(8, 0))
        dlg.update_idletasks()
        try:
            px, py = self.winfo_rootx(), self.winfo_rooty()
            dlg.geometry(f"+{px + 120}+{py + 110}")
        except Exception:
            pass
        self.wait_window(dlg)
        if not res["ok"]:
            return None
        return {r for r, v in _vars.items() if v.get()}

    def _cargar_declaraciones_carpeta(self):
        """Carga MASIVA por carpeta (o ZIP): mismo procesamiento automático que
        «Cargar declaración», para subir muchas de una vez sin ir una por una."""
        d = filedialog.askdirectory(
            title="Carpeta con PDFs de declaraciones/acuses del SAT")
        if not d:
            return
        self._cargar_y_procesar([d])

    def _procesar_declaraciones(self, cargados, rfcs_permitidos=None):
        """Motor común (archivos y carpeta): por cada declaración/pago leído marca su
        obligación con RFC/periodo/impuesto/operación/fecha DEL PDF, copia el PDF a la
        carpeta del periodo y alimenta Contribuciones/Saldos las de ISR/IVA. Solo
        empresas registradas; anuales / sin periodo / no clasificadas se reportan;
        idempotente (dedupe por rfc/año/mes/obligación)."""
        import shutil
        from pathlib import Path as _P
        from sat_saldos.config import clasifica_concepto

        def _f(x):
            try: return float(x or 0)
            except Exception: return 0.0

        def _fecha_iso(fp):
            try:
                return fp.date().isoformat() if hasattr(fp, "date") else (str(fp)[:10] if fp else "")
            except Exception:
                return ""

        hechas = set()                      # (rfc, año, mes, oblig) sin duplicar
        rutas_saldos: set = set()           # PDFs con ISR/IVA → alimentan Saldos
        marcadas = 0
        empresas: set = set()
        no_reg: set = set()
        omit_no_reg = omit_anual = omit_sin_periodo = omit_no_clasif = 0

        def _procesa(rfc, ejercicio, periodo, es_anual, operacion, fp, archivo, conceptos):
            nonlocal marcadas, omit_no_reg, omit_anual, omit_sin_periodo, omit_no_clasif
            rfc = (rfc or "").strip().upper()
            if not rfc:
                return
            # Solo los RFC verificados/elegidos por el usuario
            if rfcs_permitidos is not None and rfc not in rfcs_permitidos:
                return
            if not buscar_por_rfc(rfc):
                omit_no_reg += 1; no_reg.add(rfc); return
            if es_anual:
                omit_anual += 1; return
            if not periodo or not (1 <= int(periodo) <= 12) or not ejercicio:
                omit_sin_periodo += 1; return
            anio = int(ejercicio); mes = int(periodo)
            # ¿trae montos ISR/IVA? (para decidir el feed a Saldos)
            tiene_amt = False
            obligs = set()
            for c in conceptos or []:
                cat = clasifica_concepto(getattr(c, "nombre", "") or "")
                ob = self._CAT_A_OBLIG.get(cat)
                if not ob:
                    continue
                obligs.add(ob)
                if (_f(getattr(c, "a_cargo", None)) or _f(getattr(c, "a_favor", None))
                        or _f(getattr(c, "cantidad_pagar", None))
                        or _f(getattr(c, "cantidad_a_pagar", None))
                        or _f(getattr(c, "cantidad_cargo", None))
                        or _f(getattr(c, "cantidad_a_cargo", None))):
                    tiene_amt = True
            if not obligs:
                omit_no_clasif += 1
                return
            # Copiar el PDF a la carpeta del periodo de la empresa (si está configurada)
            cfg = CE.obtener_config(rfc)
            destino_dir = CE.asegurar_carpeta_mes(cfg, anio, mes)
            guardado = archivo or ""
            if destino_dir and archivo and os.path.isfile(archivo):
                try:
                    dest = os.path.join(destino_dir, os.path.basename(archivo))
                    if os.path.abspath(dest) != os.path.abspath(archivo):
                        shutil.copy2(archivo, dest)
                    guardado = dest
                except Exception:
                    guardado = archivo
            fecha = _fecha_iso(fp)
            for ob in obligs:
                clave = (rfc, anio, mes, ob)
                if clave in hechas:
                    continue
                hechas.add(clave)
                try:
                    CE.marcar_presentacion(rfc, anio, mes, ob, archivo=guardado,
                                           fecha_presentacion=fecha,
                                           numero_operacion=operacion or "",
                                           usuario=self._usuario)
                    marcadas += 1; empresas.add(rfc)
                except Exception:
                    pass
            if tiene_amt and guardado:
                rutas_saldos.add(guardado)

        for dec in getattr(cargados, "declaraciones", []) or []:
            _procesa(dec.rfc, dec.ejercicio, dec.periodo, getattr(dec, "es_anual", False),
                     dec.operacion, dec.fecha_presentacion, dec.archivo, dec.conceptos)
        for pg in getattr(cargados, "pagos", []) or []:
            _procesa(getattr(pg, "rfc", None), getattr(pg, "ejercicio", None),
                     getattr(pg, "periodo", None), False, getattr(pg, "operacion", None),
                     getattr(pg, "fecha_presentacion", None), getattr(pg, "archivo", None),
                     getattr(pg, "conceptos", None))

        # Alimentar Contribuciones/Saldos con las que traen ISR/IVA
        alimentadas = 0
        if rutas_saldos and hasattr(self, "_saldos_rutas"):
            try:
                self._saldos_sumar([_P(x) for x in rutas_saldos])
                alimentadas = len(rutas_saldos)
            except Exception:
                alimentadas = 0

        try:
            self._cargar_obligs(); self._reload_compliance_panel()
        except Exception:
            pass

        partes = [f"✅ {marcadas} obligación(es) marcada(s) en {len(empresas)} empresa(s)."]
        if alimentadas:
            partes.append(f"{alimentadas} con ISR/IVA → «Contribuciones, Pagos y Saldos "
                          "a Favor» (pulsa ▶ ANALIZAR ahí).")
        if omit_no_reg:
            partes.append(f"{omit_no_reg} de RFC no registrado ({', '.join(sorted(no_reg))}) "
                          "— regístralo en el catálogo de empresas.")
        if omit_anual:
            partes.append(f"{omit_anual} anual(es) — no se marcan automáticamente.")
        if omit_sin_periodo:
            partes.append(f"{omit_sin_periodo} sin periodo mensual claro.")
        if omit_no_clasif:
            partes.append(f"{omit_no_clasif} no reconocida(s) como ISR/IVA/retención — revísalas.")
        if not any((marcadas, alimentadas, omit_no_reg, omit_anual, omit_sin_periodo, omit_no_clasif)):
            partes = ["No se leyó ninguna declaración/acuse en la selección."]
        messagebox.showinfo("Carga de declaraciones", "\n".join(partes), parent=self)

    # ══════════════════════════════════════════════════════════════════════
    #  REPORTES POR PESTAÑA (Excel / PDF / CSV / TXT) del mes-año seleccionado
    # ══════════════════════════════════════════════════════════════════════
    def _generar_reporte(self, rep):
        """Pide formato y ruta, y exporta el Reporte recibido."""
        import reportes as RP
        if not rep.filas:
            if not messagebox.askyesno(
                    "Reporte vacío",
                    "No hay datos para el periodo seleccionado.\n\n¿Generar de todos "
                    "modos (solo encabezados)?", parent=self):
                return
        dlg = tk.Toplevel(self); dlg.title("Generar reporte"); dlg.transient(self)
        dlg.configure(bg=T.SURFACE); dlg.resizable(False, False); dlg.grab_set()
        tk.Label(dlg, text=rep.titulo, bg=T.SURFACE, fg=T.TEXT,
                 font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=16, pady=(14, 2))
        tk.Label(dlg, text="  ·  ".join([x for x in rep.meta_lineas()[:2]]),
                 bg=T.SURFACE, fg=T.TEXT_MUTED, font=FONT_SMALL, justify="left",
                 wraplength=460).pack(anchor="w", padx=16, pady=(0, 2))
        tk.Label(dlg, text=f"{len(rep.filas)} renglón(es)", bg=T.SURFACE,
                 fg=T.ACCENT, font=FONT_SMALL).pack(anchor="w", padx=16, pady=(0, 8))
        tk.Label(dlg, text="Formato:", bg=T.SURFACE, fg=C_TEXT,
                 font=FONT_HEAD).pack(anchor="w", padx=16)
        v_fmt = tk.StringVar(value="xlsx")
        ff = tk.Frame(dlg, bg=T.SURFACE); ff.pack(anchor="w", padx=16, pady=(2, 8))
        for f in RP.FORMATOS:
            tk.Radiobutton(ff, text=RP.NOMBRE_FORMATO[f], variable=v_fmt, value=f,
                           bg=T.SURFACE, fg=T.TEXT, activebackground=T.SURFACE,
                           selectcolor=C_WHITE, font=FONT_SMALL).pack(side="left", padx=(0, 10))
        bf = tk.Frame(dlg, bg=T.SURFACE); bf.pack(fill="x", padx=16, pady=12)

        def _ok():
            fmt = v_fmt.get()
            dlg.destroy()
            ruta = filedialog.asksaveasfilename(
                title="Guardar reporte", defaultextension=RP.EXT_FORMATO[fmt],
                initialfile=RP.nombre_sugerido(rep, fmt),
                filetypes=[(RP.NOMBRE_FORMATO[fmt], "*" + RP.EXT_FORMATO[fmt]),
                           ("Todos", "*.*")])
            if not ruta:
                return
            try:
                RP.exportar(rep, ruta, fmt)
            except Exception as e:
                messagebox.showerror("Error al generar el reporte", str(e), parent=self)
                return
            if messagebox.askyesno("Reporte generado",
                                   f"✅ Se generó:\n{ruta}\n\n¿Abrirlo ahora?",
                                   parent=self):
                try:
                    os.startfile(ruta)
                except Exception:
                    pass
        self._rbtn(bf, "📊  Generar", _ok, "accent").pack(side="left")
        self._rbtn(bf, "Cancelar", dlg.destroy, "ghost").pack(side="left", padx=(8, 0))
        dlg.update_idletasks()
        try:
            px, py = self.winfo_rootx(), self.winfo_rooty()
            dlg.geometry(f"+{px + 130}+{py + 120}")
        except Exception:
            pass

    def _reporte_panel(self):
        """Reporte del Panel General: cumplimiento de TODAS las empresas del periodo."""
        import reportes as RP
        try:
            mes = CE.MESES_ES.index(self._cp_mes.get()); anio = int(self._cp_anio.get())
        except Exception:
            messagebox.showwarning("Periodo", "Elige mes y año.", parent=self); return
        filas, tot_ob, tot_pres = [], 0, 0
        for rfc, nombre in listar_empresas():
            cfg = CE.obtener_config(rfc)
            obs = CE.obligaciones_del_mes(cfg, mes)
            try:
                pres = CE.obtener_presentaciones(rfc, anio, mes)
            except Exception:
                pres = {}
            n_pres = sum(1 for o in obs if o["id"] in pres)
            arcs = CE.listar_archivos_mes(cfg, anio, mes)
            pend = len(obs) - n_pres
            tot_ob += len(obs); tot_pres += n_pres
            filas.append([nombre, rfc, cfg.get("destinatario", ""), len(obs), n_pres,
                          pend, len(arcs),
                          f"{(n_pres / len(obs) * 100):.0f}%" if obs else "—"])
        rep = RP.Reporte(
            titulo="Panel General de Obligaciones",
            subtitulo="Cumplimiento por empresa del periodo",
            periodo=f"{CE.MESES_ES[mes]} {anio}",
            encabezados=["Empresa", "RFC", "Destinatario", "Obligaciones",
                         "Presentadas", "Pendientes", "Archivos", "Cumplimiento"],
            filas=filas,
            resumen=[("Empresas", len(filas)), ("Obligaciones totales", tot_ob),
                     ("Presentadas", tot_pres), ("Pendientes", tot_ob - tot_pres)])
        self._generar_reporte(rep)

    def _reporte_obligaciones(self):
        """Reporte de la pestaña Obligaciones: estatus por obligación de la empresa."""
        import reportes as RP
        rfc = self._c_sel_rfc.get().strip().upper()
        nombre = buscar_por_rfc(rfc) if rfc else None
        if not nombre:
            messagebox.showwarning("Sin empresa", "Selecciona una empresa registrada.",
                                   parent=self); return
        try:
            mes = CE.MESES_ES.index(self._co_mes.get()); anio = int(self._co_anio.get())
        except Exception:
            messagebox.showwarning("Periodo", "Elige mes y año.", parent=self); return
        cfg = CE.obtener_config(rfc)
        activas = set(cfg.get("obligaciones_activas", CE.TODOS_IDS))
        try:
            CE.sincronizar_presentaciones(rfc, anio, mes)
        except Exception:
            pass
        pres = CE.obtener_presentaciones(rfc, anio, mes)
        filas, n_pres = [], 0
        for o in CE.OBLIGACIONES_CATALOGO:
            if o["id"] not in activas or mes not in o["meses_aplica"]:
                continue
            pr = pres.get(o["id"])
            if pr:
                n_pres += 1
            filas.append([o["nombre"], o["periodicidad"], "Sí" if pr else "No",
                          (pr or {}).get("fecha_presentacion", ""),
                          (pr or {}).get("numero_operacion", ""),
                          os.path.basename((pr or {}).get("archivo", "") or "")])
        rep = RP.Reporte(
            titulo="Obligaciones", subtitulo="Estatus de cumplimiento del periodo",
            empresa=nombre, rfc=rfc, periodo=f"{CE.MESES_ES[mes]} {anio}",
            encabezados=["Obligación", "Periodicidad", "Presentada",
                         "Fecha presentación", "No. operación", "Archivo"],
            filas=filas,
            resumen=[("Obligaciones del mes", len(filas)), ("Presentadas", n_pres),
                     ("Pendientes", len(filas) - n_pres),
                     ("Cumplimiento", f"{(n_pres / len(filas) * 100):.0f}%" if filas else "—")])
        self._generar_reporte(rep)

    def _reporte_saldos(self):
        """Reporte de Contribuciones, Pagos y Saldos a Favor (del RFC y periodo)."""
        import reportes as RP
        rfc = (self._saldos_var_rfc.get() or "").strip().upper()
        if not rfc:
            messagebox.showwarning("Sin RFC",
                "Carga y analiza documentos primero (pestaña Archivos → ANALIZAR).",
                parent=self); return
        try:
            mes = CE.MESES_ES.index(self._sa_mes.get()); anio = int(self._sa_anio.get())
        except Exception:
            messagebox.showwarning("Periodo", "Elige mes y año.", parent=self); return
        filas = []
        p = None
        try:
            from persistencia import FirebirdPersistencia
            p = FirebirdPersistencia(); p.inicializar()
        except Exception:
            p = None
        if p is not None:
            try:
                for d in p.leer_saldos_iva(rfc, anio, mes):
                    filas.append(["IVA", CE.MESES_ES[int(d["mes"])], d.get("operacion", ""),
                                  d.get("fecha", ""), d.get("a_cargo", 0),
                                  d.get("a_favor", 0), d.get("acreditado", 0),
                                  d.get("remanente", 0)])
                for d in p.leer_saldos_pago(rfc, anio, mes):
                    filas.append(["Pago " + (d.get("tipo") or ""), CE.MESES_ES[int(d["mes"])],
                                  d.get("operacion", ""), d.get("fecha_pago", ""),
                                  d.get("contribucion", 0), 0, 0,
                                  d.get("importe_pagado", 0)])
            except Exception:
                pass
            try:
                p.cerrar()
            except Exception:
                pass
        rep = RP.Reporte(
            titulo="Contribuciones, Pagos y Saldos a Favor",
            subtitulo="Información procesada guardada en la base de datos",
            empresa=buscar_por_rfc(rfc) or "", rfc=rfc,
            periodo=f"{CE.MESES_ES[mes]} {anio}",
            encabezados=["Concepto", "Mes", "No. operación", "Fecha", "A cargo",
                         "A favor", "Acreditado", "Remanente / Pagado"],
            filas=filas,
            resumen=[("Renglones", len(filas))])
        self._generar_reporte(rep)

    def _reporte_archivos(self):
        """Reporte de la pestaña Archivos: inventario de la carpeta del periodo."""
        import reportes as RP
        rfc = self._ca_rfc.get().strip().upper()
        nombre = buscar_por_rfc(rfc) if rfc else None
        if not nombre:
            messagebox.showwarning("Sin empresa", "Selecciona una empresa registrada.",
                                   parent=self); return
        try:
            mes = CE.MESES_ES.index(self._ca_mes.get()); anio = int(self._ca_anio.get())
        except Exception:
            messagebox.showwarning("Periodo", "Elige mes y año.", parent=self); return
        cfg = CE.obtener_config(rfc)
        arcs = CE.listar_archivos_mes(cfg, anio, mes)
        ruta = CE.carpeta_mes(cfg, anio, mes) or ""
        filas = [[a["nombre"], os.path.splitext(a["nombre"])[1].lower().lstrip("."),
                  CE.formato_tamano(a["tamano"]), a["mtime"]] for a in arcs]
        total = sum(a["tamano"] for a in arcs)
        rep = RP.Reporte(
            titulo="Archivos", subtitulo=f"Inventario de la carpeta del periodo. {ruta}",
            empresa=nombre, rfc=rfc, periodo=f"{CE.MESES_ES[mes]} {anio}",
            encabezados=["Archivo", "Tipo", "Tamaño", "Fecha modificación"],
            filas=filas,
            resumen=[("Archivos", len(filas)), ("Tamaño total", CE.formato_tamano(total))])
        self._generar_reporte(rep)

    def _saldos_persistir(self):
        """Guarda en la FDB el resultado ya procesado de Contribuciones/Saldos, para
        dejarlo almacenado y poder reportarlo después."""
        res = getattr(self, "_saldos_res", None)
        if not res:
            return
        rfc = (self._saldos_var_rfc.get() or getattr(res, "rfc", "") or "").strip().upper()
        if not rfc:
            return
        try:
            from persistencia import FirebirdPersistencia
            p = FirebirdPersistencia(); p.inicializar()
        except Exception:
            return
        try:
            acum = getattr(res, "iva_acumulado", None)
            for r in (getattr(acum, "renglones", None) or []):
                ej = getattr(r, "ejercicio", None); pe = getattr(r, "periodo", None)
                if not ej or not pe:
                    continue
                fp = getattr(r, "fecha_presentacion", None)
                p.guardar_saldo_iva(
                    rfc, ej, pe, operacion=str(getattr(r, "operacion", "") or ""),
                    fecha=(fp.strftime("%Y-%m-%d") if hasattr(fp, "strftime") else str(fp or "")),
                    a_cargo=getattr(r, "a_cargo", 0) or 0,
                    a_favor=getattr(r, "a_favor_generado", 0) or 0,
                    acreditado=getattr(r, "acreditamiento", 0) or 0,
                    remanente=getattr(r, "remanente", 0) or 0)
            for pg in (getattr(res, "pagos", None) or []):
                ej = getattr(pg, "ejercicio", None); pe = getattr(pg, "periodo", None)
                if not ej or not pe:
                    continue
                p.guardar_saldo_pago(
                    rfc, ej, pe, str(getattr(pg, "operacion", "") or "-"),
                    tipo=str(getattr(pg, "tipo", "") or ""),
                    vencimiento=str(getattr(pg, "vencimiento", "") or ""),
                    contribucion=getattr(pg, "contribucion", 0) or 0,
                    actualizacion=getattr(pg, "actualizacion", 0) or 0,
                    recargos=getattr(pg, "recargos", 0) or 0,
                    importe_pagar=getattr(pg, "importe_a_pagar", 0) or 0,
                    importe_pagado=getattr(pg, "importe_pagado", 0) or 0,
                    fecha_pago=str(getattr(pg, "fecha_pago", "") or ""),
                    estatus=("Pagado" if getattr(pg, "pagado", False) else "Sin comprobante"))
        except Exception:
            pass
        finally:
            try:
                p.cerrar()
            except Exception:
                pass

    def _resolver_conflicto(self, conf):
        """Diálogo modal: muestra el valor del usuario vs. el ya guardado por otro
        proceso y deja decidir. Devuelve 'sobrescribir' o 'descartar' (default).
        Se pasa como `resolver` a la reimportación."""
        def _txt(x):
            desc = getattr(x, "concepto", None) or getattr(x, "descripcion", "")
            return (f"Fecha: {getattr(x,'fecha','')}\n"
                    f"Desc: {desc}\n"
                    f"Depósito: {getattr(x,'deposito',0)}   Retiro: {getattr(x,'retiro',0)}\n"
                    f"UUID: {getattr(x,'uuid_cfdi','') or '—'}")

        dlg = tk.Toplevel(self); dlg.title("Conflicto de datos"); dlg.transient(self)
        dlg.configure(bg=T.SURFACE); dlg.resizable(False, False); dlg.grab_set()
        eleccion = {"v": "descartar"}
        tk.Label(dlg, text="⚠ Otro proceso ya modificó este registro",
                 bg=T.SURFACE, fg=T.TEXT, font=("Segoe UI", 12, "bold")
                 ).pack(anchor="w", padx=16, pady=(14, 2))
        tk.Label(dlg, text=f"{conf.tabla}  ID={conf.registro_id}   "
                 f"(tu versión {conf.version_local} · guardada {conf.version_guardada})",
                 bg=T.SURFACE, fg=T.TEXT_MUTED, font=FONT_SMALL).pack(anchor="w", padx=16)
        cols = tk.Frame(dlg, bg=T.SURFACE); cols.pack(fill="x", padx=12, pady=10)
        for i, (titulo, val, color) in enumerate((
                ("TU CAMBIO (Excel)", conf.local, "#1565C0"),
                ("YA GUARDADO (otro)", conf.guardado, "#C62828"))):
            c = tk.Frame(cols, bg=T.SURFACE, highlightbackground=color,
                         highlightthickness=1); c.grid(row=0, column=i, padx=6, sticky="nsew")
            tk.Label(c, text=titulo, bg=color, fg="white", font=("Segoe UI", 9, "bold")
                     ).pack(fill="x")
            tk.Label(c, text=_txt(val), bg=T.SURFACE, fg=T.TEXT, justify="left",
                     font=("Consolas", 9)).pack(anchor="w", padx=8, pady=6)
        bf = tk.Frame(dlg, bg=T.SURFACE); bf.pack(fill="x", padx=16, pady=(0, 12))
        def _elige(v): eleccion["v"] = v; dlg.destroy()
        self._rbtn(bf, "Sobrescribir con lo mío", lambda: _elige("sobrescribir"),
                   "danger").pack(side="left")
        self._rbtn(bf, "Descartar mi cambio", lambda: _elige("descartar"),
                   "ghost").pack(side="left", padx=(8, 0))
        dlg.update_idletasks()
        try:
            px, py = self.winfo_rootx(), self.winfo_rooty()
            dlg.geometry(f"+{px + 100}+{py + 100}")
        except Exception:
            pass
        self.wait_window(dlg)
        return eleccion["v"]


    def _build_compliance_panel(self):
        f = self._cpanel
        tk.Label(f, text="Panel de Obligaciones Fiscales — Todas las Empresas",
                 bg=C_BG, fg=C_TEXT, font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=16, pady=(14, 2))

        # Selector de mes/año + botón recargar
        ctrl = tk.Frame(f, bg=C_BG); ctrl.pack(fill="x", padx=16, pady=(0, 8))
        tk.Label(ctrl, text="Mes:", bg=C_BG, fg=C_TEXT, font=FONT_HEAD).pack(side="left")
        self._cp_mes = ttk.Combobox(ctrl, values=CE.MESES_ES[1:], width=13, state="readonly")
        self._cp_mes.set(CE.MESES_ES[date.today().month]); self._cp_mes.pack(side="left", padx=(4, 12))
        tk.Label(ctrl, text="Año:", bg=C_BG, fg=C_TEXT, font=FONT_HEAD).pack(side="left")
        self._cp_anio = ttk.Combobox(ctrl, values=[str(y) for y in CE.anios_disponibles()],
                                      width=8, state="readonly")
        self._cp_anio.set(self._anio_default()); self._cp_anio.pack(side="left", padx=(4, 16))
        self._rbtn(ctrl, "🔄 Actualizar", self._reload_compliance_panel, "primary",
                   small=True).pack(side="left")
        self._rbtn(ctrl, "✉✉ Envío masivo", self._envio_masivo_compliance,
                   "warn", small=True).pack(side="left", padx=(12, 0))
        self._rbtn(ctrl, "📊 Generar reporte", self._reporte_panel, "primary",
                   small=True).pack(side="left", padx=(8, 0))

        # Tabla resumen
        cols = ("Empresa", "RFC", "Destinatario", "Oblig./mes", "Carpeta", "Archivos", "Estado")
        self._cp_tree = ttk.Treeview(f, columns=cols, show="headings", height=12)
        widths = [200, 130, 220, 80, 70, 80, 130]
        for c, w in zip(cols, widths):
            self._cp_tree.heading(c, text=c)
            self._cp_tree.column(c, width=w, anchor="w" if c in ("Empresa","Destinatario") else "center")
        self._cp_tree.tag_configure("ok",   foreground=C_GREEN)
        self._cp_tree.tag_configure("warn", foreground=C_WARN)
        self._cp_tree.tag_configure("err",  foreground=C_RED)
        sb = ttk.Scrollbar(f, orient="vertical", command=self._cp_tree.yview)
        self._cp_tree.configure(yscrollcommand=sb.set)
        self._cp_tree.pack(side="left", fill="both", expand=True, padx=(16, 0), pady=(0, 12))
        sb.pack(side="left", fill="y", pady=(0, 12))
        self._cp_tree.bind("<Double-1>", self._on_panel_dblclick)
        self._reload_compliance_panel()

    def _reload_compliance_panel(self):
        try:
            mes  = CE.MESES_ES.index(self._cp_mes.get())
            anio = int(self._cp_anio.get())
        except Exception:
            return
        for r in self._cp_tree.get_children(): self._cp_tree.delete(r)
        empresas = listar_empresas()
        hoy = date.today()
        limite = CE.dia_limite_habil(anio, mes, 27)
        for rfc, nombre in empresas:
            cfg = CE.obtener_config(rfc)
            obs = CE.obligaciones_del_mes(cfg, mes)
            arcs = CE.listar_archivos_mes(cfg, anio, mes)
            dest = cfg.get("destinatario", "")
            tiene_carp = "✅" if cfg.get("carpeta_raiz") else "⚠"
            n_arc = str(len(arcs)) if cfg.get("carpeta_raiz") else "–"
            if not dest:
                estado, tag = "⚠ Sin destinatario", "err"
            elif not cfg.get("carpeta_raiz"):
                estado, tag = "⚠ Sin carpeta", "warn"
            elif not obs:
                estado, tag = "⚠ Sin obligaciones", "warn"
            else:
                estado, tag = "✅ Listo", "ok"
            self._cp_tree.insert("", "end", tags=(tag,),
                values=(nombre, rfc, dest, str(len(obs)), tiene_carp, n_arc, estado))

    def _on_panel_dblclick(self, _):
        sel = self._cp_tree.selection()
        if not sel: return
        vals = self._cp_tree.item(sel[0])["values"]
        _, rfc = vals[0], vals[1]
        if hasattr(self, "_c_sel_rfc"):
            self._c_sel_rfc.set(rfc)
        try: self.nb.select(self._tab_compliance)
        except Exception: pass
        # Abrir ventana de envío directamente
        try:
            mes  = CE.MESES_ES.index(self._cp_mes.get())
            anio = int(self._cp_anio.get())
            nombre = buscar_por_rfc(rfc)
            if nombre:
                cfg = CE.obtener_config(rfc)
                _DialogEnvioCompliance(self, nombre, rfc, cfg, self._email_cfg,
                                       mes, anio, self._c_log)
        except Exception as e:
            messagebox.showerror("Error", str(e), parent=self)

    def _envio_masivo_compliance(self):
        try:
            mes  = CE.MESES_ES.index(self._cp_mes.get())
            anio = int(self._cp_anio.get())
        except Exception: return
        empresas = [(rfc, nombre) for rfc, nombre in listar_empresas()
                    if CE.obtener_config(rfc).get("destinatario")]
        if not empresas:
            messagebox.showwarning("Sin destinatarios",
                "Ninguna empresa tiene destinatario configurado en Compliance."); return
        if not messagebox.askyesno("Envío masivo",
            f"¿Abrir ventana de envío para {len(empresas)} empresa(s) de "
            f"{CE.MESES_ES[mes]} {anio}?"): return

        def abrir(idx):
            if idx >= len(empresas): return
            rfc, nombre = empresas[idx]
            cfg = CE.obtener_config(rfc)
            v = _DialogEnvioCompliance(self, nombre, rfc, cfg, self._email_cfg,
                                       mes, anio, self._c_log)
            v.bind("<Destroy>", lambda e, i=idx+1: self.after(300, abrir, i))
        abrir(0)

    def _build_compliance_obligs(self):
        f = self._cobligs
        tk.Label(f, text="Obligaciones activas por empresa",
                 bg=C_BG, fg=C_TEXT, font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=16, pady=(12, 4))

        # Selector empresa
        sel_f = tk.Frame(f, bg=C_BG); sel_f.pack(fill="x", padx=16, pady=(0, 6))
        tk.Label(sel_f, text="RFC Empresa:", bg=C_BG, fg=C_TEXT, font=FONT_HEAD).pack(side="left")
        self._c_sel_rfc = tk.StringVar()
        rfcs = [rfc for rfc, _ in listar_empresas()]
        self._cb_sel_rfc = ttk.Combobox(sel_f, textvariable=self._c_sel_rfc, values=rfcs,
                          width=22, state="readonly")
        if rfcs: self._cb_sel_rfc.set(rfcs[0])
        self._cb_sel_rfc.pack(side="left", padx=(8, 12))
        self._rbtn(sel_f, "Cargar", self._cargar_obligs, "primary", small=True
                   ).pack(side="left")
        # Periodo (contexto para marcar/ver declaraciones presentadas por mes)
        tk.Label(sel_f, text="Mes:", bg=C_BG, fg=C_TEXT, font=FONT_HEAD).pack(side="left", padx=(14, 2))
        self._co_mes = ttk.Combobox(sel_f, values=CE.MESES_ES[1:], width=11, state="readonly")
        self._co_mes.set(CE.MESES_ES[date.today().month]); self._co_mes.pack(side="left")
        self._co_mes.bind("<<ComboboxSelected>>", lambda _e: self._cargar_obligs())
        tk.Label(sel_f, text="Año:", bg=C_BG, fg=C_TEXT, font=FONT_HEAD).pack(side="left", padx=(8, 2))
        self._co_anio = ttk.Combobox(sel_f, values=[str(y) for y in CE.anios_disponibles()],
                                     width=7, state="readonly")
        self._co_anio.set(self._anio_default()); self._co_anio.pack(side="left")
        self._co_anio.bind("<<ComboboxSelected>>", lambda _e: self._cargar_obligs())
        self._lbl_empresa_obligs = tk.Label(sel_f, text="", bg=C_BG, fg=C_GRAY2, font=FONT_SMALL)
        self._lbl_empresa_obligs.pack(side="left", padx=12)

        # Botones acción
        bf = tk.Frame(f, bg=C_BG); bf.pack(fill="x", padx=16, pady=(0, 6))
        tk.Label(bf, text="Para CARGAR declaraciones usa la pestaña «Archivos».",
                 bg=C_BG, fg=C_GRAY2, font=FONT_SMALL).pack(side="left", padx=(0, 12))
        self._rbtn(bf, "✓ Todas", lambda: self._sel_obligs(True), "accent", small=True
                   ).pack(side="left", padx=(0, 6))
        self._rbtn(bf, "✗ Ninguna", lambda: self._sel_obligs(False), "ghost", small=True
                   ).pack(side="left", padx=(0, 6))
        self._rbtn(bf, "💾 Guardar selección", self._guardar_obligs, "primary", small=True
                   ).pack(side="left")
        self._rbtn(bf, "📊 Generar reporte", self._reporte_obligaciones, "primary",
                   small=True).pack(side="left", padx=(8, 0))

        # Tabla con Treeview — más fiable que canvas-frame en tkinter
        tree_cont = tk.Frame(f, bg=C_BG); tree_cont.pack(fill="both", expand=True, padx=16, pady=(0, 8))
        cols = ("activa", "nombre", "periodicidad", "meses", "presentada")
        self._tv_obligs = ttk.Treeview(tree_cont, columns=cols, show="headings",
                                       height=12, selectmode="none")
        self._tv_obligs.heading("activa",       text="✔")
        self._tv_obligs.heading("nombre",        text="Obligación")
        self._tv_obligs.heading("periodicidad",  text="Periodicidad")
        self._tv_obligs.heading("meses",         text="Meses en que aplica")
        self._tv_obligs.heading("presentada",    text="Presentada (periodo)")
        self._tv_obligs.column("activa",      width=36,  minwidth=36,  stretch=False, anchor="center")
        self._tv_obligs.column("nombre",      width=340, minwidth=200, stretch=True,  anchor="w")
        self._tv_obligs.column("periodicidad",width=100, minwidth=80,  stretch=False, anchor="center")
        self._tv_obligs.column("meses",       width=210, minwidth=140, stretch=True,  anchor="w")
        self._tv_obligs.column("presentada",  width=230, minwidth=150, stretch=True,  anchor="w")
        vsb = ttk.Scrollbar(tree_cont, orient="vertical",   command=self._tv_obligs.yview)
        self._tv_obligs.configure(yscrollcommand=vsb.set)
        self._tv_obligs.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self._tv_obligs.bind("<Button-1>", self._toggle_oblig)

        # Colores alternos
        self._tv_obligs.tag_configure("par",  background="#F8FAFF")
        self._tv_obligs.tag_configure("impar", background=C_WHITE)

        self._oblig_vars: dict[str, bool] = {}  # id → activa (True/False)

        # Cargar catálogo completo después de que la ventana esté renderizada
        self.after(200, self._cargar_obligs)

    def _toggle_oblig(self, event):
        """Alterna activa/inactiva al hacer clic en cualquier fila del treeview."""
        region = self._tv_obligs.identify_region(event.x, event.y)
        if region not in ("cell", "tree"):
            return
        iid = self._tv_obligs.identify_row(event.y)
        if not iid:
            return
        oid = iid  # usamos el id de obligación como iid de la fila
        if oid not in self._oblig_vars:
            return
        nuevo = not self._oblig_vars[oid]
        self._oblig_vars[oid] = nuevo
        marca = "✔" if nuevo else ""
        self._tv_obligs.set(iid, "activa", marca)

    def _cargar_obligs(self):
        rfc    = self._c_sel_rfc.get().strip().upper()
        nombre = buscar_por_rfc(rfc) if rfc else None
        if nombre:
            self._lbl_empresa_obligs.config(text=f"  {nombre}", fg=C_TEXT)
            cfg     = CE.obtener_config(rfc)
            activas = set(cfg.get("obligaciones_activas", CE.TODOS_IDS))
        else:
            if rfc:
                self._lbl_empresa_obligs.config(text="  RFC no registrado", fg=C_RED)
            else:
                self._lbl_empresa_obligs.config(text="  (selecciona empresa)", fg=C_GRAY2)
            activas = set(CE.TODOS_IDS)  # mostrar todo activo por defecto

        # Presentaciones del periodo elegido (Mes/Año). Antes de leer, AUTO-SINCRONIZA
        # con el archivo en disco: si el PDF de una declaración ya no está, se
        # desmarca sola y pasa a la bitácora (conserva fecha/operación).
        presentadas = {}
        try:
            _mes  = CE.MESES_ES.index(self._co_mes.get())
            _anio = int(self._co_anio.get())
            if rfc and nombre:
                CE.sincronizar_presentaciones(rfc, _anio, _mes)
                presentadas = CE.obtener_presentaciones(rfc, _anio, _mes)
        except Exception:
            presentadas = {}

        # Limpiar y reconstruir treeview
        for item in self._tv_obligs.get_children():
            self._tv_obligs.delete(item)
        self._oblig_vars.clear()
        self._tv_obligs.tag_configure("presentada", background="#E8F5E9")

        for i, o in enumerate(CE.OBLIGACIONES_CATALOGO):
            es_activa = o["id"] in activas
            self._oblig_vars[o["id"]] = es_activa
            marca     = "✔" if es_activa else ""
            meses_str = "  " + ",  ".join(CE.MESES_ES[m][:3] for m in o["meses_aplica"])
            pr = presentadas.get(o["id"])
            if pr:
                pres_str = f"✅ {pr.get('fecha_presentacion','')}  ·  op {pr.get('numero_operacion','')}"
                tag = "presentada"
            else:
                pres_str = ""
                tag = "par" if i % 2 == 0 else "impar"
            self._tv_obligs.insert("", "end", iid=o["id"], tags=(tag,),
                                   values=(marca, f"  {o['nombre']}", o["periodicidad"],
                                           meses_str, pres_str))

    def _sel_obligs(self, v: bool):
        for oid in list(self._oblig_vars.keys()):
            self._oblig_vars[oid] = v
            self._tv_obligs.set(oid, "activa", "✔" if v else "")

    def _guardar_obligs(self):
        rfc = self._c_sel_rfc.get().strip().upper()
        if not rfc:
            messagebox.showwarning("Sin empresa", "Selecciona primero una empresa."); return
        if not buscar_por_rfc(rfc):
            messagebox.showwarning("RFC no registrado",
                "Registra la empresa primero en el catálogo de empresas."); return
        cfg = CE.obtener_config(rfc)
        cfg["obligaciones_activas"] = [oid for oid, activa in self._oblig_vars.items() if activa]
        CE.guardar_config(rfc, cfg)
        messagebox.showinfo("Guardado", "✅ Obligaciones guardadas.")
        self._reload_compliance_panel()

    def _build_compliance_archivos(self):
        f = self._carch
        tk.Label(f, text="Carga y explorador de archivos por empresa y mes",
                 bg=C_BG, fg=C_TEXT, font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=16, pady=(12, 2))

        # ── Punto ÚNICO de carga: lee los PDF del SAT, marca las obligaciones
        #    (pestaña Obligaciones) y alimenta Contribuciones/Saldos, todo aquí. ──
        carga = tk.Frame(f, bg=C_BG); carga.pack(fill="x", padx=16, pady=(0, 4))
        self._rbtn(carga, "📤 Cargar declaración(es)", self._cargar_declaracion,
                   "accent", small=True).pack(side="left", padx=(0, 6))
        self._rbtn(carga, "📁 Cargar carpeta", self._cargar_declaraciones_carpeta,
                   "primary", small=True).pack(side="left")
        tk.Label(carga, text="  Lee los PDF, marca Obligaciones y alimenta "
                 "Contribuciones/Saldos automáticamente.",
                 bg=C_BG, fg=C_GRAY2, font=FONT_SMALL).pack(side="left", padx=8)

        ctrl = tk.Frame(f, bg=C_BG); ctrl.pack(fill="x", padx=16, pady=(0, 8))
        tk.Label(ctrl, text="RFC:", bg=C_BG, fg=C_TEXT, font=FONT_HEAD).pack(side="left")
        self._ca_rfc = tk.StringVar()
        rfcs = [rfc for rfc, _ in listar_empresas()]
        self._ca_combo = ttk.Combobox(ctrl, textvariable=self._ca_rfc, values=rfcs,
                                      width=18, state="readonly")
        if rfcs: self._ca_combo.set(rfcs[0])
        self._ca_combo.pack(side="left", padx=(4, 12))
        tk.Label(ctrl, text="Mes:", bg=C_BG, fg=C_TEXT, font=FONT_HEAD).pack(side="left")
        self._ca_mes = ttk.Combobox(ctrl, values=CE.MESES_ES[1:], width=13, state="readonly")
        self._ca_mes.set(CE.MESES_ES[date.today().month]); self._ca_mes.pack(side="left", padx=4)
        tk.Label(ctrl, text="Año:", bg=C_BG, fg=C_TEXT, font=FONT_HEAD).pack(side="left")
        self._ca_anio = ttk.Combobox(ctrl, values=[str(y) for y in CE.anios_disponibles()],
                                      width=8, state="readonly")
        self._ca_anio.set(self._anio_default()); self._ca_anio.pack(side="left", padx=4)
        self._rbtn(ctrl, "🔍 Ver", self._refrescar_archivos, "primary", small=True
                   ).pack(side="left", padx=(12, 0))
        self._rbtn(ctrl, "📂 Abrir carpeta", self._abrir_carpeta, "ghost", small=True
                   ).pack(side="left", padx=(8, 0))
        self._rbtn(ctrl, "📊 Generar reporte", self._reporte_archivos, "primary",
                   small=True).pack(side="left", padx=(8, 0))
        self._ca_lbl_ruta = tk.Label(ctrl, text="", bg=C_BG, fg=C_GRAY2, font=FONT_SMALL)
        self._ca_lbl_ruta.pack(side="left", padx=12)

        cols = ("Archivo", "Tamaño", "Fecha modificación")
        self._ca_tree = ttk.Treeview(f, columns=cols, show="headings", height=14)
        for c, w in zip(cols, [450, 100, 180]):
            self._ca_tree.heading(c, text=c)
            self._ca_tree.column(c, width=w, anchor="w" if c == "Archivo" else "center")
        sb = ttk.Scrollbar(f, orient="vertical", command=self._ca_tree.yview)
        self._ca_tree.configure(yscrollcommand=sb.set)
        self._ca_tree.pack(side="left", fill="both", expand=True, padx=(16, 0), pady=(0, 12))
        sb.pack(side="left", fill="y", pady=(0, 12))
        if rfcs: self._refrescar_archivos()

    def _refrescar_archivos(self):
        rfc = self._ca_rfc.get().strip().upper()
        try:
            mes  = CE.MESES_ES.index(self._ca_mes.get())
            anio = int(self._ca_anio.get())
        except Exception: return
        for r in self._ca_tree.get_children(): self._ca_tree.delete(r)
        cfg  = CE.obtener_config(rfc)
        ruta = CE.carpeta_mes(cfg, anio, mes)
        if ruta:
            self._ca_lbl_ruta.config(text=ruta, fg=C_GRAY2)
        else:
            esperada = os.path.join(cfg.get("carpeta_raiz", ""), f"{anio:04d}-{mes:02d}")
            self._ca_lbl_ruta.config(text=f"⚠ No existe: {esperada}", fg=C_RED)
            return
        iconos = {".pdf":"📄",".xml":"📋",".xlsx":"📊",".xls":"📊",
                  ".zip":"🗜",".rar":"🗜",".sua":"💾",".txt":"📝"}
        for a in CE.listar_archivos_mes(cfg, anio, mes):
            ext = os.path.splitext(a["nombre"])[1].lower()
            self._ca_tree.insert("", "end",
                values=(f"{iconos.get(ext,'📎')}  {a['nombre']}",
                        CE.formato_tamano(a["tamano"]), a["mtime"]))

    def _abrir_carpeta(self):
        rfc = self._ca_rfc.get().strip().upper()
        try:
            mes  = CE.MESES_ES.index(self._ca_mes.get())
            anio = int(self._ca_anio.get())
        except Exception: return
        cfg  = CE.obtener_config(rfc)
        ruta = CE.carpeta_mes(cfg, anio, mes)
        if ruta: os.startfile(ruta)
        else:
            raiz = cfg.get("carpeta_raiz", "")
            if raiz and os.path.isdir(raiz): os.startfile(raiz)
            else: messagebox.showinfo("Carpeta no encontrada", "Verifica la configuración.", parent=self)

    def _build_compliance_config(self):
        f = self._cconfig
        tk.Label(f, text="Configuración de Compliance por Empresa",
                 bg=C_BG, fg=C_TEXT, font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=20, pady=(14, 6))

        # Selector empresa
        sel_f = tk.Frame(f, bg=C_BG); sel_f.pack(fill="x", padx=20, pady=(0, 10))
        tk.Label(sel_f, text="RFC:", bg=C_BG, fg=C_TEXT, font=FONT_HEAD).pack(side="left")
        self._cc_rfc = tk.StringVar()
        rfcs = [rfc for rfc, _ in listar_empresas()]
        self._cc_combo = ttk.Combobox(sel_f, textvariable=self._cc_rfc, values=rfcs,
                                       width=20, state="readonly")
        if rfcs: self._cc_combo.set(rfcs[0])
        self._cc_combo.pack(side="left", padx=(8, 12))
        self._rbtn(sel_f, "Cargar", self._cargar_cc, "primary", small=True).pack(side="left")
        self._rbtn(sel_f, "📄 Subir CSF / CIF", self._subir_cif_config_empresa,
                   "primary", small=True).pack(side="left", padx=(8, 0))
        if _SAT_PORTAL_OK:
            self._rbtn(sel_f, "🔑 API key (Portal SAT)", self._config_api_key_anthropic,
                       "primary", small=True).pack(side="left", padx=(8, 0))
        self._cc_lbl = tk.Label(sel_f, text="", bg=C_BG, fg=C_GRAY2, font=FONT_SMALL)
        self._cc_lbl.pack(side="left", padx=10)

        # Régimen fiscal de la empresa seleccionada (se llena vía CSF)
        reg_f = tk.Frame(f, bg=C_BG); reg_f.pack(fill="x", padx=20, pady=(0, 10))
        self._cc_reg_lbl = tk.Label(reg_f, text="", bg=C_BG, fg=C_GRAY2,
                                     font=FONT_SMALL, anchor="w", justify="left", wraplength=560)
        self._cc_reg_lbl.pack(anchor="w")

        card = tk.Frame(f, bg=C_WHITE); card.pack(padx=20, fill="x", pady=(0, 10))
        fields = [
            ("Correo destinatario (para:)", "destinatario"),
            ("CC (separados por coma):",    "cc"),
            ("Plantilla asunto:",           "asunto_template"),
        ]
        self._cc_vars: dict[str, tk.StringVar] = {}
        for i, (lb, key) in enumerate(fields):
            tk.Label(card, text=lb, bg=C_WHITE, fg=C_GRAY2, font=FONT_NORMAL).grid(
                row=i, column=0, sticky="w", padx=16, pady=(10, 0))
            var = tk.StringVar()
            ef  = tk.Frame(card, bg=C_GRAY)
            ef.grid(row=i, column=1, sticky="ew", padx=16, pady=(10, 0))
            tk.Entry(ef, textvariable=var, font=("Segoe UI", 10), width=58,
                     bg=C_GRAY, fg=C_TEXT, relief="flat", bd=4,
                     insertbackground=C_ACCENT).pack(fill="x")
            self._cc_vars[key] = var

        # Carpeta raíz
        r = len(fields)
        tk.Label(card, text="Carpeta raíz de archivos:", bg=C_WHITE, fg=C_GRAY2,
                 font=FONT_NORMAL).grid(row=r, column=0, sticky="w", padx=16, pady=(10, 0))
        cf = tk.Frame(card, bg=C_WHITE)
        cf.grid(row=r, column=1, sticky="ew", padx=16, pady=(10, 0))
        self._cc_carpeta = tk.StringVar()
        tk.Entry(cf, textvariable=self._cc_carpeta, font=("Segoe UI", 10), width=46,
                 bg=C_GRAY, fg=C_TEXT, relief="flat", bd=4).pack(side="left", fill="x", expand=True)
        self._rbtn(cf, "📂 Buscar…", self._buscar_carpeta_cc, "ghost", small=True
                   ).pack(side="left", padx=(6, 0))

        # Notas
        tk.Label(card, text="Notas adicionales:", bg=C_WHITE, fg=C_GRAY2,
                 font=FONT_NORMAL).grid(row=r+1, column=0, sticky="nw", padx=16, pady=(10, 10))
        self._cc_notas = tk.Text(card, width=58, height=2, font=("Segoe UI", 9),
                                  bg=C_GRAY, fg=C_TEXT, relief="flat", bd=4,
                                  insertbackground=C_ACCENT)
        self._cc_notas.grid(row=r+1, column=1, padx=16, pady=(10, 10), sticky="ew")
        card.columnconfigure(1, weight=1)

        tk.Label(card, text="Variables asunto: {empresa}  {mes}  {anio}  {dia}",
                 bg=C_WHITE, fg=C_GRAY2, font=FONT_SMALL).grid(
            row=r+2, column=1, sticky="w", padx=16, pady=(0, 10))

        bf = tk.Frame(f, bg=C_BG); bf.pack(fill="x", padx=20, pady=(0, 10))
        self._rbtn(bf, "📅 Años de trabajo", self._abrir_anios_trabajo, "primary"
                   ).pack(side="left", padx=(0, 8))
        self._rbtn(bf, "💾 Guardar configuración", self._guardar_cc, "accent"
                   ).pack(side="left", padx=(0, 12))
        self._rbtn(bf, "📨 Preparar envío", self._preparar_envio_cc, "primary"
                   ).pack(side="left")

        # Log compliance
        tk.Label(f, text="Historial de envíos:", bg=C_BG, fg=C_TEXT,
                 font=FONT_HEAD).pack(anchor="w", padx=20, pady=(10, 2))
        self._c_log_text = scrolledtext.ScrolledText(
            f, height=4, wrap="word", font=FONT_MONO,
            bg="#1A1A19", fg="#C3C2B7", relief="flat", bd=0)
        self._c_log_text.pack(fill="x", padx=20, pady=(0, 12))

        if rfcs: self._cargar_cc()

    def _c_log(self, msg: str):
        ts = datetime.now().strftime("%H:%M:%S")
        self._c_log_text.config(state="normal")
        self._c_log_text.insert("end", f"[{ts}] {msg}\n")
        self._c_log_text.see("end")
        self._c_log_text.config(state="disabled")

    def _cargar_cc(self):
        rfc    = self._cc_rfc.get().strip().upper()
        nombre = buscar_por_rfc(rfc) if rfc else None
        self._cc_lbl.config(text=nombre or "RFC no registrado")
        self._refrescar_reg_cc(rfc)
        if not nombre: return
        cfg = CE.obtener_config(rfc)
        for key, var in self._cc_vars.items():
            var.set(cfg.get(key, ""))
        self._cc_carpeta.set(cfg.get("carpeta_raiz", ""))
        self._cc_notas.delete("1.0", "end")
        self._cc_notas.insert("1.0", cfg.get("notas_adicionales", ""))

    def _refrescar_reg_cc(self, rfc: str):
        """Actualiza la etiqueta de régimen fiscal (CSF) de la empresa seleccionada."""
        emp = obtener_empresa(rfc) if rfc else None
        if not emp:
            self._cc_reg_lbl.config(text="", fg=C_GRAY2)
            return
        codigos = emp.get("regimenes", [])
        nombres = emp.get("regimenes_nombres", [])
        if codigos and nombres:
            partes = [f"{n} ({c})" for n, c in zip(nombres, codigos)]
            self._cc_reg_lbl.config(text="✔  Régimen(es): " + ",  ".join(partes), fg="#2E7D32")
        elif codigos:
            self._cc_reg_lbl.config(text="Régimen(es): " + ", ".join(codigos), fg="#2E7D32")
        else:
            self._cc_reg_lbl.config(
                text="Sin CSF registrada — sube la Constancia de Situación Fiscal",
                fg="#E65100")

    def _subir_cif_config_empresa(self):
        """
        Sube la Constancia de Situación Fiscal (CSF/CIF) de una empresa.
        Único lugar del sistema para esta acción — antes existían botones
        duplicados en Convertidor PDF→Excel y Conciliación CFDI.
        """
        pdf = filedialog.askopenfilename(
            title="Seleccionar Constancia de Situación Fiscal (PDF)",
            filetypes=[("PDF", "*.pdf"), ("Todos", "*.*")],
        )
        if not pdf:
            return
        try:
            ok, msg, datos = registrar_desde_cif(pdf)
            if ok and datos:
                rfcs = [rfc for rfc, _ in listar_empresas()]
                self._cc_combo["values"] = rfcs
                self._cc_rfc.set(datos.rfc)
                self._cargar_cc()
                messagebox.showinfo("CSF registrada", msg)
            else:
                messagebox.showerror("Error al leer CSF", msg)
        except Exception as e:
            messagebox.showerror("Error al procesar CSF", str(e))

    def _config_api_key_anthropic(self):
        """
        Guarda/borra la API key de Anthropic usada SOLO para la
        auto-reparación de selectores del canal Portal SAT
        (sat_descarga_portal.py). Se cifra en disco con el mismo esquema
        XOR-SHA256 atado al MAC que usa repo_conciliaciones.py — nunca se
        guarda en texto plano.
        """
        actual = _SAT_PORTAL.hay_api_key_configurada()
        estado = "configurada" if actual else "no configurada"
        nueva = simpledialog.askstring(
            "API key de Anthropic (Portal SAT)",
            f"Estado actual: {estado}.\n\n"
            "Esta key solo se usa cuando el robot del Portal SAT no reconoce "
            "un campo de la página (cambio de layout) — le pide a la API de "
            "Anthropic un selector nuevo. Sin key, el canal Portal sigue "
            "funcionando igual; solo se pierde la auto-reparación.\n\n"
            "Deja vacío y acepta para borrar la key guardada. Cancela para no cambiar nada.",
            show="*", parent=self,
        )
        if nueva is None:
            return
        if nueva.strip():
            _SAT_PORTAL.guardar_api_key_anthropic(nueva.strip())
            messagebox.showinfo("Guardado", "API key guardada (cifrada en disco).", parent=self)
        else:
            _SAT_PORTAL.eliminar_api_key_anthropic()
            messagebox.showinfo("Borrado", "API key eliminada.", parent=self)

    def _buscar_carpeta_cc(self):
        d = filedialog.askdirectory(title="Carpeta raíz de archivos de la empresa")
        if d: self._cc_carpeta.set(d)

    def _guardar_cc(self):
        rfc = self._cc_rfc.get().strip().upper()
        if not rfc: return
        cfg = CE.obtener_config(rfc)
        for key, var in self._cc_vars.items():
            cfg[key] = var.get()
        cfg["carpeta_raiz"]      = self._cc_carpeta.get()
        cfg["notas_adicionales"] = self._cc_notas.get("1.0", "end-1c")
        CE.guardar_config(rfc, cfg)
        messagebox.showinfo("Guardado", "✅ Configuración compliance guardada.")
        self._reload_compliance_panel()

    def _preparar_envio_cc(self):
        rfc    = self._cc_rfc.get().strip().upper()
        nombre = buscar_por_rfc(rfc) if rfc else None
        if not nombre:
            messagebox.showwarning("Sin empresa", "Selecciona y carga una empresa.", parent=self); return
        cfg = CE.obtener_config(rfc)
        try:
            mes  = CE.MESES_ES.index(self._cp_mes.get()) if hasattr(self, "_cp_mes") else date.today().month
            anio = int(self._cp_anio.get()) if hasattr(self, "_cp_anio") else date.today().year
        except Exception:
            mes, anio = date.today().month, date.today().year
        _DialogEnvioCompliance(self, nombre, rfc, cfg, self._email_cfg,
                               mes, anio, self._c_log)

    # ══════════════════════════════════════════════════════════════════════════
    #  TAB: CONFIGURACIÓN GENERAL
    # ══════════════════════════════════════════════════════════════════════════
    def _build_config_tab(self):
        f = self._tab_config

        # ── Cabecera de módulo (blanca, estilo prototipo) ────────────────────
        self._module_header(
            f, "⚙", "Configuración",
            "Correo saliente, usuarios y base de datos")

        subnb = ttk.Notebook(f); subnb.pack(fill="both", expand=True)

        self._cfg_email_tab  = tk.Frame(subnb, bg=C_BG)
        self._cfg_users_tab  = tk.Frame(subnb, bg=C_BG)
        self._cfg_db_tab     = tk.Frame(subnb, bg=C_BG)

        subnb.add(self._cfg_email_tab, text="  📧  Email Multi-proveedor  ")
        subnb.add(self._cfg_users_tab, text="  👤  Usuarios  ")
        subnb.add(self._cfg_db_tab,    text="  🗄  Base de datos  ")

        self._build_email_config()
        self._build_users_config()
        self._build_db_config()

    # ── Base de datos (Firebird) ──────────────────────────────────────────────
    def _build_db_config(self):
        f = self._cfg_db_tab
        try:
            import db_config as _dc
            cfg = _dc.cargar()
        except Exception:
            cfg = {"motor": "firebird", "host": "localhost", "port": 3050,
                   "database": "", "user": "SYSDBA", "password": "",
                   "charset": "UTF8", "fbclient": ""}

        tk.Label(f, text="Conexión a la base de datos (Firebird)",
                 bg=C_BG, fg=C_TEXT, font=("Segoe UI", 12, "bold")
                 ).pack(anchor="w", padx=20, pady=(14, 2))
        tk.Label(f, text="Parámetros estándar de conexión Firebird. La base guarda "
                        "la conciliación y sus CFDI (para reimportar y generar pólizas).",
                 bg=C_BG, fg=C_GRAY2, font=FONT_SMALL).pack(anchor="w", padx=20, pady=(0, 10))

        card = tk.Frame(f, bg=C_WHITE, highlightbackground=T.BORDER, highlightthickness=1)
        card.pack(fill="x", padx=20)
        card.columnconfigure(1, weight=1)
        self._db_vars = {}

        def _fila(r, etq, clave, ancho=44, show=None, hint=""):
            tk.Label(card, text=etq, bg=C_WHITE, fg=C_TEXT, font=FONT_HEAD
                     ).grid(row=r, column=0, sticky="w", padx=(14, 8), pady=(10 if r == 0 else 6, 2))
            var = tk.StringVar(value=str(cfg.get(clave, "")))
            self._db_vars[clave] = var
            ent = tk.Entry(card, textvariable=var, font=FONT_NORMAL, width=ancho,
                           bg=C_GRAY, fg=C_TEXT, relief="flat", bd=4,
                           show=(show or ""))
            ent.grid(row=r, column=1, sticky="ew", pady=(10 if r == 0 else 6, 2))
            if hint:
                tk.Label(card, text=hint, bg=C_WHITE, fg=C_GRAY2, font=("Segoe UI", 8)
                         ).grid(row=r, column=2, sticky="w", padx=(8, 12))
            return ent

        _fila(0, "Servidor (host):", "host", hint="localhost si es local")
        _fila(1, "Puerto:",          "port", ancho=10, hint="3050 (default de Firebird)")
        _fila(2, "Ruta de la BD (.fdb):", "database")
        _bwrap = tk.Frame(card, bg=C_WHITE); _bwrap.grid(row=2, column=2, sticky="w", padx=(8, 12))
        self._rbtn(_bwrap, "📂", self._db_buscar_fdb, "ghost", small=True).pack()
        _fila(3, "Usuario:",   "user", ancho=20, hint="SYSDBA por defecto")
        self._db_pwd_entry = _fila(4, "Contraseña:", "password", ancho=20, show="•")
        _fila(5, "Charset:",   "charset", ancho=12, hint="UTF8")
        _fila(6, "fbclient.dll:", "fbclient",
              hint="vacío = autodetección (64-bit)")
        _fw = tk.Frame(card, bg=C_WHITE); _fw.grid(row=6, column=2, sticky="w", padx=(8, 12))
        self._rbtn(_fw, "📂", self._db_buscar_fbclient, "ghost", small=True).pack()

        _pf = tk.Frame(card, bg=C_WHITE); _pf.grid(row=7, column=1, sticky="w", pady=(2, 10))
        self._rbtn(_pf, "👁 Mostrar/ocultar contraseña", self._db_toggle_pwd,
                   "ghost", small=True).pack(side="left")

        bf = tk.Frame(f, bg=C_BG); bf.pack(fill="x", padx=20, pady=(12, 6))
        self._rbtn(bf, "🔌 Probar conexión", self._db_probar, "primary"
                   ).pack(side="left")
        self._rbtn(bf, "💾 Guardar configuración", self._db_guardar, "accent"
                   ).pack(side="left", padx=(10, 0))
        self._db_lbl_estado = tk.Label(bf, text="", bg=C_BG, font=("Segoe UI", 9))
        self._db_lbl_estado.pack(side="left", padx=16)

        tk.Label(f, text="Nota: si Firebird no está instalado o el servicio no "
                        "corre, la conciliación funciona igual pero SIN guardar en "
                        "la BD (no habrá IDs ni reimportación).",
                 bg=C_BG, fg=C_GRAY2, font=FONT_SMALL, wraplength=680, justify="left"
                 ).pack(anchor="w", padx=20, pady=(8, 0))

    def _db_buscar_fdb(self):
        ruta = filedialog.asksaveasfilename(
            title="Ruta del archivo de base de datos Firebird",
            defaultextension=".fdb", filetypes=[("Firebird", "*.fdb"), ("Todos", "*.*")])
        if ruta:
            self._db_vars["database"].set(ruta)

    def _db_buscar_fbclient(self):
        ruta = filedialog.askopenfilename(
            title="Selecciona fbclient.dll (64-bit)",
            filetypes=[("DLL", "fbclient.dll"), ("Todos", "*.*")])
        if ruta:
            self._db_vars["fbclient"].set(ruta)

    def _db_toggle_pwd(self):
        e = self._db_pwd_entry
        e.config(show="" if e.cget("show") else "•")

    def _db_cfg_actual(self) -> dict:
        return {k: v.get().strip() for k, v in self._db_vars.items()}

    def _db_probar(self):
        import db_config as _dc
        self._db_lbl_estado.config(text="Probando…", fg=C_BLUE)
        self.update_idletasks()
        ok, msg = _dc.probar_conexion(self._db_cfg_actual())
        self._db_lbl_estado.config(text=("✔ Conectado" if ok else "✕ Error"),
                                   fg=("#2E7D32" if ok else C_RED))
        (messagebox.showinfo if ok else messagebox.showerror)(
            "Prueba de conexión", msg, parent=self)

    def _db_guardar(self):
        import db_config as _dc
        try:
            _dc.guardar(self._db_cfg_actual())
            self._db_lbl_estado.config(text="✔ Guardado", fg="#2E7D32")
            messagebox.showinfo("Configuración guardada",
                f"Se guardó en:\n{_dc.RUTA_CONFIG}", parent=self)
        except Exception as e:
            messagebox.showerror("Error al guardar", str(e), parent=self)

    # ── Email multi-proveedor ─────────────────────────────────────────────────
    def _build_email_config(self):
        f = self._cfg_email_tab
        tk.Label(f, text="Configuración de correo saliente",
                 bg=C_BG, fg=C_TEXT, font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=20, pady=(14, 6))

        # Tarjetas de proveedor
        prov_f = tk.Frame(f, bg=C_BG); prov_f.pack(fill="x", padx=20, pady=(0, 12))
        tk.Label(prov_f, text="Proveedor de correo:", bg=C_BG, fg=C_TEXT,
                 font=FONT_HEAD).pack(anchor="w", pady=(0, 6))

        self._email_prov_var = tk.StringVar(value=self._email_cfg.get("proveedor", "gmail"))
        cards_f = tk.Frame(prov_f, bg=C_BG); cards_f.pack(fill="x")

        prov_colors = {
            "gmail":   ("#EA4335", "Gmail"),
            "outlook": ("#0078D4", "Outlook / Hotmail"),
            "yahoo":   ("#6001D2", "Yahoo Mail"),
            "smtp":    ("#374151", "SMTP Custom"),
        }
        self._prov_cards: dict[str, tk.Frame] = {}
        for i, (key, (color, label)) in enumerate(prov_colors.items()):
            card = tk.Frame(cards_f, bg=color if self._email_prov_var.get()==key else C_GRAY,
                            padx=14, pady=10, cursor="hand2")
            card.grid(row=0, column=i, padx=(0, 8), sticky="ew")
            tk.Label(card, text=label,
                     bg=color if self._email_prov_var.get()==key else C_GRAY,
                     fg=C_WHITE if self._email_prov_var.get()==key else C_TEXT,
                     font=("Segoe UI", 10, "bold")).pack()
            tk.Radiobutton(card, value=key, variable=self._email_prov_var,
                           bg=color if self._email_prov_var.get()==key else C_GRAY,
                           activebackground=color, cursor="hand2",
                           command=lambda k=key, c=color: self._on_prov_change(k, c)
                           ).pack()
            self._prov_cards[key] = (card, color)
            for w in [card] + list(card.winfo_children()):
                w.bind("<Button-1>", lambda e, k=key, c=color:
                    (self._email_prov_var.set(k), self._on_prov_change(k, c)))
        for i in range(4): cards_f.columnconfigure(i, weight=1, uniform="p")

        # Nota del proveedor
        self._lbl_prov_nota = tk.Label(f, text="", bg="#EFF6FF", fg="#1E40AF",
                                        font=FONT_SMALL, wraplength=600, justify="left",
                                        padx=16, pady=8)
        self._lbl_prov_nota.pack(fill="x", padx=20, pady=(0, 10))

        # Campos
        card = tk.Frame(f, bg=C_WHITE); card.pack(padx=20, fill="x", pady=(0, 10))
        self._email_vars: dict[str, tk.StringVar] = {}
        rows = [
            ("Correo remitente:",     "remitente", False),
            ("Contraseña / App pwd:", "password",  True),
            ("Host SMTP:",            "host_smtp", False),
            ("Puerto:",               "puerto",    False),
        ]
        self._pwd_entry_ref = None
        for i, (lb, key, is_pwd) in enumerate(rows):
            tk.Label(card, text=lb, bg=C_WHITE, fg=C_GRAY2, font=FONT_NORMAL).grid(
                row=i, column=0, sticky="w", padx=16, pady=(10, 0))
            var = tk.StringVar(value=str(self._email_cfg.get(key, "")))
            ef  = tk.Frame(card, bg=C_GRAY)
            ef.grid(row=i, column=1, sticky="ew", padx=16, pady=(10, 0))
            ent = tk.Entry(ef, textvariable=var, show="*" if is_pwd else "",
                           font=("Segoe UI", 10), width=50,
                           bg=C_GRAY, fg=C_TEXT, relief="flat", bd=4,
                           insertbackground=C_ACCENT)
            ent.pack(fill="x")
            self._email_vars[key] = var
            if is_pwd: self._pwd_entry_ref = ent

        self._ssl_var = tk.BooleanVar(value=self._email_cfg.get("ssl", False))
        tk.Checkbutton(card, text="Usar SSL directo (puerto 465) en lugar de STARTTLS (587)",
                       variable=self._ssl_var, bg=C_WHITE, fg=C_TEXT,
                       font=FONT_NORMAL, cursor="hand2").grid(
            row=len(rows), column=1, sticky="w", padx=16, pady=(10, 10))
        card.columnconfigure(1, weight=1)

        # Auto-rellenar host al cambiar remitente
        self._email_vars["remitente"].trace_add("write", self._on_remitente_change)

        # Toggle pwd
        self._rbtn(card, "👁 Mostrar/ocultar contraseña", self._toggle_email_pwd,
                   "ghost", small=True).grid(
            row=len(rows)+1, column=1, sticky="w", padx=16, pady=(0, 10))

        bf = tk.Frame(f, bg=C_BG); bf.pack(fill="x", padx=20, pady=(0, 10))
        self._rbtn(bf, "💾 Guardar", self._guardar_email, "accent"
                   ).pack(side="left", padx=(0, 12))
        self._rbtn(bf, "🧪 Probar conexión", self._probar_email, "primary"
                   ).pack(side="left")
        self._lbl_email_status = tk.Label(bf, text="", bg=C_BG, font=("Segoe UI", 10))
        self._lbl_email_status.pack(side="left", padx=16)

        self._on_prov_change(self._email_prov_var.get(),
                             prov_colors.get(self._email_prov_var.get(), ("", ""))[0])

    def _on_prov_change(self, proveedor, color):
        prov_colors = {"gmail":"#EA4335","outlook":"#0078D4","yahoo":"#6001D2","smtp":"#374151"}
        for key, (card, col) in self._prov_cards.items():
            is_sel = (key == proveedor)
            bg = prov_colors.get(key, col) if is_sel else C_GRAY
            fg = C_WHITE if is_sel else C_TEXT
            card.config(bg=bg)
            for w in card.winfo_children():
                try: w.config(bg=bg, fg=fg)
                except Exception: pass

        nota = PROVIDERS.get(proveedor, {}).get("nota", "")
        self._lbl_prov_nota.config(text=nota)
        # Auto-rellenar host/puerto/ssl
        cfg_p = config_proveedor(proveedor)
        if cfg_p.get("host_smtp"):
            self._email_vars["host_smtp"].set(cfg_p["host_smtp"])
            self._email_vars["puerto"].set(str(cfg_p["puerto"]))
            self._ssl_var.set(cfg_p["ssl"])

    def _on_remitente_change(self, *_):
        email = self._email_vars["remitente"].get()
        prov  = detectar_proveedor(email)
        if prov != "smtp":
            self._email_prov_var.set(prov)
            prov_colors = {"gmail":"#EA4335","outlook":"#0078D4","yahoo":"#6001D2","smtp":"#374151"}
            self._on_prov_change(prov, prov_colors.get(prov, ""))

    def _toggle_email_pwd(self):
        if self._pwd_entry_ref:
            cur = self._pwd_entry_ref.cget("show")
            self._pwd_entry_ref.config(show="" if cur == "*" else "*")

    def _guardar_email(self):
        for key, var in self._email_vars.items():
            self._email_cfg[key] = var.get()
        self._email_cfg["proveedor"] = self._email_prov_var.get()
        self._email_cfg["ssl"]       = self._ssl_var.get()
        try: self._email_cfg["puerto"] = int(self._email_cfg.get("puerto", 587))
        except ValueError: pass
        _guardar_email_cfg(self._email_cfg)
        messagebox.showinfo("Guardado", "✅ Configuración de email guardada.")

    def _probar_email(self):
        self._lbl_email_status.config(text="⏳ Probando…", fg=C_WARN)
        self.update()
        cfg = {
            "remitente": self._email_vars["remitente"].get(),
            "password":  self._email_vars["password"].get(),
            "host_smtp": self._email_vars["host_smtp"].get(),
            "puerto":    int(self._email_vars["puerto"].get() or 587),
            "ssl":       self._ssl_var.get(),
        }
        def tarea():
            ok, msg = probar_conexion(cfg)
            color = C_GREEN if ok else C_RED
            self.after(0, lambda: self._lbl_email_status.config(text=msg.split("\n")[0], fg=color))
        threading.Thread(target=tarea, daemon=True).start()

    # ── Usuarios ──────────────────────────────────────────────────────────────
    def _build_users_config(self):
        f = self._cfg_users_tab
        tk.Label(f, text="Gestión de Usuarios",
                 bg=C_BG, fg=C_TEXT, font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=20, pady=(14, 6))

        # Lista de usuarios
        lst_f = tk.Frame(f, bg=C_BG); lst_f.pack(padx=20, fill="x", pady=(0, 10))
        tk.Label(lst_f, text="Usuarios registrados:", bg=C_BG, fg=C_TEXT, font=FONT_HEAD).pack(anchor="w")
        self._usr_list = tk.Listbox(lst_f, font=FONT_MONO, height=6, bg=C_GRAY,
                                     fg=C_TEXT, relief="flat", bd=4, selectbackground=C_ACCENT)
        self._usr_list.pack(fill="x", pady=(4, 8))
        self._reload_users()

        bf = tk.Frame(f, bg=C_BG); bf.pack(fill="x", padx=20)
        self._rbtn(bf, "🗑 Eliminar usuario", self._eliminar_usuario, "danger", small=True
                   ).pack(side="left", padx=(0, 10))
        self._rbtn(bf, "🔄 Actualizar lista", self._reload_users, "ghost", small=True
                   ).pack(side="left")

        # Crear usuario
        tk.Frame(f, bg=C_GRAY, height=1).pack(fill="x", padx=20, pady=14)
        tk.Label(f, text="Crear nuevo usuario:", bg=C_BG, fg=C_TEXT,
                 font=FONT_HEAD).pack(anchor="w", padx=20)
        card = tk.Frame(f, bg=C_WHITE); card.pack(padx=20, fill="x", pady=(6, 0))
        self._nu_vars: dict[str, tk.StringVar] = {}
        for i, (lb, key, show) in enumerate([
            ("Usuario:", "user", ""), ("Contraseña:", "pwd", "*"), ("Confirmar:", "pwd2", "*")
        ]):
            tk.Label(card, text=lb, bg=C_WHITE, fg=C_GRAY2, font=FONT_NORMAL).grid(
                row=i, column=0, sticky="w", padx=16, pady=(10, 0))
            var = tk.StringVar()
            ef  = tk.Frame(card, bg=C_GRAY)
            ef.grid(row=i, column=1, sticky="ew", padx=16, pady=(10, 0))
            tk.Entry(ef, textvariable=var, show=show, font=("Segoe UI", 10),
                     bg=C_GRAY, fg=C_TEXT, relief="flat", bd=4,
                     insertbackground=C_ACCENT).pack(fill="x")
            self._nu_vars[key] = var
        self._nu_msg = tk.Label(card, text="", bg=C_WHITE, fg=C_RED, font=FONT_SMALL)
        self._nu_msg.grid(row=3, column=1, sticky="w", padx=16, pady=(4, 10))
        card.columnconfigure(1, weight=1)

        self._rbtn(f, "➕ Crear usuario", self._crear_usuario_cfg, "accent"
                   ).pack(anchor="w", padx=20, pady=10)

    def _reload_users(self):
        self._usr_list.delete(0, "end")
        for u in listar_usuarios():
            self._usr_list.insert("end", f"  👤  {u}")

    def _eliminar_usuario(self):
        sel = self._usr_list.curselection()
        if not sel: return
        u = self._usr_list.get(sel[0]).strip().replace("👤", "").strip()
        if u == self._usuario:
            messagebox.showwarning("No permitido", "No puedes eliminar tu propio usuario.", parent=self); return
        ok, msg = eliminar_usuario(u)
        if ok: self._reload_users()
        else:  messagebox.showerror("Error", msg, parent=self)

    def _crear_usuario_cfg(self):
        u  = self._nu_vars["user"].get().strip()
        p  = self._nu_vars["pwd"].get()
        p2 = self._nu_vars["pwd2"].get()
        if p != p2:
            self._nu_msg.config(text="Las contraseñas no coinciden."); return
        ok, msg, code = crear_usuario(u, p, creado_por=self._usuario)
        if ok:
            self._nu_msg.config(text="")
            for v in self._nu_vars.values(): v.set("")
            messagebox.showinfo("Usuario creado",
                f"✅ Usuario '{u}' creado.\nCódigo de recuperación:\n\n{code}\n\n"
                "Guárdalo, solo se muestra una vez.", parent=self)
            self._reload_users()
        else:
            self._nu_msg.config(text=msg)


    # ══════════════════════════════════════════════════════════════════════════
    #  TAB: CONCILIACIÓN CFDI
    # ══════════════════════════════════════════════════════════════════════════
    def _build_conciliacion_tab(self):
        f = self._tab_concil

        if not _CONCIL_OK:
            tk.Label(f, text="⚠  Módulos de conciliación no encontrados.\n"
                              "Verifica que cfdi_parser.py, conciliacion.py y\n"
                              "exportar_conciliacion.py estén en la misma carpeta.",
                     bg=C_BG, fg=C_RED, font=FONT_HEAD, justify="center"
                     ).pack(expand=True)
            return

        # ── Cabecera del módulo (blanca, estilo prototipo) ────────────────────
        hdr = tk.Frame(f, bg=C_WHITE, highlightbackground=T.BORDER,
                       highlightthickness=1); hdr.pack(fill="x")
        if _HAS_CTK:
            ctk.CTkLabel(hdr, text="⇄", fg_color=T.ACCENT_SOFT, text_color=T.ACCENT,
                         corner_radius=10, width=38, height=38,
                         font=("Segoe UI", 14, "bold")).pack(side="left", padx=(16, 12), pady=10)
        else:
            tk.Label(hdr, text="⇄", bg=T.ACCENT_SOFT, fg=T.ACCENT, width=3,
                     font=("Segoe UI", 14, "bold")).pack(side="left", padx=(16, 12), pady=10)
        _htt = tk.Frame(hdr, bg=C_WHITE); _htt.pack(side="left", pady=8)
        tk.Label(_htt, text="Conciliación Bancaria ↔ CFDI", bg=C_WHITE, fg=T.TEXT,
                 font=("Segoe UI", 13, "bold")).pack(anchor="w")
        tk.Label(_htt, text="Cruza movimientos bancarios con los CFDI del SAT y exporta el papel de trabajo",
                 bg=C_WHITE, fg=T.TEXT_MUTED, font=FONT_SMALL).pack(anchor="w")

        # ── Área scrollable (barra vertical a la DERECHA); header fijo arriba ──
        _scan = tk.Canvas(f, bg=C_BG, highlightthickness=0)
        _svsb = ttk.Scrollbar(f, orient="vertical", command=_scan.yview)
        _scan.configure(yscrollcommand=_svsb.set)
        _svsb.pack(side="right", fill="y")
        _scan.pack(side="left", fill="both", expand=True)
        _sinner = tk.Frame(_scan, bg=C_BG)
        _swin = _scan.create_window((0, 0), window=_sinner, anchor="nw")
        _sinner.bind("<Configure>",
                     lambda e: _scan.configure(scrollregion=_scan.bbox("all")))
        _scan.bind("<Configure>",
                   lambda e: _scan.itemconfigure(_swin, width=e.width))
        def _concil_wheel(e):
            _scan.yview_scroll(int(-e.delta / 120), "units")
        _scan.bind("<Enter>", lambda e: _scan.bind_all("<MouseWheel>", _concil_wheel))
        _scan.bind("<Leave>", lambda e: _scan.unbind_all("<MouseWheel>"))
        f = _sinner   # a partir de aquí, todo va DENTRO del área scrollable

        # ── Cuerpo: dos columnas superiores ───────────────────────────────────
        body = tk.Frame(f, bg=C_BG); body.pack(fill="both", expand=True, padx=12, pady=8)
        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=1)

        # ── Columna A: XMLs (tarjeta) ─────────────────────────────────────────
        ca = self._card_badge(body, "1", "Carpetas con XML (CFDI)",
                              grid=dict(row=0, column=0, sticky="nsew", padx=(0, 6), pady=(0, 6)))
        ca.columnconfigure(0, weight=1)

        bx = tk.Frame(ca, bg=C_WHITE); bx.grid(row=0, column=0, sticky="ew", pady=(2, 8))
        self._rbtn(bx, "＋  Agregar carpeta", self._concil_add_carpeta, "primary"
                   ).grid(row=0, column=0, sticky="w")
        self._rbtn(bx, "📄  Agregar XML (sin carpeta)", self._concil_add_xml, "primary",
                   small=True).grid(row=1, column=0, columnspan=3, sticky="w", pady=(4, 0))
        self._rbtn(bx, "📥 Descargar del SAT", self._abrir_descarga_sat, "accent", small=True
                   ).grid(row=0, column=1, padx=(6, 0))
        self._rbtn(bx, "✕ Limpiar", self._concil_clear_carpetas, "ghost", small=True
                   ).grid(row=0, column=2, padx=(6, 0))

        self._concil_lst_xml = tk.Listbox(ca, height=6, font=FONT_SMALL,
                                           bg="#F9F9F9", fg=C_TEXT,
                                           selectbackground=C_ACCENT,
                                           relief="flat", bd=0)
        self._concil_lst_xml.grid(row=1, column=0, sticky="ew", pady=(0, 4))

        self._concil_lbl_xml = tk.Label(ca, text="0 carpetas  |  0 XMLs cargados",
                                         bg=C_WHITE, fg=C_GRAY2, font=FONT_SMALL)
        self._concil_lbl_xml.grid(row=2, column=0, sticky="e", pady=(0, 2))

        # ── Columna B: PDFs (tarjeta + filas con pill de banco) ───────────────
        cb = self._card_badge(body, "2", "Estados de cuenta (PDF)",
                              grid=dict(row=0, column=1, sticky="nsew", padx=(6, 0), pady=(0, 6)))
        cb.columnconfigure(0, weight=1)

        bp = tk.Frame(cb, bg=C_WHITE); bp.grid(row=0, column=0, sticky="ew", pady=(2, 8))
        self._rbtn(bp, "＋  Agregar PDFs", self._concil_add_pdf, "primary"
                   ).grid(row=0, column=0, sticky="w")
        self._rbtn(bp, "✕ Limpiar", self._concil_clear_pdfs, "ghost", small=True
                   ).grid(row=0, column=1, padx=(6, 0))

        _wrap = tk.Frame(cb, bg=C_WHITE, highlightbackground=T.BORDER, highlightthickness=1)
        _wrap.grid(row=1, column=0, sticky="ew", pady=(0, 4))
        _cv = tk.Canvas(_wrap, bg=C_WHITE, highlightthickness=0, height=150)
        sb_p = ttk.Scrollbar(_wrap, orient="vertical", command=_cv.yview)
        self._concil_pdf_inner = tk.Frame(_cv, bg=C_WHITE)
        self._concil_pdf_inner.bind(
            "<Configure>", lambda e: _cv.configure(scrollregion=_cv.bbox("all")))
        _pwin = _cv.create_window((0, 0), window=self._concil_pdf_inner, anchor="nw")
        _cv.bind("<Configure>", lambda e: _cv.itemconfigure(_pwin, width=e.width))
        _cv.configure(yscrollcommand=sb_p.set)
        _cv.pack(side="left", fill="both", expand=True)
        sb_p.pack(side="right", fill="y")

        self._concil_lbl_pdf = tk.Label(cb, text="0 PDFs",
                                         bg=C_WHITE, fg=C_GRAY2, font=FONT_SMALL)
        self._concil_lbl_pdf.grid(row=2, column=0, sticky="e", pady=(0, 2))

        # ── Paso 3: configuración + periodo (tarjeta, etiqueta arriba) ────────
        cc = self._card_badge(f, "3", "Configuración y periodo",
                              pack=dict(fill="x", padx=12, pady=(0, 6)))
        cc.columnconfigure(0, weight=1)

        # Variable que almacena sólo el RFC (lo que usa la conciliación)
        self._concil_rfc = tk.StringVar()
        # Variable de display: "RFC  —  Nombre"
        self._concil_emp_var = tk.StringVar()

        # Fila de campos estilo prototipo: etiqueta ARRIBA, compactos a la izq
        top = tk.Frame(cc, bg=C_WHITE)
        top.grid(row=0, column=0, sticky="ew", padx=8, pady=(6, 2))

        gemp = tk.Frame(top, bg=C_WHITE); gemp.pack(side="left", padx=(0, 16))
        tk.Label(gemp, text="Empresa", bg=C_WHITE, fg=T.TEXT_MUTED,
                 font=FONT_SMALL, anchor="w").pack(anchor="w")
        _erow = tk.Frame(gemp, bg=C_WHITE); _erow.pack(anchor="w")
        self._concil_emp_combo = ttk.Combobox(
            _erow, textvariable=self._concil_emp_var,
            state="readonly", width=38, font=FONT_NORMAL)
        self._concil_emp_combo.pack(side="left")
        tk.Button(_erow, text="🔄", command=self._concil_refresh_empresas,
                  bg=C_WHITE, fg=C_TEXT, highlightthickness=1, highlightbackground=T.BORDER, font=FONT_SMALL,
                  relief="flat", padx=6, pady=2, cursor="hand2").pack(side="left", padx=(4, 0))

        gan = tk.Frame(top, bg=C_WHITE); gan.pack(side="left", padx=(0, 16))
        tk.Label(gan, text="Año a conciliar", bg=C_WHITE, fg=T.TEXT_MUTED,
                 font=FONT_SMALL, anchor="w").pack(anchor="w")
        self._concil_anio = ttk.Combobox(
            gan, values=[str(y) for y in CE.anios_disponibles()],
            width=8, state="readonly", font=FONT_HEAD)
        self._concil_anio.set(self._anio_default())
        self._concil_anio.pack(anchor="w")

        gdst = tk.Frame(top, bg=C_WHITE); gdst.pack(side="left", fill="x", expand=True)
        tk.Label(gdst, text="Guardar en", bg=C_WHITE, fg=T.TEXT_MUTED,
                 font=FONT_SMALL, anchor="w").pack(anchor="w")
        _drow = tk.Frame(gdst, bg=C_WHITE); _drow.pack(anchor="w", fill="x")
        self._concil_destino = tk.StringVar()
        tk.Entry(_drow, textvariable=self._concil_destino, font=FONT_SMALL,
                 bg=C_GRAY, fg=C_TEXT, relief="flat", bd=4).pack(side="left", fill="x", expand=True)
        tk.Button(_drow, text="…", command=self._concil_elegir_destino,
                  bg=C_WHITE, fg=C_TEXT, highlightthickness=1, highlightbackground=T.BORDER, font=FONT_NORMAL,
                  relief="flat", padx=8, pady=2, cursor="hand2").pack(side="left", padx=(4, 0))

        # Extraer RFC cuando el usuario selecciona una empresa y mostrar régimen
        def _on_emp_select(*_):
            val = self._concil_emp_var.get()
            rfc = val.split("  —  ")[0].strip() if "  —  " in val else val.strip()
            self._concil_rfc.set(rfc)
            # Actualizar label de régimen
            emp = obtener_empresa(rfc) if rfc else None
            if emp:
                codigos = emp.get("regimenes", [])
                nombres = emp.get("regimenes_nombres", [])
                if codigos and nombres:
                    partes = [f"{n} ({c})" for n, c in zip(nombres, codigos)]
                    self._concil_reg_lbl.config(
                        text=f"  ✔  Régimen(es): {',  '.join(partes)}", fg="#2E7D32")
                elif codigos:
                    self._concil_reg_lbl.config(
                        text=f"  Régimen(es): {', '.join(codigos)}", fg="#2E7D32")
                else:
                    self._concil_reg_lbl.config(
                        text="  Sin CSF registrada — sube la Constancia para validar régimen fiscal",
                        fg="#E65100")
            else:
                self._concil_reg_lbl.config(text="", fg=C_GRAY2)

        self._concil_emp_var.trace_add("write", _on_emp_select)

        # RFC readonly (informativo)
        self._concil_rfc_lbl = tk.Label(cc, textvariable=self._concil_rfc,
                                         bg=C_WHITE, fg=C_GRAY2, font=("Consolas", 8))
        self._concil_rfc_lbl.grid(row=1, column=0, sticky="w", padx=10, pady=(4, 0))

        # Régimen fiscal de la empresa seleccionada
        self._concil_reg_lbl = tk.Label(cc, text="", bg=C_WHITE, fg=C_GRAY2,
                                         font=FONT_SMALL, anchor="w")
        self._concil_reg_lbl.grid(row=2, column=0, sticky="w", padx=10, pady=(0, 2))

        # ── Fila de acciones: Forma de pago + Catálogos COI + nota ────────────
        _acc3 = tk.Frame(cc, bg=C_WHITE)
        _acc3.grid(row=3, column=0, sticky="w", padx=10, pady=(2, 8))
        self._rbtn(_acc3, "💳  Forma de pago…", self._concil_abrir_formas_pago,
                   "accent", small=True).pack(side="left")
        self._concil_forma_pago = tk.StringVar(value="Todas (01–99)")
        tk.Label(_acc3, textvariable=self._concil_forma_pago, bg=C_WHITE,
                 fg=T.ACCENT_INK, font=("Segoe UI", 9, "bold")).pack(side="left", padx=(8, 16))
        if _COI_OK:
            self._rbtn(_acc3, "🗂  Catálogos COI…", self._concil_abrir_catalogos_win,
                       "primary", small=True).pack(side="left", padx=(0, 16))

        # — Meses para cálculo de IVA (multiselección) —
        _nombres_mes = [
            (1,"Ene"),(2,"Feb"),(3,"Mar"),(4,"Abr"),
            (5,"May"),(6,"Jun"),(7,"Jul"),(8,"Ago"),
            (9,"Sep"),(10,"Oct"),(11,"Nov"),(12,"Dic"),
        ]
        self._concil_meses_iva: dict[int, tk.BooleanVar] = {}
        tk.Label(_acc3, text="Meses IVA:", bg=C_WHITE, fg=T.TEXT_MUTED,
                 font=FONT_SMALL).pack(side="left", padx=(16, 4))
        self._concil_lbl_meses_iva = tk.Label(
            _acc3, text="Todos", bg=C_WHITE, fg=T.ACCENT_INK,
            font=("Segoe UI", 9, "bold"))
        self._concil_lbl_meses_iva.pack(side="left", padx=(0, 4))

        _mb_iva = tk.Menubutton(_acc3, text="▼", bg=C_WHITE, fg=C_TEXT,
                                relief="flat", font=FONT_SMALL, cursor="hand2",
                                highlightthickness=1, highlightbackground=T.BORDER,
                                padx=6, pady=1)
        _mb_iva.pack(side="left")
        _menu_iva = tk.Menu(_mb_iva, tearoff=0)
        _mb_iva["menu"] = _menu_iva

        def _upd_lbl_meses(*_):
            sel = [m for m, v in self._concil_meses_iva.items() if v.get()]
            if not sel or len(sel) == 12:
                self._concil_lbl_meses_iva.config(text="Todos")
            else:
                nombres = [n for m, n in _nombres_mes if m in sel]
                self._concil_lbl_meses_iva.config(text=", ".join(nombres))

        for num, corto in _nombres_mes:
            var = tk.BooleanVar(value=False)
            self._concil_meses_iva[num] = var
            _menu_iva.add_checkbutton(
                label=f"{num:02d} - {corto}", variable=var,
                command=_upd_lbl_meses)

        tk.Label(_acc3, text="Solo del año elegido.",
                 bg=C_WHITE, fg=T.ACCENT, font=FONT_SMALL).pack(side="left", padx=(16, 0))

        # Poblar la lista al construir el tab
        self._concil_refresh_empresas()

        # ── Catálogos COI → ahora en VENTANA (botón en PASO 3) ────────────────
        if _COI_OK and _GRUPOS_OK:
            self._concil_rfc.trace_add(
                "write", lambda *_: self._concil_refrescar_lbl_grupos())

        # ── Fases de conciliación (casillas) ──────────────────────────────────
        # Agrupadas por NIVEL DE CONFIANZA, con el mismo color que la columna
        # TIPO MATCH del reporte Excel: lo que aquí se marca en violeta sale
        # marcado en violeta en el Excel.
        ce = self._card_badge(f, "⚙", "Fases de conciliación",
                              pack=dict(fill="x", padx=12, pady=(0, 6)))

        tk.Label(
            ce, text="Marca las que quieras aplicar. Se ejecutan en orden: "
                     "cada fase solo ve lo que la anterior dejó sin conciliar.",
            bg=C_WHITE, fg=T.TEXT_MUTED, font=("Segoe UI", 8),
            justify="left", anchor="w", wraplength=620
        ).pack(fill="x", padx=12, pady=(4, 6))

        # fid → nivel. Un id que no esté aquí cae en su propio grupo "Otras",
        # así que agregar una fase nueva al catálogo nunca la deja invisible.
        _NIVEL_DE = {
            "exacto":       0, "multi_mov": 1, "combinatorio": 1, "grupos": 1,
            "ppd_con_rfc":  2, "ppd_con_nom": 2,
            "ppd_sin_rfc":  3, "ppd_sin_nom": 3, "solo_importe": 3,
        }
        _NIVELES = [
            ("Directo",         T.TIER_DIRECTO,
             "La evidencia más fuerte: identidad + importe + mes."),
            ("Compuesto",       T.TIER_COMPUESTO,
             "Sumas y agrupaciones. Verificar la asignación individual."),
            ("Con complemento", T.TIER_COMPLEM,
             "Respaldado por el REP (complemento de pago)."),
            ("Débil",           T.TIER_DEBIL,
             "Sin complemento o solo por importe. Exige revisión."),
            ("Otras",           T.TEXT_MUTED, ""),
        ]

        self._concil_fase_vars: dict = {}
        _por_nivel: dict[int, list] = {}
        for fid, etiqueta in FASES_CONCILIACION:
            _por_nivel.setdefault(_NIVEL_DE.get(fid, 4), []).append((fid, etiqueta))

        for _ix, (_nom, _color, _desc) in enumerate(_NIVELES):
            _fases = _por_nivel.get(_ix)
            if not _fases:
                continue
            # Barra de color a la izquierda = la marca de identidad del nivel.
            _fila = tk.Frame(ce, bg=C_WHITE)
            _fila.pack(fill="x", padx=12, pady=(2, 0))
            tk.Frame(_fila, bg=_color, width=4).pack(side="left", fill="y")
            _cuerpo = tk.Frame(_fila, bg=C_WHITE)
            _cuerpo.pack(side="left", fill="x", expand=True, padx=(8, 0))

            _enc = tk.Frame(_cuerpo, bg=C_WHITE)
            _enc.pack(fill="x")
            tk.Label(_enc, text=_nom, bg=C_WHITE, fg=_color,
                     font=("Segoe UI", 9, "bold"), anchor="w").pack(side="left")
            if _desc:
                tk.Label(_enc, text="  " + _desc, bg=C_WHITE, fg=T.TEXT_MUTED,
                         font=("Segoe UI", 8), anchor="w").pack(side="left")

            for fid, etiqueta in _fases:
                var = tk.BooleanVar(value=(fid in FASES_DEFAULT))
                self._concil_fase_vars[fid] = var
                tk.Checkbutton(
                    _cuerpo, text=etiqueta, variable=var,
                    bg=C_WHITE, fg=C_TEXT, font=FONT_SMALL, anchor="w",
                    activebackground=C_WHITE, selectcolor=C_WHITE,
                    cursor="hand2"
                ).pack(fill="x", padx=(4, 0))

        tk.Label(
            ce, text="La Fase 7 cruza solo por importe del mismo mes, sin validar "
                     "al cliente: úsala como último recurso. Todo lo que concilie "
                     "queda marcado para revisión.",
            bg=C_WHITE, fg=T.TEXT_MUTED, font=("Segoe UI", 8), justify="left",
            wraplength=620, anchor="w"
        ).pack(fill="x", padx=12, pady=(6, 8))

        # ── Botón principal: alterna CONCILIAR (verde) ⇆ DETENER (rojo) según el
        #    estado. Así el "detener" no ocupa una fila extra (acabado más limpio
        #    y sin desbordar la pestaña). ─────────────────────────────────────
        # ── Fila de acción: CONCILIAR (expande) + Pólizas COI (al lado) ───────
        # ── Fila XML: Actualizar XML + Reconciliar (a la altura de CONCILIAR) ─
        _xmlrow = tk.Frame(f, bg=C_BG)
        _xmlrow.pack(fill="x", padx=12, pady=(6, 2))
        self._rbtn(_xmlrow, "🔄  Actualizar XML", self._concil_subir_xml,
                   "accent").pack(side="left", fill="x", expand=True)
        self._rbtn(_xmlrow, "🔗  Reconciliar (sin CFDI)", self._concil_reconciliar_dialog,
                   "primary").pack(side="left", fill="x", expand=True, padx=(8, 0))

        _accrow = tk.Frame(f, bg=C_BG)
        _accrow.pack(fill="x", padx=12, pady=(2, 4))
        if _HAS_CTK:
            self._concil_btn = ctk.CTkButton(
                _accrow, text="⚡  LEER PDFs  →  CONCILIAR  →  EXPORTAR EXCEL",
                command=self._concil_ejecutar, corner_radius=10,
                fg_color=C_GREEN, text_color=C_WHITE,
                font=("Segoe UI", 12, "bold"), height=46)
        else:
            self._concil_btn = tk.Button(
                _accrow, text="⚡  LEER PDFs  →  CONCILIAR  →  EXPORTAR EXCEL",
                command=self._concil_ejecutar,
                bg=C_GREEN, fg=C_WHITE, font=("Segoe UI", 12, "bold"),
                relief="flat", pady=10, cursor="hand2")
        self._concil_btn.pack(side="left", fill="x", expand=True)

        # ── Botón Pólizas COI (misma fila, a la derecha) ──────────────────────
        if _COI_OK:
            if _HAS_CTK:
                self._concil_btn_polizas = ctk.CTkButton(
                    _accrow, text="Pólizas COI  (tras conciliar)",
                    command=self._concil_abrir_polizas, corner_radius=10,
                    fg_color="#37474F", text_color="#90A4AE", width=210,
                    font=("Segoe UI", 10, "bold"), height=46, state="disabled")
            else:
                self._concil_btn_polizas = tk.Button(
                    _accrow, text="Pólizas COI  (tras conciliar)",
                    command=self._concil_abrir_polizas,
                    bg="#37474F", fg="#90A4AE", font=("Segoe UI", 10, "bold"),
                    relief="flat", pady=10, cursor="hand2", state="disabled")
            self._concil_btn_polizas.pack(side="left", padx=(8, 0),
                                          fill="y" if not _HAS_CTK else None)

        # Estado de última conciliación (para pasar a VentanaPolizas)
        self._concil_ultimo: dict | None = None

        self._concil_prog = ttk.Progressbar(f, mode="indeterminate")
        self._concil_prog.pack(fill="x", padx=12, pady=(0, 4))

        # ── Log ───────────────────────────────────────────────────────────────
        self._concil_log = scrolledtext.ScrolledText(
            f, height=6, wrap="word", font=FONT_MONO,
            bg="#1A1A19", fg="#C3C2B7", relief="flat", bd=0, state="disabled")
        self._concil_log.pack(fill="both", expand=True, padx=12, pady=(0, 8))
        self._concil_log.tag_configure("ok",   foreground="#4ADE80")
        self._concil_log.tag_configure("err",  foreground="#F87171")
        self._concil_log.tag_configure("info", foreground="#60A5FA")
        self._concil_log.tag_configure("warn", foreground="#FACC15")

    # ── Conciliación: acciones ────────────────────────────────────────────────

    def _concil_guardar_bd(self, rfc, resultados, recibidos=None, emitidos=None):
        """Guarda la conciliación en Firebird y estampa el ID ESTABLE (BD) en
        cada MovBanco como `id_fila` → aparece en la columna A del Excel.
        Guarda TAMBIÉN todos los CFDI (tabla CFDI) para poder regenerar las
        pólizas desde la BD sin volver a los XML. La BD (IDs monotónicos) hace del
        ID la llave de sincronización para reimportar sin sobrescribir."""
        import persistencia as _persist
        try:
            periodo = self._concil_anio.get()
        except Exception:
            periodo = ""

        def _tipo_mov(mb):
            return "recibido" if (mb.retiro or 0) else ("emitido" if (mb.deposito or 0) else "")

        pares = []   # [(MovBanco, FilaConciliacion), ...]  (misma referencia)
        for res in resultados:
            _bloque = []   # este banco: se ordena por posicion original del PDF
            for mc in res.conciliados:
                mb = mc.banco
                _suma = ",".join(getattr(c, "uuid", "") for c in getattr(mc, "cfdi_suma", []) if getattr(c, "uuid", ""))
                _bloque.append((mb, _persist.FilaConciliacion(
                    rfc=rfc, banco=res.nombre, fecha=str(mb.fecha or ""),
                    importe=float(mb.deposito or 0) - float(mb.retiro or 0),
                    concepto=mb.descripcion or "", uuid_cfdi=mb.uuid or "",
                    uuid_rep=getattr(mb, "uuid_rep", "") or "",
                    tipo_match=mc.tipo_match or "", revisar=bool(mc.revisar),
                    tipo=mc.tipo or _tipo_mov(mb), uuid_suma=_suma,
                    deposito=float(mb.deposito or 0), retiro=float(mb.retiro or 0))))
            for mb in res.sin_cfdi:
                _bloque.append((mb, _persist.FilaConciliacion(
                    rfc=rfc, banco=res.nombre, fecha=str(mb.fecha or ""),
                    importe=float(mb.deposito or 0) - float(mb.retiro or 0),
                    concepto=mb.descripcion or "", tipo=_tipo_mov(mb),
                    deposito=float(mb.deposito or 0), retiro=float(mb.retiro or 0))))
            # Guardar en el ORDEN DEL ESTADO DE CUENTA (PDF): asi el ID estable se
            # asigna en ese orden y el detalle (ORDER BY ID) NO reordena.
            _bloque.sort(key=lambda pr: getattr(pr[0], "orden", 0))
            pares.extend(_bloque)
        if not pares:
            return

        def _cfdi_data(c):
            return _persist.CFDIData(
                uuid=getattr(c, "uuid", "") or "", tipo=getattr(c, "tipo", "") or "",
                rfc_emisor=getattr(c, "rfc_emisor", "") or "",
                rfc_receptor=getattr(c, "rfc_receptor", "") or "",
                nombre_emisor=getattr(c, "nombre_emisor", "") or "",
                nombre_receptor=getattr(c, "nombre_receptor", "") or "",
                subtotal=float(getattr(c, "subtotal", 0) or 0),
                total=float(getattr(c, "total", 0) or 0),
                descuento=float(getattr(c, "descuento", 0) or 0),
                iva=float(getattr(c, "iva16", 0) or 0),
                iva_retenido=float(getattr(c, "iva_retenido", 0) or 0),
                isr_retenido=float(getattr(c, "isr_retenido", 0) or 0),
                clave_prod_serv=getattr(c, "clave_prod_serv", "") or "",
                fecha_emision=getattr(c, "fecha_emision", "") or "")

        p = _persist.FirebirdPersistencia()
        p.inicializar()
        try:
            filas = [f for _, f in pares]
            p.guardar_conciliacion(rfc, periodo, filas)   # asigna IDs in-place
            for mb, f in pares:
                mb.id_fila = f.id      # estampar el ID → sale en columna A
            cfdis = [_cfdi_data(c) for c in (list(recibidos or []) + list(emitidos or []))]
            if cfdis:
                p.guardar_cfdis(cfdis)
            self._concil_log_write(
                f"  BD Firebird: {len(filas)} filas + {len(cfdis)} CFDI guardados "
                f"(ID estable en col A).", "info")
        finally:
            p.cerrar()

    def _wget(self, name):
        """Devuelve el widget self.<name> si existe y NO ha sido destruido
        (p.ej. las etiquetas de la ventana de Catálogos, que puede estar cerrada)."""
        w = getattr(self, name, None)
        try:
            return w if (w is not None and w.winfo_exists()) else None
        except Exception:
            return None

    def _set_concil_btn(self, text=None, color=None, command=None, state=None):
        """Reconfigura el botón principal soportando CTkButton (fg_color) y
        tk.Button (bg) de forma transparente."""
        b = self._concil_btn
        kw = {}
        if text is not None:    kw["text"] = text
        if command is not None: kw["command"] = command
        if state is not None:   kw["state"] = state
        if color is not None:
            kw["fg_color" if _HAS_CTK else "bg"] = color
        (b.configure if _HAS_CTK else b.config)(**kw)

    def _concil_abrir_formas_pago(self):
        """Ventana de SELECCIÓN MÚLTIPLE de Formas de Pago (c_FormaPago 01–99).
        Guarda los códigos elegidos en ``self._concil_formas_sel`` (set) y
        refresca la etiqueta. Vacío = todas."""
        if not hasattr(self, "_concil_formas_sel"):
            self._concil_formas_sel = set()
        win = tk.Toplevel(self); win.title("Formas de pago (SAT)")
        win.configure(bg=T.SURFACE); win.geometry("470x560"); win.transient(self)
        tk.Label(win, text="Selecciona una o varias formas de pago",
                 font=("Segoe UI", 12, "bold"), bg=T.SURFACE, fg=T.TEXT
                 ).pack(anchor="w", padx=16, pady=(14, 2))
        tk.Label(win, text="Catálogo c_FormaPago del SAT (01 – 99) · marca las que quieras conciliar",
                 font=FONT_SMALL, bg=T.SURFACE, fg=T.TEXT_MUTED).pack(anchor="w", padx=16, pady=(0, 8))

        vars_ = {}
        bar = tk.Frame(win, bg=T.SURFACE); bar.pack(fill="x", padx=16, pady=(0, 6))
        def _marca(val):
            for v in vars_.values():
                v.set(val)
        tk.Button(bar, text="Todos", relief="flat", font=("Segoe UI", 8, "bold"),
                  cursor="hand2", bg=T.SURFACE_ALT, fg=T.TEXT, bd=0, padx=10, pady=4,
                  command=lambda: _marca(True)).pack(side="left", padx=(0, 6))
        tk.Button(bar, text="Ninguno", relief="flat", font=("Segoe UI", 8, "bold"),
                  cursor="hand2", bg=T.SURFACE_ALT, fg=T.TEXT, bd=0, padx=10, pady=4,
                  command=lambda: _marca(False)).pack(side="left")

        cont = tk.Frame(win, bg=T.SURFACE); cont.pack(fill="both", expand=True, padx=12)
        canvas = tk.Canvas(cont, bg=T.SURFACE, highlightthickness=0)
        sb = tk.Scrollbar(cont, orient="vertical", command=canvas.yview)
        inner = tk.Frame(canvas, bg=T.SURFACE)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True); sb.pack(side="right", fill="y")

        for cod, desc in FORMAS_PAGO:
            v = tk.BooleanVar(value=(cod in self._concil_formas_sel)); vars_[cod] = v
            row = tk.Frame(inner, bg=T.SURFACE_ALT); row.pack(fill="x", pady=1)
            tk.Checkbutton(row, variable=v, bg=T.SURFACE_ALT, activebackground=T.SURFACE_ALT,
                           selectcolor=T.SURFACE, bd=0, highlightthickness=0).pack(side="left", padx=(6, 2))
            tk.Label(row, text=cod, width=4, font=("Segoe UI", 10, "bold"),
                     bg=T.ACCENT_SOFT, fg=T.ACCENT_INK).pack(side="left", padx=(2, 8), pady=2)
            tk.Label(row, text=desc, anchor="w", font=("Segoe UI", 10),
                     bg=T.SURFACE_ALT, fg=T.TEXT).pack(side="left", fill="x", expand=True)

        def aceptar():
            sel = {c for c in vars_ if vars_[c].get()}
            self._concil_formas_sel = sel
            self._concil_forma_pago.set(
                ", ".join(sorted(sel)) if sel else "Todas (01–99)")
            win.destroy()
        bf = tk.Frame(win, bg=T.SURFACE); bf.pack(fill="x", padx=16, pady=10)
        tk.Button(bf, text="✔  Aceptar selección", relief="flat", font=("Segoe UI", 9, "bold"),
                  cursor="hand2", bg=T.ACCENT, fg=T.ACCENT_FG, bd=0, padx=14, pady=6,
                  command=aceptar).pack(side="left")
        tk.Button(bf, text="Cancelar", relief="flat", font=("Segoe UI", 9, "bold"),
                  cursor="hand2", bg=T.SURFACE_ALT, fg=T.TEXT_MUTED, bd=0, padx=14, pady=6,
                  command=win.destroy).pack(side="left", padx=8)

    def _concil_subir_excel(self, despues=None):
        """SOLO SUBE: guarda en la BD las correcciones del .xlsm revisado (match por
        ID ESTABLE de la columna A). `despues` = callback que se ejecuta tras subir
        (p.ej. refrescar la ventana de Conciliaciones)."""
        rfc = self._concil_rfc_sel()
        if not rfc:
            messagebox.showwarning("Sin empresa",
                                   "Selecciona una empresa en el Paso 3 primero.")
            return
        ruta = filedialog.askopenfilename(
            title="Selecciona el Excel de conciliacion revisado (para GUARDAR)",
            filetypes=[("Excel", "*.xlsm *.xlsx"), ("Todos", "*.*")])
        if not ruta:
            return
        try:
            import reimportar
            st = reimportar.subir_excel(ruta, rfc,
                                        resolver=self._resolver_conflicto,
                                        usuario=self._usuario)
        except Exception as e:
            messagebox.showerror("Error al subir el Excel", str(e))
            return
        _msg = ("Filas leidas: " + str(st.get("filas", 0)) + "\n" +
                "Guardadas/actualizadas en la BD: " + str(st.get("upsert", 0)) + "\n")
        if st.get("conflictos"):
            _msg += ("Conflictos de concurrencia: " + str(st["conflictos"]) +
                     "  (descartados: " + str(st.get("descartados", 0)) +
                     "; en bitacora)\n")
        _sin = st.get("sin_id", 0)
        if _sin:
            _msg += ("\n⚠ " + str(_sin) + " fila(s) NO se actualizaron porque su ID "
                     "(columna A) no existe en la base para esta empresa.\n"
                     "Sube el MISMO Excel que exporto/reimprimio la app (con los IDs "
                     "de la BD en la columna A): usa 'Reimprimir -> Excel', editalo y "
                     "vuelve a subirlo. Un Excel renumerado a mano no casa por ID.")
        elif st.get("upsert"):
            _msg += "\nTus correcciones quedaron guardadas."
        messagebox.showinfo("Excel subido", _msg)
        if callable(despues):
            try: despues()
            except Exception: pass

    def _concil_reimprimir_polizas(self):
        """SOLO REIMPRIME: reconstruye DESDE la BD (sin leer ningun Excel ni escribir
        nada) y abre Polizas COI. Asi no puede cruzarse con lo que estes subiendo."""
        if not _COI_OK:
            return
        rfc = self._concil_rfc_sel()
        if not rfc:
            messagebox.showwarning("Sin empresa",
                                   "Selecciona una empresa en el Paso 3 primero.")
            return
        try:
            import reimportar
            resultados, recibidos, emitidos, stats = reimportar.construir_desde_bd(rfc)
        except Exception as e:
            messagebox.showerror("Error al reimprimir", str(e))
            return
        if not any(r.conciliados for r in resultados):
            messagebox.showwarning(
                "Sin conciliaciones",
                "No hay filas conciliadas con CFDI en la BD para esta empresa.\n\n"
                "Sube primero el Excel revisado con el boton Subir Excel revisado, "
                "o concilia y guarda en la BD.")
            return
        _dest = (getattr(self, "_concil_ultimo", None) or {}).get("destino", "")
        self._concil_ultimo = {
            "resultados": resultados, "recibidos": recibidos,
            "emitidos": emitidos, "destino": _dest, "rfc": rfc,
        }
        _msg = ("Filas en la BD: " + str(stats["filas"]) + "\n" +
                "Conciliadas (con CFDI): " + str(stats["conciliados"]) + "\n")
        if stats.get("sin_datos_cfdi"):
            _msg += ("Sin datos de CFDI en la BD: " + str(stats["sin_datos_cfdi"]) +
                     "  (no generan poliza)\n")
        _msg += "\nAbriendo Polizas COI..."
        messagebox.showinfo("Reimpresion lista", _msg)
        VentanaPolizas(parent=self, rfc_empresa=rfc,
                       todos_recibidos=recibidos, todos_emitidos=emitidos,
                       resultados=resultados, output_excel_path=_dest)

    def _concil_reimprimir_excel(self):
        """Regenera el EXCEL de conciliacion DESDE la BD, con opcion de INTEGRAR una
        o varias conciliaciones (del mes, meses anteriores/futuros, o todo el año
        elegido) en un solo reporte."""
        rfc = self._concil_rfc_sel()
        if not rfc:
            messagebox.showwarning("Sin empresa",
                                   "Selecciona una empresa en el Paso 3 primero.")
            return
        try:
            import persistencia as _persist
            p = _persist.FirebirdPersistencia(); p.inicializar()
            try:
                concs = p.leer_conciliaciones(rfc)
            finally:
                p.cerrar()
        except Exception as e:
            messagebox.showerror("Error de base de datos", str(e), parent=self)
            return
        if not concs:
            messagebox.showwarning(
                "Sin datos",
                "No hay conciliaciones guardadas en la BD para esta empresa.\n\n"
                "Concilia y guarda primero, o sube el Excel revisado.")
            return
        nombre = buscar_por_rfc(rfc) or rfc

        dlg = tk.Toplevel(self); dlg.title("Reimprimir a Excel — elegir conciliaciones")
        dlg.transient(self); dlg.configure(bg=T.SURFACE); dlg.resizable(True, True)
        try: dlg.grab_set()
        except Exception: pass
        tk.Label(dlg, text=f"Integrar conciliaciones — {nombre}", bg=T.SURFACE,
                 fg=T.TEXT, font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=16, pady=(14, 2))
        tk.Label(dlg, text="Marca una o varias conciliaciones para combinarlas en un "
                 "solo Excel (mes, varios meses o todo el año).", bg=T.SURFACE,
                 fg=T.TEXT_MUTED, font=FONT_SMALL, justify="left", wraplength=560
                 ).pack(anchor="w", padx=16, pady=(0, 8))

        # Filtros rapidos por año
        anios = sorted({str(c.get("periodo") or "") for c in concs if c.get("periodo")}, reverse=True)
        fr = tk.Frame(dlg, bg=T.SURFACE); fr.pack(fill="x", padx=16, pady=(0, 6))
        _vars = {}
        def _set_all(v):
            for iid in _vars: _vars[iid].set(v)
        def _set_year(y):
            for c in concs:
                _vars[c["id"]].set(str(c.get("periodo") or "") == y)
        self._rbtn(fr, "Todas", lambda: _set_all(True), "ghost", small=True).pack(side="left", padx=(0, 4))
        self._rbtn(fr, "Ninguna", lambda: _set_all(False), "ghost", small=True).pack(side="left", padx=(0, 8))
        if anios:
            tk.Label(fr, text="Año:", bg=T.SURFACE, fg=T.TEXT_MUTED, font=FONT_SMALL).pack(side="left")
            cb_y = ttk.Combobox(fr, values=anios, width=8, state="readonly"); cb_y.set(anios[0])
            cb_y.pack(side="left", padx=(4, 4))
            self._rbtn(fr, "Marcar año", lambda: _set_year(cb_y.get()), "ghost", small=True).pack(side="left")

        cont = tk.Frame(dlg, bg=T.SURFACE); cont.pack(fill="both", expand=True, padx=12)
        canvas = tk.Canvas(cont, bg=T.SURFACE, highlightthickness=0, height=260)
        sb = ttk.Scrollbar(cont, orient="vertical", command=canvas.yview)
        inner = tk.Frame(canvas, bg=T.SURFACE)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True); sb.pack(side="left", fill="y")
        tk.Label(inner, text="  ✓   Año   Mes(es)         Guardada            Filas / Conc.",
                 bg=T.SURFACE, fg=T.TEXT_MUTED, font=("Consolas", 9)).pack(anchor="w")
        for c in concs:
            v = tk.BooleanVar(value=True); _vars[c["id"]] = v
            txt = ("  %-4s %-14s %-19s %s/%s"
                   % (c.get("periodo") or "—", c.get("meses") or "—",
                      str(c.get("creado") or "")[:19], c["n_filas"], c["n_conciliadas"]))
            tk.Checkbutton(inner, text=txt, variable=v, bg=T.SURFACE, fg=T.TEXT,
                           activebackground=T.SURFACE, selectcolor=C_WHITE,
                           font=("Consolas", 9), anchor="w").pack(fill="x", anchor="w")

        bf = tk.Frame(dlg, bg=T.SURFACE); bf.pack(fill="x", padx=16, pady=12)
        def _generar():
            ids = [cid for cid, vv in _vars.items() if vv.get()]
            if not ids:
                messagebox.showinfo("Selecciona", "Marca al menos una conciliacion.",
                                    parent=dlg)
                return
            try:
                import reimportar
                resultados, recibidos, emitidos, stats = reimportar.construir_desde_bd(
                    rfc, conciliacion_ids=ids)
            except Exception as e:
                messagebox.showerror("Error al reconstruir desde la BD", str(e), parent=dlg)
                return
            if not resultados or not any((r.conciliados or r.sin_cfdi) for r in resultados):
                messagebox.showwarning("Sin datos",
                    "Las conciliaciones elegidas no tienen filas.", parent=dlg)
                return
            _mult = len(ids) > 1
            ruta = filedialog.asksaveasfilename(
                title="Guardar reimpresion de la conciliacion (Excel con macros)",
                defaultextension=".xlsm",
                initialfile=(f"Conciliaciones_{rfc}.xlsm" if _mult else f"Conciliacion_{rfc}.xlsm"),
                filetypes=[("Excel con macros", "*.xlsm"), ("Excel", "*.xlsx"),
                           ("Todos", "*.*")])
            if not ruta:
                return
            try:
                destino = _exp_concil_multi(
                    resultados=resultados, todos_recibidos=recibidos,
                    todos_emitidos=emitidos, output_path=ruta, rfc_empresa=rfc) or ruta
            except Exception as e:
                messagebox.showerror("Error al generar el Excel", str(e), parent=dlg)
                return
            dlg.destroy()
            if messagebox.askyesno("Reimpresion lista",
                    f"✅ {len(ids)} conciliacion(es) integradas en:\n{destino}\n\n"
                    "¿Abrirlo ahora?", parent=self):
                try: os.startfile(destino)
                except Exception: pass
        self._rbtn(bf, "📊  Generar Excel", _generar, "accent").pack(side="left")
        self._rbtn(bf, "Cancelar", dlg.destroy, "ghost").pack(side="left", padx=(8, 0))
        dlg.update_idletasks()
        try:
            px, py = self.winfo_rootx(), self.winfo_rooty()
            dlg.geometry(f"620x460+{px + 110}+{py + 90}")
        except Exception:
            pass


    def _concil_ver_guardado(self):
        """Consulta en la BD las conciliaciones guardadas de la empresa seleccionada.
        Tiene boton Actualizar (re-consulta) para reflejar lo recien subido sin
        reabrir la ventana, y boton para ver el DETALLE."""
        rfc = self._concil_rfc_sel()
        if not rfc:
            messagebox.showwarning("Sin empresa",
                                   "Selecciona una empresa en el Paso 3 primero.")
            return
        nombre = buscar_por_rfc(rfc) or rfc
        dlg = tk.Toplevel(self); dlg.title("Conciliaciones")
        dlg.transient(self); dlg.configure(bg=T.SURFACE); dlg.resizable(True, True)
        try: dlg.grab_set()
        except Exception: pass
        tk.Label(dlg, text=f"Conciliaciones — {nombre}", bg=T.SURFACE,
                 fg=T.TEXT, font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=16, pady=(14, 2))
        lbl_info = tk.Label(dlg, text="", bg=T.SURFACE, fg=T.TEXT_MUTED, font=FONT_SMALL)
        lbl_info.pack(anchor="w", padx=16, pady=(0, 8))

        cont = tk.Frame(dlg, bg=T.SURFACE); cont.pack(fill="both", expand=True, padx=12)
        cols = ("id", "periodo", "meses", "creado", "filas", "conc")
        tv = ttk.Treeview(cont, columns=cols, show="headings", height=12)
        for c, txt, w, anc in (("id", "ID", 60, "center"), ("periodo", "Año", 70, "center"),
                               ("meses", "Mes(es)", 150, "w"),
                               ("creado", "Guardada", 150, "center"),
                               ("filas", "Filas", 70, "e"), ("conc", "Conciliadas", 90, "e")):
            tv.heading(c, text=txt); tv.column(c, width=w, anchor=anc)
        sb = ttk.Scrollbar(cont, orient="vertical", command=tv.yview)
        tv.configure(yscrollcommand=sb.set)
        tv.pack(side="left", fill="both", expand=True); sb.pack(side="left", fill="y")

        def _recargar():
            try:
                import persistencia as _persist
                p = _persist.FirebirdPersistencia(); p.inicializar()
                try:
                    concs = p.leer_conciliaciones(rfc)
                finally:
                    p.cerrar()
            except Exception as e:
                messagebox.showerror("Error de base de datos", str(e), parent=dlg)
                return
            for it in tv.get_children():
                tv.delete(it)
            for c in concs:
                _cr = str(c.get("creado") or "")[:19]
                tv.insert("", "end", iid=str(c["id"]),
                          values=(c["id"], c.get("periodo") or "—", c.get("meses") or "—",
                                  _cr, c["n_filas"], c["n_conciliadas"]))
            lbl_info.config(
                text=f"RFC: {rfc}   ·   {len(concs)} conciliacion(es) en la base"
                     + ("" if concs else "   ·   (aun no hay conciliaciones guardadas)"))

        bf = tk.Frame(dlg, bg=T.SURFACE); bf.pack(fill="x", padx=16, pady=12)
        def _detalle():
            sel = tv.selection()
            if not sel:
                messagebox.showinfo("Selecciona", "Elige una conciliacion de la lista.",
                                    parent=dlg)
                return
            self._concil_ver_detalle(int(sel[0]), nombre, dlg)
        def _eliminar():
            sel = tv.selection()
            if not sel:
                messagebox.showinfo("Selecciona",
                                    "Elige una conciliacion de la lista.", parent=dlg)
                return
            cid = int(sel[0])
            vals = tv.item(sel[0])["values"]
            _per = vals[1] if len(vals) > 1 else ""
            _mes = vals[2] if len(vals) > 2 else ""
            if not messagebox.askyesno(
                    "Eliminar conciliación",
                    f"¿Eliminar la conciliación #{cid} ({_per} · {_mes})?\n\n"
                    "Se borrarán sus filas guardadas y su historial. Esta acción "
                    "NO se puede deshacer.", parent=dlg, icon="warning"):
                return
            try:
                import persistencia as _persist
                p = _persist.FirebirdPersistencia(); p.inicializar()
                try:
                    n = p.eliminar_conciliacion(cid)
                finally:
                    p.cerrar()
            except Exception as e:
                messagebox.showerror("Error de base de datos", str(e), parent=dlg)
                return
            messagebox.showinfo("Eliminada",
                                f"Conciliación #{cid} eliminada ({n} fila[s]).",
                                parent=dlg)
            _recargar()
        tv.bind("<Double-1>", lambda _e: _detalle())
        self._rbtn(bf, "📄  Ver detalle", _detalle, "accent").pack(side="left")
        self._rbtn(bf, "📤  Subir Excel revisado",
                   lambda: self._concil_subir_excel(despues=_recargar),
                   "primary").pack(side="left", padx=(8, 0))
        self._rbtn(bf, "🔄  Actualizar", _recargar, "accent").pack(side="left", padx=(8, 0))
        self._rbtn(bf, "🗑  Eliminar", _eliminar, "danger").pack(side="left", padx=(8, 0))
        self._rbtn(bf, "Cerrar", dlg.destroy, "ghost").pack(side="left", padx=(8, 0))
        _recargar()
        dlg.update_idletasks()
        try:
            px, py = self.winfo_rootx(), self.winfo_rooty()
            dlg.geometry(f"+{px + 120}+{py + 110}")
        except Exception:
            pass


    def _concil_ver_detalle(self, conciliacion_id, empresa="", parent_win=None):
        """Detalle por movimiento de una conciliacion. Ventana REDIMENSIONABLE con
        botones nativos de maximizar/minimizar, MODAL sobre la lista (se cierra
        primero el detalle) y con scroll horizontal funcional."""
        try:
            import persistencia as _persist
            p = _persist.FirebirdPersistencia(); p.inicializar()
            try:
                filas = p.leer_filas(int(conciliacion_id))
            finally:
                p.cerrar()
        except Exception as e:
            messagebox.showerror("Error de base de datos", str(e), parent=self)
            return

        # NO transient -> conserva los botones de maximizar/minimizar de Windows.
        dlg = tk.Toplevel(self); dlg.title(f"Detalle conciliacion #{conciliacion_id}")
        dlg.configure(bg=T.SURFACE); dlg.resizable(True, True); dlg.minsize(720, 400)

        def _cerrar():
            # Al cerrar el detalle, devolver el grab a la lista (para poder cerrarla).
            try: dlg.grab_release()
            except Exception: pass
            dlg.destroy()
            try:
                if parent_win is not None and parent_win.winfo_exists():
                    parent_win.grab_set()
            except Exception: pass
        dlg.protocol("WM_DELETE_WINDOW", _cerrar)
        # El detalle toma el grab: es modal sobre la lista -> se cierra primero.
        try: dlg.grab_set()
        except Exception: pass

        tk.Label(dlg, text=f"Detalle de la conciliacion #{conciliacion_id}",
                 bg=T.SURFACE, fg=T.TEXT, font=("Segoe UI", 12, "bold")
                 ).pack(anchor="w", padx=16, pady=(14, 2))
        _conc = sum(1 for f in filas if (f.uuid_cfdi or "").strip())
        tk.Label(dlg, text=f"{empresa}   ·   {len(filas)} filas   ·   {_conc} con CFDI",
                 bg=T.SURFACE, fg=T.TEXT_MUTED, font=FONT_SMALL).pack(anchor="w", padx=16, pady=(0, 8))

        cont = tk.Frame(dlg, bg=T.SURFACE); cont.pack(fill="both", expand=True, padx=12, pady=(0, 4))
        cols = ("id", "banco", "fecha", "concepto", "retiro", "deposito",
                "uuid", "tipo", "match")
        tv = ttk.Treeview(cont, columns=cols, show="headings", height=16)
        for c, txt, w, anc in (
                ("id", "ID", 55, "center"), ("banco", "Banco", 120, "w"),
                ("fecha", "Fecha", 95, "center"), ("concepto", "Concepto", 300, "w"),
                ("retiro", "Retiro", 100, "e"), ("deposito", "Depósito", 100, "e"),
                ("uuid", "UUID CFDI", 290, "w"), ("tipo", "Tipo", 80, "center"),
                ("match", "Tipo match", 140, "w")):
            tv.heading(c, text=txt)
            tv.column(c, width=w, minwidth=w, anchor=anc, stretch=False)
        tv.tag_configure("conc", background="#E8F5E9")
        tv.tag_configure("sin", background="#FDECEA")
        sbx = ttk.Scrollbar(cont, orient="horizontal", command=tv.xview)
        sby = ttk.Scrollbar(cont, orient="vertical", command=tv.yview)
        tv.configure(xscrollcommand=sbx.set, yscrollcommand=sby.set)
        tv.grid(row=0, column=0, sticky="nsew")
        sby.grid(row=0, column=1, sticky="ns")
        sbx.grid(row=1, column=0, sticky="ew")
        cont.rowconfigure(0, weight=1); cont.columnconfigure(0, weight=1)
        for f in filas:
            tag = "conc" if (f.uuid_cfdi or "").strip() else "sin"
            tv.insert("", "end", tags=(tag,), values=(
                f.id, f.banco, f.fecha, (f.concepto or "")[:160],
                f"{f.retiro:,.2f}" if f.retiro else "",
                f"{f.deposito:,.2f}" if f.deposito else "",
                f.uuid_cfdi or "", f.tipo or "", f.tipo_match or ""))

        bf = tk.Frame(dlg, bg=T.SURFACE); bf.pack(fill="x", padx=16, pady=12)
        self._rbtn(bf, "Cerrar", _cerrar, "ghost").pack(side="left")
        dlg.update_idletasks()
        try:
            dlg.geometry("1040x580+%d+%d" % (self.winfo_rootx() + 70, self.winfo_rooty() + 70))
        except Exception:
            pass


    def _concil_abrir_catalogos_win(self):
        """Ventana con los Catálogos COI (Plan de cuentas, Clientes/Proveedores,
        Grupos de RFC), en dos columnas de botones. Los datos viven en memoria."""
        if not _COI_OK:
            return
        win = tk.Toplevel(self)
        win.title("Catálogos COI — para generación de pólizas")
        win.configure(bg=T.SURFACE); win.transient(self); win.resizable(False, False)
        tk.Label(win, text="Catálogos COI (para generación de pólizas)",
                 bg=T.SURFACE, fg=T.TEXT, font=("Segoe UI", 12, "bold")
                 ).pack(anchor="w", padx=16, pady=(14, 2))
        tk.Label(win, text="Se usan para generar las pólizas COI tras conciliar.",
                 bg=T.SURFACE, fg=T.TEXT_MUTED, font=FONT_SMALL).pack(anchor="w", padx=16, pady=(0, 8))

        cd = tk.Frame(win, bg=T.SURFACE); cd.pack(fill="x", padx=8, pady=(0, 6))
        cd.columnconfigure(1, weight=1)
        cd.columnconfigure(2, uniform="btns", minsize=190)
        cd.columnconfigure(3, uniform="btns", minsize=130)

        tk.Label(cd, text="Plan de cuentas:", bg=T.SURFACE, fg=C_TEXT, font=FONT_HEAD
                 ).grid(row=0, column=0, sticky="w", padx=(12, 6), pady=(8, 5))
        self._concil_lbl_cuentas = tk.Label(cd, text="No cargado", bg=T.SURFACE, fg=C_RED, font=FONT_SMALL)
        self._concil_lbl_cuentas.grid(row=0, column=1, sticky="w", pady=(8, 5))
        self._rbtn(cd, "Importar Excel COI", self._concil_importar_cuentas, "primary", small=True
                   ).grid(row=0, column=2, sticky="ew", padx=6, pady=(8, 5))

        tk.Label(cd, text="Clientes / Proveedores:", bg=T.SURFACE, fg=C_TEXT, font=FONT_HEAD
                 ).grid(row=1, column=0, sticky="w", padx=(12, 6), pady=5)
        self._concil_lbl_terceros = tk.Label(cd, text="No cargado", bg=T.SURFACE, fg=C_RED, font=FONT_SMALL)
        self._concil_lbl_terceros.grid(row=1, column=1, sticky="w", pady=5)
        self._rbtn(cd, "Auto-poblar desde XMLs", self._concil_poblar_terceros, "primary", small=True
                   ).grid(row=1, column=2, sticky="ew", padx=6, pady=5)
        self._rbtn(cd, "Gestionar", self._concil_gestionar_catalogos, "ghost", small=True
                   ).grid(row=1, column=3, sticky="ew", padx=(0, 12), pady=5)

        if _GRUPOS_OK:
            tk.Label(cd, text="Grupos de RFC:", bg=T.SURFACE, fg=C_TEXT, font=FONT_HEAD
                     ).grid(row=2, column=0, sticky="w", padx=(12, 6), pady=(5, 10))
            self._concil_lbl_grupos = tk.Label(cd, text="0 grupos configurados", bg=T.SURFACE, fg=C_GRAY2, font=FONT_SMALL)
            self._concil_lbl_grupos.grid(row=2, column=1, sticky="w", pady=(5, 10))
            self._rbtn(cd, "Gestionar grupos", self._concil_abrir_grupos_rfc, "primary", small=True
                       ).grid(row=2, column=2, sticky="ew", padx=6, pady=(5, 10))

        # ── Subir Excel revisado (a la BD)  ·  Reimprimir → Pólizas (desde la BD)
        tk.Frame(win, bg=T.BORDER, height=1).pack(fill="x", padx=12, pady=(8, 0))
        _rip = tk.Frame(win, bg=T.SURFACE); _rip.pack(fill="x", padx=12, pady=(8, 2))
        self._rbtn(_rip, "🖨  Reimprimir → Pólizas COI",
                   self._concil_reimprimir_polizas, "accent"
                   ).pack(side="left", padx=(8, 0))
        self._rbtn(_rip, "📊  Reimprimir → Excel",
                   self._concil_reimprimir_excel, "primary"
                   ).pack(side="left", padx=(8, 0))
        self._rbtn(_rip, "🔎  Detalles de Conciliación",
                   self._concil_ver_guardado, "ghost"
                   ).pack(side="left", padx=(8, 0))
        tk.Label(_rip, text="Lee el .xlsm ya revisado (match por ID) y genera las "
                            "pólizas con tus correcciones.",
                 bg=T.SURFACE, fg=T.TEXT_MUTED, font=FONT_SMALL).pack(side="left", padx=10)

        tk.Button(win, text="Cerrar", command=win.destroy, bg=T.SURFACE_ALT, fg=T.TEXT_MUTED,
                  relief="flat", bd=0, padx=16, pady=6, cursor="hand2", font=FONT_SMALL
                  ).pack(anchor="e", padx=16, pady=(2, 12))

        self._concil_refrescar_catalogos()
        if _GRUPOS_OK:
            self._concil_refrescar_lbl_grupos()
        win.update_idletasks()
        try:
            px, py = self.winfo_rootx(), self.winfo_rooty()
            pw, ph = self.winfo_width(), self.winfo_height()
            w, h = win.winfo_width(), win.winfo_height()
            win.geometry(f"+{px + (pw - w) // 2}+{py + max(0, (ph - h) // 3)}")
        except Exception:
            pass

    def _concil_refresh_empresas(self):
        """Recarga el Combobox de empresas desde el catálogo del convertidor."""
        try:
            empresas = listar_empresas()   # [(rfc, nombre), ...]
            opciones = [f"{rfc}  —  {nombre}" for rfc, nombre in empresas]
            self._concil_emp_combo["values"] = opciones
            rfc_actual = self._concil_rfc.get()
            for op in opciones:
                if op.startswith(rfc_actual):
                    self._concil_emp_var.set(op)
                    break
            else:
                if opciones:
                    self._concil_emp_var.set(opciones[0])
        except Exception:
            pass

    def _concil_add_carpeta(self):
        carpeta = filedialog.askdirectory(title="Seleccionar carpeta con archivos XML")
        if carpeta and carpeta not in self._concil_carpetas:
            self._concil_carpetas.append(carpeta)
            self._concil_lst_xml.insert(tk.END, carpeta)
            self._concil_lbl_xml.config(
                text=f"{len(self._concil_carpetas)} carpeta(s)  |  XMLs pendientes de escanear")

    def _concil_add_xml(self):
        """Agrega archivos .xml INDIVIDUALES (sin carpeta) a la lista de XML.
        Se procesan igual que las carpetas (leer_cfdis_carpetas acepta ambos)."""
        rutas = filedialog.askopenfilenames(
            title="Seleccionar archivos XML (CFDI)",
            filetypes=[("XML (CFDI)", "*.xml"), ("Todos", "*.*")])
        _n = 0
        for _r in rutas:
            if _r and _r not in self._concil_carpetas:
                self._concil_carpetas.append(_r)
                self._concil_lst_xml.insert(tk.END, _r)
                _n += 1
        if _n:
            self._concil_lbl_xml.config(
                text=f"{len(self._concil_carpetas)} elemento(s)  |  XMLs pendientes de escanear")

    def _abrir_descarga_sat(self):
        if not _SAT_OK:
            messagebox.showerror(
                "Módulo no disponible",
                "No se pudo cargar sat_descarga.py o falta el paquete 'satcfdi'.\n"
                "Instala con: pip install satcfdi",
            )
            return
        rfc_sugerido = self._concil_rfc.get().strip().upper()
        _DialogDescargaSAT(self, rfc_sugerido, self._on_sat_descarga_completa)

    def _on_sat_descarga_completa(self, carpeta: Path, resultado):
        """Agrega la carpeta descargada del SAT al Paso 1 de Conciliación CFDI."""
        ruta = str(carpeta)
        if ruta not in self._concil_carpetas:
            self._concil_carpetas.append(ruta)
            self._concil_lst_xml.insert(tk.END, f"{ruta}  (SAT: {resultado.total_cfdi} CFDIs)")
            self._concil_lbl_xml.config(
                text=f"{len(self._concil_carpetas)} carpeta(s)  |  XMLs pendientes de escanear")

    def _concil_clear_carpetas(self):
        self._concil_carpetas.clear()
        self._concil_lst_xml.delete(0, tk.END)
        self._concil_lbl_xml.config(text="0 carpetas  |  0 XMLs cargados")

    def _concil_add_pdf(self):
        paths = filedialog.askopenfilenames(
            title="Seleccionar estados de cuenta PDF",
            filetypes=[("PDF", "*.pdf"), ("Todos", "*.*")])
        existing = {d["path"] for d in self._concil_pdfs}
        for p in paths:
            if p not in existing:
                self._concil_pdfs.append({
                    "path": p, "nombre": Path(p).name,
                    "banco": "—", "movs": 0, "info": None,
                    "estado": "listo",
                })
        self._concil_refresh_pdfs()

    def _concil_clear_pdfs(self):
        self._concil_pdfs.clear()
        self._concil_refresh_pdfs()

    def _concil_refresh_pdfs(self):
        for w in self._concil_pdf_inner.winfo_children():
            w.destroy()
        for i, d in enumerate(self._concil_pdfs):
            movs_txt = str(d["movs"]) if d["movs"] else "—"
            _rbg = C_WHITE if i % 2 else T.SURFACE_ALT
            row = tk.Frame(self._concil_pdf_inner, bg=_rbg); row.pack(fill="x")
            est = d.get("estado", "listo")
            _pbg = "#FADBD8" if est == "error" else T.ACCENT_SOFT
            _pfg = C_RED if est == "error" else T.ACCENT_INK
            _banco = d["banco"] or "—"
            if _HAS_CTK:
                ctk.CTkLabel(row, text=_banco, fg_color=_pbg, text_color=_pfg,
                             corner_radius=9, font=("Segoe UI", 8, "bold"), height=20
                             ).pack(side="left", padx=(8, 8), pady=4)
            else:
                tk.Label(row, text=_banco, bg=_pbg, fg=_pfg,
                         font=("Segoe UI", 8, "bold")).pack(side="left", padx=(8, 8), pady=4)
            tk.Label(row, text=d["nombre"], bg=_rbg, fg=C_TEXT,
                     font=FONT_SMALL, anchor="w").pack(side="left", fill="x", expand=True)
            tk.Label(row, text=movs_txt, bg=_rbg, fg=C_GRAY2,
                     font=("Segoe UI", 9, "bold")).pack(side="right", padx=10)
        self._concil_lbl_pdf.config(text=f"{len(self._concil_pdfs)} PDF(s)")

    def _concil_elegir_destino(self):
        hoy  = datetime.now().strftime("%Y%m%d")
        dest = filedialog.asksaveasfilename(
            title="Guardar conciliación como…",
            defaultextension=".xlsx",
            initialfile=f"Conciliacion_{hoy}.xlsx",
            filetypes=[("Excel", "*.xlsx")])
        if dest:
            self._concil_destino.set(dest)

    def _concil_log_write(self, msg: str, tag: str = ""):
        def _do():
            ts = datetime.now().strftime("%H:%M:%S")
            self._concil_log.config(state="normal")
            self._concil_log.insert("end", f"[{ts}] {msg}\n", tag)
            self._concil_log.see("end")
            self._concil_log.config(state="disabled")
        self.after(0, _do)

    def _concil_log_clear(self):
        self._concil_log.config(state="normal")
        self._concil_log.delete("1.0", tk.END)
        self._concil_log.config(state="disabled")

    def _concil_subir_xml(self):
        """Botón "Actualizar XML": SOLO sube/actualiza los XML (CFDI) en la BD sin
        duplicar (por UUID). No concilia (para eso está "Reconciliar")."""
        if not _CONCIL_OK:
            messagebox.showerror("Error", "Módulos de conciliación no disponibles."); return
        rfc = self._concil_rfc.get().strip().upper()
        if not rfc:
            messagebox.showwarning("RFC requerido", "Ingresa el RFC de tu empresa."); return
        if not self._concil_carpetas:
            messagebox.showwarning("Sin XMLs",
                "Selecciona al menos una carpeta con archivos XML (CFDIs)."); return
        self._concil_prog.start(10)
        self._concil_log_clear()
        threading.Thread(target=self._concil_subir_xml_run, args=(rfc,),
                         daemon=True).start()

    def _concil_subir_xml_run(self, rfc: str):
        try:
            self._concil_log_write(
                f"Escaneando {len(self._concil_carpetas)} carpeta(s) de XMLs…", "info")
            cfdis = leer_cfdis_carpetas(*self._concil_carpetas)
            recibidos, emitidos, descartados = clasificar_cfdis(cfdis, rfc)
            self._concil_log_write(
                f"  {len(cfdis)} CFDIs ({len(recibidos)} recibidos, {len(emitidos)} emitidos"
                + (f", {len(descartados)} descartados (RFC no coincide)" if descartados else "")
                + ")", "ok")
            self.after(0, lambda: self._concil_lbl_xml.config(
                text=f"{len(self._concil_carpetas)} carpeta(s)  |  {len(cfdis)} XMLs cargados"))
            import persistencia as _persist_x
            _cfdis_data = []
            for _c in list(recibidos) + list(emitidos):
                _cfdis_data.append(_persist_x.CFDIData(
                    uuid=getattr(_c, "uuid", "") or "", tipo=getattr(_c, "tipo", "") or "",
                    rfc_emisor=getattr(_c, "rfc_emisor", "") or "",
                    rfc_receptor=getattr(_c, "rfc_receptor", "") or "",
                    nombre_emisor=getattr(_c, "nombre_emisor", "") or "",
                    nombre_receptor=getattr(_c, "nombre_receptor", "") or "",
                    subtotal=float(getattr(_c, "subtotal", 0) or 0),
                    total=float(getattr(_c, "total", 0) or 0),
                    descuento=float(getattr(_c, "descuento", 0) or 0),
                    iva=float(getattr(_c, "iva16", 0) or 0),
                    iva_retenido=float(getattr(_c, "iva_retenido", 0) or 0),
                    isr_retenido=float(getattr(_c, "isr_retenido", 0) or 0),
                    clave_prod_serv=getattr(_c, "clave_prod_serv", "") or "",
                    fecha_emision=getattr(_c, "fecha_emision", "") or ""))
            _px = _persist_x.FirebirdPersistencia(); _px.inicializar()
            try:
                _st = _px.guardar_cfdis(_cfdis_data, solo_nuevos=True)
            finally:
                _px.cerrar()
            self._concil_log_write(
                f"  XML en BD: {_st.get('nuevos', 0)} nuevos, "
                f"{_st.get('omitidos', 0)} ya existian (omitidos).", "ok")
            self.after(0, lambda n=_st.get('nuevos', 0), o=_st.get('omitidos', 0):
                       messagebox.showinfo(
                "XML actualizados",
                f"XML nuevos agregados : {n}\n"
                f"Ya existían (omitidos) : {o}\n\n"
                "Usa \"Reconciliar (sin CFDI)\" para conciliar contra los estados "
                "de cuenta guardados."))
        except Exception as exc:
            import traceback
            self._concil_log_write(f"ERROR: {exc}", "err")
            self._concil_log_write(traceback.format_exc(), "err")
            self.after(0, lambda e=str(exc): messagebox.showerror("Error al actualizar XML", e))
        finally:
            self.after(0, self._concil_prog.stop)

    # ── Reconciliar (solo filas 'sin CFDI') con submenú de meses ──────────────
    _MESNOM_CONCIL = {1: "enero", 2: "febrero", 3: "marzo", 4: "abril", 5: "mayo",
                      6: "junio", 7: "julio", 8: "agosto", 9: "septiembre",
                      10: "octubre", 11: "noviembre", 12: "diciembre"}

    def _concil_reconciliar_dialog(self):
        """Botón "Reconciliar (sin CFDI)": abre un submenú para elegir qué meses de
        estados de cuenta re-conciliar vs qué meses de XML usar. Solo toca filas
        'sin CFDI' / no conciliadas; nunca las ya conciliadas."""
        if not _CONCIL_OK:
            messagebox.showerror("Error", "Módulos de conciliación no disponibles."); return
        rfc = self._concil_rfc.get().strip().upper()
        if not rfc:
            messagebox.showwarning("RFC requerido", "Ingresa el RFC de tu empresa."); return
        if not self._concil_carpetas:
            messagebox.showwarning("Sin XMLs",
                "Selecciona la(s) carpeta(s) de XML a conciliar."); return
        self._concil_prog.start(10)
        self._concil_log_clear()
        threading.Thread(target=self._concil_reconciliar_prep, args=(rfc,),
                         daemon=True).start()

    def _concil_reconciliar_prep(self, rfc: str):
        try:
            from conciliacion import _mes_fecha as _mf
            from collections import Counter
            self._concil_log_write("Escaneando XML y leyendo estados de cuenta guardados…", "info")
            cfdis = leer_cfdis_carpetas(*self._concil_carpetas)
            recibidos, emitidos, _desc = clasificar_cfdis(cfdis, rfc)
            xml_cnt = Counter()
            for _c in list(recibidos) + list(emitidos):
                _m = _mf(getattr(_c, "fecha_emision", "") or "")
                if _m:
                    xml_cnt[_m] += 1
            import persistencia as _persist_x
            _px = _persist_x.FirebirdPersistencia(); _px.inicializar()
            try:
                filas = _px.leer_filas_por_rfc(rfc)
            finally:
                _px.cerrar()
            bank_cnt = Counter(); bank_sin = Counter()
            for f in filas:
                _m = _mf(f.fecha)
                if not _m:
                    continue
                bank_cnt[_m] += 1
                if not (f.uuid_cfdi and f.tipo_match):
                    bank_sin[_m] += 1
            self.after(0, lambda: self._concil_reconciliar_build(
                rfc, recibidos, emitidos, dict(xml_cnt), dict(bank_cnt), dict(bank_sin)))
        except Exception as exc:
            import traceback
            self._concil_log_write(traceback.format_exc(), "err")
            self.after(0, lambda e=str(exc): messagebox.showerror("Error", e))
        finally:
            self.after(0, self._concil_prog.stop)

    def _concil_reconciliar_build(self, rfc, recibidos, emitidos, xml_cnt, bank_cnt, bank_sin):
        MN = self._MESNOM_CONCIL
        if not bank_cnt:
            messagebox.showwarning(
                "Sin estados de cuenta",
                "No hay estados de cuenta guardados en la BD para este RFC.\n\n"
                "Concilia primero con PDF (botón de conciliación) para tener "
                "movimientos bancarios que reconciliar."); return
        if not xml_cnt:
            messagebox.showwarning("Sin XML",
                "Las carpetas seleccionadas no tienen XML con fecha para conciliar."); return

        dlg = tk.Toplevel(self); dlg.title("Reconciliar — elegir meses")
        dlg.transient(self); dlg.configure(bg=T.SURFACE); dlg.resizable(False, True)
        try: dlg.grab_set()
        except Exception: pass
        tk.Label(dlg, text="Reconciliar filas 'sin CFDI'", bg=T.SURFACE, fg=T.TEXT,
                 font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=16, pady=(14, 2))
        tk.Label(dlg, text="Elige qué meses de estados de cuenta re-conciliar y qué "
                 "meses de XML usar. Solo se tocan las filas 'sin CFDI'.",
                 bg=T.SURFACE, fg=T.TEXT_MUTED, font=FONT_SMALL, justify="left",
                 wraplength=520).pack(anchor="w", padx=16, pady=(0, 10))

        cols = tk.Frame(dlg, bg=T.SURFACE); cols.pack(fill="both", expand=True, padx=12)
        # Columna izquierda: meses de ESTADOS DE CUENTA (banco)
        lf = tk.LabelFrame(cols, text=" Meses de estados de cuenta ", bg=T.SURFACE,
                           fg=T.TEXT, font=("Segoe UI", 9, "bold"))
        lf.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        rf = tk.LabelFrame(cols, text=" Meses de XML a conciliar ", bg=T.SURFACE,
                           fg=T.TEXT, font=("Segoe UI", 9, "bold"))
        rf.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        cols.columnconfigure(0, weight=1); cols.columnconfigure(1, weight=1)
        bank_vars = {}; xml_vars = {}
        for _m in sorted(bank_cnt):
            v = tk.BooleanVar(value=(bank_sin.get(_m, 0) > 0)); bank_vars[_m] = v
            tk.Checkbutton(lf, text=f"{MN.get(_m, _m)}  ({bank_sin.get(_m, 0)} sin CFDI)",
                           variable=v, bg=T.SURFACE, fg=T.TEXT, activebackground=T.SURFACE,
                           selectcolor=C_WHITE, font=FONT_SMALL, anchor="w"
                           ).pack(fill="x", anchor="w", padx=6, pady=1)
        for _m in sorted(xml_cnt):
            v = tk.BooleanVar(value=True); xml_vars[_m] = v
            tk.Checkbutton(rf, text=f"{MN.get(_m, _m)}  ({xml_cnt[_m]} CFDI)",
                           variable=v, bg=T.SURFACE, fg=T.TEXT, activebackground=T.SURFACE,
                           selectcolor=C_WHITE, font=FONT_SMALL, anchor="w"
                           ).pack(fill="x", anchor="w", padx=6, pady=1)

        bf = tk.Frame(dlg, bg=T.SURFACE); bf.pack(fill="x", padx=16, pady=12)

        def _go():
            sb = {m for m, v in bank_vars.items() if v.get()}
            sx = {m for m, v in xml_vars.items() if v.get()}
            if not sb:
                messagebox.showinfo("Elige meses",
                    "Marca al menos un mes de estado de cuenta.", parent=dlg); return
            if not sx:
                messagebox.showinfo("Elige meses",
                    "Marca al menos un mes de XML.", parent=dlg); return
            dlg.destroy()
            self._concil_prog.start(10)
            threading.Thread(target=self._concil_reconciliar_run,
                             args=(rfc, recibidos, emitidos, sb, sx), daemon=True).start()

        self._rbtn(bf, "🔗  Reconciliar", _go, "accent").pack(side="left")
        self._rbtn(bf, "Cancelar", dlg.destroy, "ghost").pack(side="left", padx=(8, 0))
        dlg.update_idletasks()
        try:
            px, py = self.winfo_rootx(), self.winfo_rooty()
            dlg.geometry(f"+{px + 120}+{py + 100}")
        except Exception:
            pass

    def _concil_reconciliar_run(self, rfc, recibidos, emitidos, bank_meses, xml_meses):
        try:
            from conciliacion import _mes_fecha as _mf
            MN = self._MESNOM_CONCIL
            rec = [c for c in recibidos if _mf(getattr(c, "fecha_emision", "") or "") in xml_meses]
            emi = [c for c in emitidos if _mf(getattr(c, "fecha_emision", "") or "") in xml_meses]
            self._concil_log_write(
                "Reconciliando: estados de cuenta "
                f"[{', '.join(MN.get(m, str(m)) for m in sorted(bank_meses))}] "
                f"vs XML [{', '.join(MN.get(m, str(m)) for m in sorted(xml_meses))}] …", "info")
            _fs = {fid for fid, var in self._concil_fase_vars.items() if var.get()}
            _grp = None
            if _GRUPOS_OK:
                _gs = _cargar_grupos(rfc); _ga = [g for g in _gs if g.activo]
                if _ga:
                    _grp = _ga
            import reimportar as _reimp_x
            _rst = _reimp_x.reconciliar_desde_bd(
                rfc, meses=list(bank_meses), fases=_fs, grupos=_grp,
                usuario=getattr(self, "_usuario", ""),
                recibidos=rec, emitidos=emi, log=self._concil_log_write,
                regenerar=True)
            _r = _rst.get("reconciliadas", 0)
            self._concil_log_write(
                f"  Re-conciliación: {_r} fila(s) 'sin CFDI' ahora conciliadas.", "ok")
            if _r == 0:
                self.after(0, lambda: messagebox.showinfo(
                    "Sin cambios",
                    "No hubo cambios: ninguna fila 'sin CFDI' de los meses elegidos "
                    "encontró su CFDI."))
            else:
                self.after(0, lambda r=_r: messagebox.showinfo(
                    "Reconciliación",
                    f"{r} fila(s) 'sin CFDI' ahora conciliadas."))
        except Exception as exc:
            import traceback
            self._concil_log_write(f"ERROR: {exc}", "err")
            self._concil_log_write(traceback.format_exc(), "err")
            self.after(0, lambda e=str(exc): messagebox.showerror("Error al reconciliar", e))
        finally:
            self.after(0, self._concil_prog.stop)


    def _concil_ejecutar(self):
        if not _CONCIL_OK:
            messagebox.showerror("Error", "Módulos de conciliación no disponibles."); return
        rfc = self._concil_rfc.get().strip().upper()
        if not rfc:
            messagebox.showwarning("RFC requerido", "Ingresa el RFC de tu empresa."); return
        if not self._concil_carpetas:
            messagebox.showwarning("Sin XMLs",
                "Selecciona al menos una carpeta con archivos XML (CFDIs)."); return
        if not self._concil_pdfs:
            messagebox.showwarning("Sin PDFs",
                "Agrega al menos un estado de cuenta PDF."); return
        destino = self._concil_destino.get().strip()
        if not destino:
            messagebox.showwarning("Destino",
                "Selecciona dónde guardar el archivo Excel de conciliación."); return

        self._concil_cancel = threading.Event()
        # El botón principal se convierte en DETENER (rojo) mientras corre.
        self._set_concil_btn(
            text="■  DETENER PROCESO", color=C_RED,
            command=self._concil_detener, state="normal")
        self._concil_prog.start(10)
        self._concil_log_clear()
        threading.Thread(
            target=self._concil_run,
            args=(rfc, destino),
            daemon=True).start()

    def _concil_detener(self):
        """Marca la cancelación; el hilo de conciliación se detiene en el
        siguiente checkpoint (entre PDFs o antes del cruce)."""
        if getattr(self, "_concil_cancel", None) is not None:
            self._concil_cancel.set()
        self._set_concil_btn(state="disabled", text="■  Deteniendo…")
        self._concil_log_write("⛔ Cancelando… (terminando el paso actual)", "warn")

    def _concil_cancelado(self) -> bool:
        ev = getattr(self, "_concil_cancel", None)
        return bool(ev is not None and ev.is_set())

    def _concil_run(self, rfc: str, destino: str):
        try:
            all_movs  = []
            last_info = None

            # ── PASO 1: Leer CFDIs XML ─────────────────────────────────────
            self._concil_log_write(
                f"Escaneando {len(self._concil_carpetas)} carpeta(s) de XMLs…", "info")
            cfdis = leer_cfdis_carpetas(*self._concil_carpetas)
            recibidos, emitidos, descartados = clasificar_cfdis(cfdis, rfc)
            self._concil_log_write(
                f"  {len(cfdis)} CFDIs  ({len(recibidos)} recibidos, "
                f"{len(emitidos)} emitidos"
                + (f", {len(descartados)} descartados (RFC no coincide)" if descartados else "")
                + ")", "ok")
            self._concil_lbl_xml.config(
                text=f"{len(self._concil_carpetas)} carpeta(s)  |  "
                     f"{len(cfdis)} XMLs cargados")

            # ── Filtro por FORMA DE PAGO (c_FormaPago) elegido en el GUI ───
            #   Solo aplica si hay selección. NO destructivo con la lógica
            #   fiscal: PPD y forma "99"/vacía se conservan (su forma real vive
            #   en el REP); se descartan únicamente los PUE cuya forma del
            #   comprobante NO está entre las elegidas.
            sel_fp = getattr(self, "_concil_formas_sel", None)
            if sel_fp:
                def _pasa_fp(c):
                    if getattr(c, "metodo_pago", "") == "PPD":
                        return True
                    fp = (getattr(c, "forma_pago", "") or "").strip()
                    if fp in ("", "99"):
                        return True
                    return fp in sel_fp
                _nr0, _ne0 = len(recibidos), len(emitidos)
                recibidos = [c for c in recibidos if _pasa_fp(c)]
                emitidos  = [c for c in emitidos  if _pasa_fp(c)]
                _quit = (_nr0 - len(recibidos)) + (_ne0 - len(emitidos))
                self._concil_log_write(
                    f"  Filtro Forma de pago {', '.join(sorted(sel_fp))}: "
                    f"se omitieron {_quit} CFDI PUE con otra forma  "
                    f"(PPD/REP/'99' conservados)", "info")

            # ── PASO 2: Leer PDFs y extraer movimientos (uno por banco) ───
            bancos_input  = []   # [(nombre_banco, [MovBanco, ...])]
            colores_banco = {}   # nombre_banco → color hex
            cuentas_banco = {}   # nombre_banco → numero de cuenta del estado
            _COLORES = ["1565C0","2E7D32","6A1B9A","BF360C",
                        "00695C","4527A0","558B2F","AD1457"]
            _nombres_usados: dict[str, int] = {}

            # ── Periodo seleccionado en el GUI (filtra qué estados de cuenta
            #    entran; los XML NO se filtran). ────────────────────────────
            try:
                _anio_sel = int(self._concil_anio.get())
            except (ValueError, AttributeError):
                _anio_sel = None
            # Filtro solo por AÑO (las casillas de meses se quitaron del GUI):
            # se concilian todos los meses del año elegido.
            _meses_sel = set(range(1, 13))

            # Aviso claro de qué año se está conciliando (certeza para el usuario).
            self._concil_log_write(
                f"► Conciliando AÑO {_anio_sel}", "info")

            import re as _re
            def _periodo_de(movs, info) -> tuple[int, int]:
                """(año, mes) predominante del estado de cuenta, del CONTENIDO del
                PDF: mes por el mes predominante de los movimientos (todos los
                bancos lo traen); año del encabezado `periodo` o de alguna fecha
                con año. Devuelve 0 en lo que no se pueda determinar."""
                from collections import Counter as _Counter
                from conciliacion import _mes_fecha as _mf
                _cnt = _Counter()
                for _m in movs:
                    _mm = getattr(_m, "mes", 0) or _mf(getattr(_m, "fecha", "") or "")
                    if _mm:
                        _cnt[_mm] += 1
                _mes = _cnt.most_common(1)[0][0] if _cnt else 0
                _anio = 0
                _m = _re.search(r"\b(20\d{2})\b", getattr(info, "periodo", "") or "")
                if _m:
                    _anio = int(_m.group(1))
                if not _anio:
                    for _mv in movs:
                        _m2 = _re.search(r"\b(20\d{2})\b", getattr(_mv, "fecha", "") or "")
                        if _m2:
                            _anio = int(_m2.group(1)); break
                return _anio, _mes

            # Pasada A: leer todos los PDFs y detectar su año (sin descartar aún).
            _leidos = []   # [(d, movs, info, banco_nombre, año_detectado)]
            for d in self._concil_pdfs:
                if self._concil_cancelado():
                    self._concil_log_write("⛔ Proceso detenido por el usuario.", "warn")
                    return
                self._concil_log_write(f"Leyendo PDF: {d['nombre']}", "info")
                try:
                    with PDF_LOCK:
                        extractor = detectar_banco(d["path"])
                    if extractor is None:
                        raise ValueError("Banco no reconocido")
                    d["banco"] = extractor.NOMBRE
                    with PDF_LOCK:
                        movs, info = extractor.extraer(d["path"])
                    if not movs:
                        raise ValueError("Sin movimientos en el PDF")
                    d["movs"]   = len(movs)
                    d["info"]   = info
                    d["estado"] = "ok"
                    last_info   = info
                    _pa, _pm = _periodo_de(movs, info)
                    _leidos.append((d, movs, info, extractor.NOMBRE, _pa))
                except Exception as e:
                    d["estado"] = "error"
                    self._concil_log_write(f"  ERROR: {e}", "err")
                self.after(0, self._concil_refresh_pdfs)

            # Filtro por AÑO A PRUEBA DE FALLOS: si NINGÚN estado de cuenta
            # coincide con el año elegido (o el año no se pudo detectar bien), se
            # concilian TODOS — el filtro NUNCA deja el resultado en cero.
            _coinciden = [t for t in _leidos
                          if not (_anio_sel and t[4] and t[4] != _anio_sel)]
            if _anio_sel and _leidos and not _coinciden:
                self._concil_log_write(
                    f"  ⚠ Ningún estado de cuenta coincide con el año {_anio_sel}; "
                    f"se concilian TODOS los cargados.", "warn")
                _usar = _leidos
            else:
                _usar = _coinciden
                _keep_ids = {id(t[0]) for t in _usar}
                for t in _leidos:
                    if id(t[0]) not in _keep_ids:
                        t[0]["estado"] = "omitido"
                        self._concil_log_write(
                            f"  ⏭ Omitido (año {t[4]} ≠ {_anio_sel}): {t[0]['nombre']}", "warn")

            # Pasada B: construir la entrada de conciliación con lo que se usará.
            for d, movs, info, _banco_nom, _pa in _usar:
                base = _banco_nom[:28]
                _nombres_usados[base] = _nombres_usados.get(base, 0) + 1
                nombre_banco = base if _nombres_usados[base] == 1 \
                               else f"{base} ({_nombres_usados[base]})"
                movs_banco = [mov_a_banco(m, anio=_anio_sel or 0) for m in movs]
                bancos_input.append((nombre_banco, movs_banco))
                all_movs.extend(movs)
                idx_color = len(bancos_input) - 1
                colores_banco[nombre_banco] = _COLORES[idx_color % len(_COLORES)]
                cuentas_banco[nombre_banco] = getattr(info, "numero_cuenta", "") or ""
                self._concil_log_write(
                    f"  [{nombre_banco}] {len(movs)} movimientos", "ok")
            self.after(0, self._concil_refresh_pdfs)

            if not bancos_input:
                raise ValueError("No se extrajeron movimientos de ningún PDF.")

            self._concil_log_write(
                f"Total: {len(bancos_input)} banco(s), {len(all_movs)} movimientos", "info")

            # ── PASO 3: Conciliar multi-banco ──────────────────────────────
            if self._concil_cancelado():
                self._concil_log_write("⛔ Proceso detenido por el usuario.", "warn")
                return
            self._concil_log_write("Ejecutando conciliacion multi-banco...", "info")
            grupos_activos = None
            if _GRUPOS_OK:
                gs = _cargar_grupos(rfc)
                activos = [g for g in gs if g.activo]
                if activos:
                    self._concil_log_write(
                        f"Grupos RFC: {len(activos)} activo(s) — Fase 5 habilitada", "info")
                    grupos_activos = activos
            fases_sel = {fid for fid, var in self._concil_fase_vars.items()
                         if var.get()}
            self._concil_log_write(
                "Fases activas: "
                + ", ".join(fid for fid, _ in FASES_CONCILIACION
                            if fid in fases_sel), "info")
            resultados = conciliar_multibanco(bancos_input, recibidos, emitidos,
                                              grupos=grupos_activos,
                                              fases=fases_sel,
                                              log=self._concil_log_write)

            total_conc = sum(len(r.conciliados) for r in resultados)
            total_pend = sum(len(r.sin_cfdi)    for r in resultados)
            total_rev  = sum(sum(1 for mc in r.conciliados if mc.revisar)
                             for r in resultados)
            pct = round(total_conc / len(all_movs) * 100, 1) if all_movs else 0

            for res in resultados:
                self._concil_log_write(
                    f"  {res.nombre}: {len(res.conciliados)} conciliados, "
                    f"{len(res.sin_cfdi)} sin CFDI", "ok")
            self._concil_log_write(
                f"  TOTAL: {total_conc} conciliados ({pct}%), "
                f"{total_pend} sin CFDI"
                + (f", {total_rev} requieren revision" if total_rev else ""),
                "ok" if not total_pend else "warn")

            # ── Guardar en BD (Firebird) y estampar el ID ESTABLE en cada fila ──
            #    para que el Excel muestre la columna A (ID). NO destructivo: si la
            #    BD falla, el reporte sale igual (sin IDs) y solo se avisa.
            try:
                self._concil_guardar_bd(rfc, resultados, recibidos, emitidos)
            except Exception as _e_bd:
                self._concil_log_write(
                    f"  ⚠ BD no disponible ({_e_bd}); el reporte sale sin IDs.", "warn")

            # ── PASO 4: Exportar Excel ─────────────────────────────────────
            self._concil_log_write("Generando Excel...", "info")
            # El export puede devolver .xlsm (papel de trabajo con macro) en vez
            # del .xlsx pedido — usar SIEMPRE la ruta real que retorna para log,
            # pólizas, mensaje y os.startfile.
            _meses_sel = [m for m, v in self._concil_meses_iva.items() if v.get()]
            destino = _exp_concil_multi(
                resultados      = resultados,
                todos_recibidos = recibidos,
                todos_emitidos  = emitidos,
                output_path     = destino,
                colores_banco   = colores_banco,
                rfc_empresa     = rfc,
                cuentas_banco   = cuentas_banco,
                meses_iva       = _meses_sel or None,
            ) or destino
            self._concil_log_write(f"Archivo generado:", "ok")
            self._concil_log_write(f"  {destino}", "ok")

            # Guardar estado para módulo de pólizas
            self._concil_ultimo = {
                "resultados": resultados,
                "recibidos":  recibidos,
                "emitidos":   emitidos,
                "destino":    destino,
                "rfc":        rfc,
            }

            def _habilitar_polizas():
                if _COI_OK and hasattr(self, "_concil_btn_polizas"):
                    _bp = self._concil_btn_polizas
                    _txt = f"Polizas COI  ->  {total_conc} movimientos listos"
                    if _HAS_CTK:
                        _bp.configure(state="normal", text=_txt,
                                      text_color="#FFFFFF", fg_color=T.ACCENT)
                    else:
                        _bp.config(state="normal", text=_txt, fg="#FFFFFF")
            self.after(0, _habilitar_polizas)

            self.after(0, lambda: messagebox.showinfo(
                "Conciliacion completa",
                f"Proceso terminado exitosamente.\n\n"
                f"Bancos procesados   : {len(resultados)}\n"
                f"Movimientos totales : {len(all_movs)}\n"
                f"Conciliados         : {total_conc} ({pct}%)\n"
                f"Sin CFDI            : {total_pend}\n"
                + (f"Requieren revision  : {total_rev}\n" if total_rev else "")
                + f"\n{destino}"))
            self.after(0, lambda: os.startfile(destino))

        except Exception as exc:
            import traceback
            self._concil_log_write(f"ERROR: {exc}", "err")
            self._concil_log_write(traceback.format_exc(), "err")
            self.after(0, lambda e=str(exc): messagebox.showerror("Error en conciliación", e))

        finally:
            self.after(0, self._concil_prog.stop)
            self.after(0, lambda: self._set_concil_btn(
                state="normal", color=C_GREEN, command=self._concil_ejecutar,
                text="⚡  LEER PDFs  →  CONCILIAR  →  EXPORTAR EXCEL"))


    # ── Catálogos COI ─────────────────────────────────────────────────────────

    def _concil_rfc_sel(self) -> str:
        return self._concil_rfc.get().strip().upper()

    def _concil_refrescar_catalogos(self):
        if not _COI_OK:
            return
        lc = self._wget("_concil_lbl_cuentas")
        lt = self._wget("_concil_lbl_terceros")
        if not lc and not lt:
            return  # la ventana de Catálogos COI no está abierta
        rfc = self._concil_rfc_sel()
        if not rfc:
            if lc: lc.config(text="Selecciona una empresa", fg=C_GRAY2)
            if lt: lt.config(text="Selecciona una empresa", fg=C_GRAY2)
            return

        cuentas  = cargar_catalogo(rfc)
        terceros = cargar_terceros(rfc)

        if lc:
            if cuentas:
                lc.config(text=f"{len(cuentas)} cuentas cargadas", fg="#2E7D32")
            else:
                lc.config(text="No cargado", fg=C_RED)

        if lt:
            if terceros:
                sin_cta = sum(1 for t in terceros if not t.cuenta_contable)
                txt = f"{len(terceros)} terceros"
                txt += f"  ({sin_cta} sin cuenta)" if sin_cta else "  (todos con cuenta)"
                col = "#E65100" if sin_cta else "#2E7D32"
                lt.config(text=txt, fg=col)
            else:
                lt.config(text="No cargado", fg=C_RED)

    def _concil_importar_cuentas(self):
        if not _COI_OK:
            return
        rfc = self._concil_rfc_sel()
        if not rfc:
            messagebox.showwarning("Sin empresa",
                                   "Selecciona una empresa en el Paso 3 primero.")
            return
        path = filedialog.askopenfilename(
            title="Seleccionar Excel de Catalogo de Cuentas COI",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")],
        )
        if not path:
            return
        try:
            ok, msg, _ = _imp_cuentas(rfc, path)
            if ok:
                messagebox.showinfo("Catalogo importado", msg)
            else:
                messagebox.showerror("Error al importar", msg)
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            self._concil_refrescar_catalogos()

    def _concil_poblar_terceros(self):
        if not _COI_OK:
            return
        rfc = self._concil_rfc_sel()
        if not rfc:
            messagebox.showwarning("Sin empresa",
                                   "Selecciona una empresa en el Paso 3 primero.")
            return
        if not self._concil_carpetas:
            messagebox.showwarning("Sin XMLs",
                                   "Agrega carpetas con XMLs en el Paso 1 primero.")
            return
        try:
            from cfdi_parser import leer_cfdis_carpetas, clasificar_cfdis
            cfdis = leer_cfdis_carpetas(*self._concil_carpetas)
            recibidos, emitidos, _ = clasificar_cfdis(cfdis, rfc)
            nc, np = _poblar_terceros(rfc, recibidos, emitidos)
            messagebox.showinfo(
                "Terceros actualizados",
                f"Nuevos clientes:    {nc}\n"
                f"Nuevos proveedores: {np}\n\n"
                "Usa 'Gestionar' para asignar cuentas contables.",
            )
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            self._concil_refrescar_catalogos()

    def _concil_gestionar_catalogos(self):
        if not _COI_OK:
            return
        rfc = self._concil_rfc_sel()
        if not rfc:
            messagebox.showwarning("Sin empresa",
                                   "Selecciona una empresa en el Paso 3 primero.")
            return
        u = self._concil_ultimo or {}
        VentanaPolizas(
            parent            = self,
            rfc_empresa       = rfc,
            todos_recibidos   = u.get("recibidos", []),
            todos_emitidos    = u.get("emitidos", []),
            resultados        = u.get("resultados", []),
            output_excel_path = u.get("destino", ""),
        )
        self.after(500, self._concil_refrescar_catalogos)

    def _concil_abrir_polizas(self):
        if not _COI_OK:
            return
        if not self._concil_ultimo:
            messagebox.showwarning("Sin conciliacion",
                                   "Ejecuta la conciliacion primero.")
            return
        u = self._concil_ultimo
        VentanaPolizas(
            parent            = self,
            rfc_empresa       = u["rfc"],
            todos_recibidos   = u["recibidos"],
            todos_emitidos    = u["emitidos"],
            resultados        = u["resultados"],
            output_excel_path = u["destino"],
        )
        self.after(500, self._concil_refrescar_catalogos)

    # ── Grupos de RFC ─────────────────────────────────────────────────────────

    def _concil_refrescar_lbl_grupos(self):
        lg = self._wget("_concil_lbl_grupos")
        if not _GRUPOS_OK or not lg:
            return  # la ventana de Catálogos COI no está abierta
        rfc = self._concil_rfc_sel()
        if not rfc:
            lg.config(text="Selecciona una empresa", fg=C_GRAY2)
            return
        gs = _cargar_grupos(rfc)
        n_act = sum(1 for g in gs if g.activo)
        total = len(gs)
        if total == 0:
            lg.config(text="0 grupos configurados", fg=C_GRAY2)
        elif n_act == 0:
            lg.config(text=f"{total} grupo(s) — todos inactivos", fg="#E65100")
        else:
            lg.config(text=f"{total} grupo(s)  ({n_act} activo(s))", fg="#4527A0")

    def _concil_abrir_grupos_rfc(self):
        if not _GRUPOS_OK:
            return
        rfc = self._concil_rfc_sel()
        if not rfc:
            messagebox.showwarning("Sin empresa",
                                   "Selecciona una empresa en el Paso 3 primero.")
            return
        VentanaGruposRFC(self, rfc_empresa=rfc,
                         on_close=self._concil_refrescar_lbl_grupos)


# ═══════════════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ═══════════════════════════════════════════════════════════════════════════════
def main():
    dlg = DialogLogin()
    dlg.mainloop()
    if not dlg.usuario:
        return
    app = App(usuario=dlg.usuario)
    app.mainloop()

if __name__ == "__main__":
    main()