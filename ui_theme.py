# -*- coding: utf-8 -*-
"""
ui_theme.py — Sistema de diseño centralizado para ConvertidorBancos v5.0
=========================================================================
• Detecta automáticamente el tema Windows (light / dark) via registro.
• Escala fuentes según el DPI del monitor principal.
• Configura ttk.Style globalmente (Treeview, Notebook, Scrollbar, etc.).
• Exporta alias de compatibilidad C_* / FONT_* para migración gradual.

Uso típico
----------
    import ui_theme

    # 1. Antes de crear la ventana Tk:
    ui_theme.enable_dpi()

    # 2. Después de crear root = tk.Tk():
    ui_theme.apply_theme(root)

    # 3. Usar tokens en los widgets:
    from ui_theme import T, F
    frame.configure(bg=T.BG)
    label.configure(font=F.BODY, fg=T.TEXT)

    # Los alias de compatibilidad siguen funcionando sin cambios:
    from ui_theme import C_BG, C_ACCENT, FONT_NORMAL   # ← igual que antes
"""

from __future__ import annotations
import sys

# ── Detección de tema Windows ──────────────────────────────────────────────────

def _is_dark_mode() -> bool:
    """Lee AppsUseLightTheme del registro. Retorna True si el SO está en dark mode."""
    if sys.platform != "win32":
        return False
    try:
        import winreg
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize",
        )
        value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
        winreg.CloseKey(key)
        return value == 0   # 0 = dark, 1 = light
    except Exception:
        return False


_DPI_ENABLED = False


def enable_dpi() -> None:
    """
    Activa DPI awareness a nivel proceso, de la forma más moderna disponible.
    Llamar ANTES de crear la ventana Tk para que el SO no escale en modo borroso
    y para que `winfo_screenwidth/height` devuelva PÍXELES REALES (si el proceso
    no es DPI-aware, Windows le miente y reporta la resolución virtualizada por
    el escalado: 1920×1080 al 150 % se ve como 1280×720).

    Estrategia en cascada (la primera que exista en el equipo gana):
      1. SetProcessDpiAwarenessContext(PER_MONITOR_AWARE_V2)  → Win10 1703+
      2. SetProcessDpiAwareness(2)  = Per-Monitor            → Win8.1+
      3. SetProcessDpiAwareness(1)  = System                 → Win8.1+
      4. SetProcessDPIAware()       = System (legacy)        → Vista+
    Así el comportamiento es consistente entre PCs con distinta versión de
    Windows y distinto factor de escala (125/150/200 %).
    """
    global _DPI_ENABLED
    if _DPI_ENABLED or sys.platform != "win32":
        _DPI_ENABLED = True
        return
    import ctypes
    # 1) Per-Monitor-V2 (contexto). El handle -4 = PER_MONITOR_AWARE_V2.
    try:
        ctypes.windll.user32.SetProcessDpiAwarenessContext(
            ctypes.c_void_p(-4))
        _DPI_ENABLED = True
        return
    except Exception:
        pass
    # 2/3) shcore.SetProcessDpiAwareness: intenta Per-Monitor, luego System.
    try:
        for level in (2, 1):
            try:
                ctypes.windll.shcore.SetProcessDpiAwareness(level)
                _DPI_ENABLED = True
                return
            except Exception:
                continue
    except Exception:
        pass
    # 4) Legacy (Vista+): siempre existe.
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass
    _DPI_ENABLED = True


def _dpi_scale() -> float:
    """Retorna el factor de escala DPI del monitor principal (1.0 = 96 DPI)."""
    if sys.platform != "win32":
        return 1.0
    try:
        import ctypes
        dc = ctypes.windll.user32.GetDC(0)
        dpi = ctypes.windll.gdi32.GetDeviceCaps(dc, 88)  # LOGPIXELSX
        ctypes.windll.user32.ReleaseDC(0, dc)
        return max(1.0, dpi / 96.0)
    except Exception:
        return 1.0


def work_area(win=None) -> tuple[int, int, int, int]:
    """Devuelve `(x, y, w, h)` del ÁREA DE TRABAJO (sin barra de tareas) del
    MONITOR donde está la ventana `win` (o donde está el cursor si `win` es
    None), no del monitor primario.

    `x, y` son el origen del monitor en el escritorio virtual: en un monitor
    secundario NO es (0,0), así que centrar con `x + (w - ancho)//2` coloca la
    ventana en el monitor correcto. En equipos con un solo monitor devuelve
    `(0, 0, ancho, alto)` igual que antes.

    Se consulta vía Win32 (MonitorFromWindow/MonitorFromPoint + GetMonitorInfo);
    si algo falla o no es Windows, cae de vuelta a `winfo_screenwidth/height`
    del monitor primario con origen (0,0).
    """
    if sys.platform == "win32":
        try:
            import ctypes
            from ctypes import wintypes

            class _RECT(ctypes.Structure):
                _fields_ = [("left", wintypes.LONG), ("top", wintypes.LONG),
                            ("right", wintypes.LONG), ("bottom", wintypes.LONG)]

            class _MONITORINFO(ctypes.Structure):
                _fields_ = [("cbSize", wintypes.DWORD), ("rcMonitor", _RECT),
                            ("rcWork", _RECT), ("dwFlags", wintypes.DWORD)]

            user32 = ctypes.windll.user32
            hmon = None
            if win is not None:
                try:
                    hwnd = int(win.winfo_id())
                    # winfo_id da el HWND del widget; subimos a la ventana raíz.
                    root = user32.GetAncestor(hwnd, 2)  # GA_ROOT
                    hmon = user32.MonitorFromWindow(
                        root or hwnd, 2)  # MONITOR_DEFAULTTONEAREST
                except Exception:
                    hmon = None
            if not hmon:
                pt = wintypes.POINT()
                user32.GetCursorPos(ctypes.byref(pt))
                hmon = user32.MonitorFromPoint(pt, 2)  # DEFAULTTONEAREST
            mi = _MONITORINFO()
            mi.cbSize = ctypes.sizeof(_MONITORINFO)
            if hmon and user32.GetMonitorInfoW(hmon, ctypes.byref(mi)):
                r = mi.rcWork
                w, h = r.right - r.left, r.bottom - r.top
                if w > 0 and h > 0:
                    return int(r.left), int(r.top), int(w), int(h)
        except Exception:
            pass
    # Fallback: monitor primario vía Tk, origen (0,0).
    if win is not None:
        try:
            return 0, 0, int(win.winfo_screenwidth()), int(win.winfo_screenheight())
        except Exception:
            pass
    return 0, 0, 1920, 1080


def screen_size(win=None) -> tuple[int, int]:
    """Resolución `(ancho, alto)` del área de trabajo del monitor donde está
    `win` (o el cursor). Atajo sobre `work_area` para código que solo necesita
    el tamaño (p. ej. topes `min(px, pantalla*fraccion)`)."""
    _x, _y, w, h = work_area(win)
    return w, h


def center(win, w: int, h: int, y_offset: int = -20) -> tuple[int, int]:
    """Dimensiona y centra `win` (w×h px) en el monitor donde está la ventana o
    el cursor, respetando el origen del monitor. `y_offset` sube un poco la
    ventana respecto al centro exacto. Devuelve `(x, y)` aplicados."""
    mx, my, sw, sh = work_area(win)
    x = mx + max(0, (sw - w) // 2)
    y = my + max(0, (sh - h) // 2 + y_offset)
    win.geometry(f"{w}x{h}+{x}+{y}")
    return x, y


_DPI_CACHE: float | None = None


def dpi_scale() -> float:
    """Factor de escala DPI, cacheado (el DPI del proceso no cambia tras
    enable_dpi()). 1.0 = 96 DPI, 1.5 = 150 %, etc."""
    global _DPI_CACHE
    if _DPI_CACHE is None:
        _DPI_CACHE = _dpi_scale()
    return _DPI_CACHE


def sx(px: int) -> int:
    """Escala un tamaño en px pensado a 96 DPI por el factor DPI actual.
    Usar en topes de geometria de ventanas (min/geometry/minsize) para que el
    alto/ancho crezca junto con las fuentes en pantallas de alta densidad; la
    fraccion de pantalla que suele acompanar al min() sigue limitando el tope a
    la resolucion real."""
    return int(round(px * dpi_scale()))


def fit_geometry(win, w_px: int, h_px: int,
                 wfrac: float = 0.92, hfrac: float = 0.90,
                 center: bool = True) -> tuple[int, int]:
    """Dimensiona `win` a un tamano base en px (a 96 DPI) ESCALADO por DPI y
    ACOTADO a la resolucion real: `min(sx(px), pantalla*fraccion)`. Asi la
    ventana crece con las fuentes en alta densidad pero nunca excede la pantalla
    (clave en laptops de baja resolucion o alta densidad). Devuelve (w, h)."""
    win.update_idletasks()
    mx, my, sw, sh = work_area(win)
    w = min(sx(w_px), int(sw * wfrac))
    h = min(sx(h_px), int(sh * hfrac))
    if center:
        x = mx + (sw - w) // 2
        y = my + max(0, (sh - h) // 2 - 20)
        win.geometry(f"{w}x{h}+{x}+{y}")
    else:
        win.geometry(f"{w}x{h}")
    return w, h


# ── Clase de tokens ────────────────────────────────────────────────────────────

class _Tokens:
    """
    Contenedor de tokens del tema activo.
    Los valores se asignan por apply_theme() según light/dark.
    """
    # NOTA DE DISEÑO — la paleta está alineada con la del reporte Excel
    # (exportar_conciliacion.py). App y reporte deben leerse como un mismo
    # producto: el mismo azul de marca, los mismos semáforos y los mismos 4
    # colores de nivel de confianza. Los semánticos son tinta legible sobre
    # fondo claro (≥4.5:1), no el tono saturado de una marca de gráfica.

    # Fondos
    BG:          str = "#EEF1F5"   # fondo de ventana (gris frío del rediseño)
    SURFACE:     str = "#FFFFFF"   # tarjetas, formularios
    SURFACE_ALT: str = "#F6F8FB"   # filas alternas en tablas

    # Acento primario — verde suave del rediseño (blanco encima ≈ 4.6:1)
    ACCENT:      str = "#2F7D57"   # verde de marca
    ACCENT_FG:   str = "#FFFFFF"   # texto sobre ACCENT
    # Texto secundario SOBRE la barra de acento (tinte claro para contraste).
    ACCENT_FG_2: str = "#CFE7D9"   # sobre ACCENT
    ACCENT_SOFT: str = "#E9F3EE"   # fondos de encabezado, chips suaves
    ACCENT_INK:  str = "#2A5C45"   # texto verde oscuro sobre chips suaves (pills)

    # Texto
    TEXT:        str = "#26251F"   # texto principal
    TEXT_MUTED:  str = "#5A594F"   # texto secundario / labels

    # Bordes
    BORDER:      str = "#E1E0D9"   # hairline
    BORDER_STRONG: str = "#C3C2B7" # bordes más marcados

    # Semánticos — mismos tintes/tintas que el Excel
    SUCCESS:     str = "#14571B"   # conciliado (sobre SUCCESS_BG: 7.3:1)
    SUCCESS_BG:  str = "#DCF0DC"
    DANGER:      str = "#A32118"   # sin CFDI / error (5.9:1)
    DANGER_BG:   str = "#FADCDA"
    WARN:        str = "#7A5200"   # requiere revisión (5.9:1)
    WARN_BG:     str = "#FBEBC9"

    # Nivel de confianza del cruce — los 4 hues validados del reporte
    TIER_DIRECTO:   str = "#1B5FAF"   # Fase 1
    TIER_COMPUESTO: str = "#A8480F"   # Fases 2/3/5
    TIER_COMPLEM:   str = "#0F7A57"   # Fase 4
    TIER_DEBIL:     str = "#4A3AA7"   # Fases 6/7

    # Módulos especiales
    COI:         str = "#33322D"   # gris oscuro — identidad módulo COI
    COI_SOFT:    str = "#ECECE8"   # fondo suave COI
    BLUE:        str = "#2A78D6"   # azul acción secundaria

    # Estados Treeview
    SELECTION:   str = "#D6E4F7"   # fila seleccionada
    SELECTION_FG: str = "#1A1A2E"  # texto en fila seleccionada

    # Dark mode variants (se aplican si _is_dark_mode() == True)
    # Dark NO es un volteo automático del light: cada paso se eligió para la
    # superficie oscura (los 4 hues de nivel son los pasos dark de la misma
    # paleta categórica).
    _DARK = {
        "BG":           "#0C121C",
        "SURFACE":      "#131C29",
        "SURFACE_ALT":  "#0F1723",
        "ACCENT":       "#3E9A6C",
        "ACCENT_FG":    "#FFFFFF",
        "ACCENT_FG_2":  "#CFE7D9",
        "ACCENT_SOFT":  "#17281F",
        "ACCENT_INK":   "#AEDFC6",
        "TEXT":         "#FFFFFF",
        "TEXT_MUTED":   "#C3C2B7",
        "BORDER":       "#2C2C2A",
        "BORDER_STRONG":"#383835",
        "SUCCESS":      "#0CA30C",
        "SUCCESS_BG":   "#12250F",
        "DANGER":       "#E66767",
        "DANGER_BG":    "#2B1413",
        "WARN":         "#FAB219",
        "WARN_BG":      "#2A2010",
        "TIER_DIRECTO":   "#3987E5",
        "TIER_COMPUESTO": "#D95926",
        "TIER_COMPLEM":   "#199E70",
        "TIER_DEBIL":     "#9085E9",
        "COI":          "#898781",
        "COI_SOFT":     "#232321",
        "BLUE":         "#3987E5",
        "SELECTION":    "#16283F",
        "SELECTION_FG": "#FFFFFF",
    }

    def activate_dark(self) -> None:
        for attr, val in self._DARK.items():
            setattr(self, attr, val)

    def activate_light(self) -> None:
        light = {k: v for k, v in self.__class__.__dict__.items()
                 if isinstance(v, str) and not k.startswith("_")}
        for attr, val in light.items():
            setattr(self, attr, val)


# Instancia global de tokens (se configura en apply_theme)
T = _Tokens()


# ── Escala tipográfica ─────────────────────────────────────────────────────────

class _Fonts:
    """
    Tokens de fuente. Los tamaños se ajustan al DPI del monitor.
    Usar siempre estas tuplas en los widgets: font=F.BODY, font=F.HEAD, etc.
    """
    # Se rellenan en _build_fonts()
    TITLE:  tuple = ("Segoe UI", 14, "bold")
    HEAD:   tuple = ("Segoe UI", 10, "bold")
    BODY:   tuple = ("Segoe UI", 9)
    SMALL:  tuple = ("Segoe UI", 8)
    MONO:   tuple = ("Consolas", 9)
    BTN:    tuple = ("Segoe UI", 10, "bold")
    BTN_SM: tuple = ("Segoe UI", 9, "bold")


F = _Fonts()


def _build_fonts(scale: float) -> None:
    """Ajusta los tamaños de fuente al DPI actual."""
    def _sz(base: int) -> int:
        # Tkinter aplica su propio escalado cuando DPI awareness está activo.
        # Solo forzamos escala si el factor es significativamente > 1.
        if scale >= 1.5:
            return round(base * 1.15)
        return base

    F.TITLE  = ("Segoe UI", _sz(14), "bold")
    F.HEAD   = ("Segoe UI", _sz(10), "bold")
    F.BODY   = ("Segoe UI", _sz(9))
    F.SMALL  = ("Segoe UI", _sz(8))
    F.MONO   = ("Consolas",  _sz(9))
    F.BTN    = ("Segoe UI", _sz(10), "bold")
    F.BTN_SM = ("Segoe UI", _sz(9),  "bold")


# ── Aplicar tema global ────────────────────────────────────────────────────────

def apply_theme(root, dark: bool | None = None) -> bool:
    """
    Configura los tokens y el ttk.Style para la raíz Tk dada.

    Parámetros
    ----------
    root : tk.Tk o tk.Toplevel
        La ventana raíz de la aplicación.
    dark : bool | None
        True = forzar dark, False = forzar light, None = autodetectar del SO.

    Retorna
    -------
    bool : True si se activó el modo oscuro.
    """
    if dark is None:
        # Por defecto CLARO, no lo que diga el SO.
        #
        # Motivo: los módulos de la app hacen `from ui_theme import C_WHITE, …`.
        # Un from-import liga el nombre en el módulo consumidor con el valor que
        # tenía al importar, así que `_refresh_module_compat()` (que actualiza
        # los atributos de ui_theme) NUNCA alcanza a esos nombres ya ligados:
        # los ~270 widgets tk se quedan en claro para siempre. Si aquí se
        # autodetectaba dark, solo los widgets ttk (Treeview, Notebook) se iban
        # a oscuro y el resultado era una app clara con paneles de lista negros.
        #
        # Un dark mode de verdad exige migrar esos widgets a leer T.* en vivo;
        # mientras tanto, forzar claro es lo consistente. `dark=True` sigue
        # disponible para cuando se haga esa migración.
        dark = False

    if dark:
        T.activate_dark()
    else:
        T.activate_light()

    scale = _dpi_scale()
    _build_fonts(scale)
    _apply_ttk_style(root, dark)

    return dark


def _apply_ttk_style(root, dark: bool) -> None:
    """Configura ttk.Style globalmente para todos los widgets ttk."""
    try:
        from tkinter import ttk
    except ImportError:
        return

    style = ttk.Style(root)

    # Base: usar 'clam' como punto de partida (más personalizable que 'default' o 'vista')
    try:
        style.theme_use("clam")
    except Exception:
        pass

    # ── Treeview ────────────────────────────────────────────────────────────
    style.configure(
        "Treeview",
        background=T.SURFACE,
        foreground=T.TEXT,
        fieldbackground=T.SURFACE,
        borderwidth=0,
        relief="flat",
        rowheight=22,
        font=F.BODY,
    )
    style.configure(
        "Treeview.Heading",
        background=T.SURFACE_ALT,
        foreground=T.TEXT_MUTED,
        font=F.BTN_SM,
        relief="flat",
        borderwidth=0,
        padding=(6, 4),
    )
    style.map(
        "Treeview",
        background=[("selected", T.SELECTION)],
        foreground=[("selected", T.SELECTION_FG)],
    )
    style.map(
        "Treeview.Heading",
        background=[("active", T.ACCENT_SOFT)],
        foreground=[("active", T.ACCENT)],
    )
    style.layout("Treeview", [
        ("Treeview.treearea", {"sticky": "nswe"}),
    ])

    # ── Notebook ─────────────────────────────────────────────────────────────
    style.configure(
        "TNotebook",
        background=T.BG,
        borderwidth=0,
        tabmargins=(0, 0, 0, 0),
    )
    style.configure(
        "TNotebook.Tab",
        background=T.BG,
        foreground=T.TEXT_MUTED,
        font=F.BODY,
        padding=(14, 7),
        borderwidth=0,
    )
    style.map(
        "TNotebook.Tab",
        background=[("selected", T.SURFACE), ("active", T.SURFACE)],
        foreground=[("selected", T.TEXT),    ("active", T.TEXT)],
        expand=[("selected", (0, 0, 0, 2))],
    )

    # ── Scrollbar ─────────────────────────────────────────────────────────────
    style.configure(
        "TScrollbar",
        background=T.BORDER,
        troughcolor=T.SURFACE,
        borderwidth=0,
        arrowsize=12,
        relief="flat",
    )
    style.map(
        "TScrollbar",
        background=[("active", T.ACCENT_SOFT), ("pressed", T.ACCENT)],
    )

    # ── Separator ─────────────────────────────────────────────────────────────
    style.configure("TSeparator", background=T.BORDER)

    # ── Entry / Combobox ───────────────────────────────────────────────────────
    style.configure(
        "TEntry",
        fieldbackground=T.SURFACE,
        foreground=T.TEXT,
        bordercolor=T.BORDER,
        lightcolor=T.BORDER,
        darkcolor=T.BORDER,
        insertcolor=T.ACCENT,
        selectbackground=T.ACCENT_SOFT,
        selectforeground=T.TEXT,
        padding=(6, 4),
    )
    style.map(
        "TEntry",
        bordercolor=[("focus", T.ACCENT)],
        lightcolor=[("focus", T.ACCENT)],
    )

    # ── Frame / Label ─────────────────────────────────────────────────────────
    style.configure("TFrame", background=T.BG)
    style.configure("TLabel", background=T.BG, foreground=T.TEXT, font=F.BODY)

    # Color de fondo de la ventana raíz
    root.configure(bg=T.BG)


# ── Helpers de widget ──────────────────────────────────────────────────────────

def make_btn(parent, text: str, command=None,
             bg: str | None = None,
             fg: str | None = None,
             font=None,
             padx: int = 12,
             pady: int = 6,
             **kw):
    """
    Botón plano con los defaults del tema activo.
    bg/fg/font son opcionales — por defecto usa T.ACCENT / T.ACCENT_FG / F.BTN.
    """
    import tkinter as tk
    _bg   = bg   if bg   is not None else T.ACCENT
    _fg   = fg   if fg   is not None else T.ACCENT_FG
    _font = font if font is not None else F.BTN
    return tk.Button(
        parent,
        text=text, command=command,
        bg=_bg, fg=_fg, font=_font,
        relief="flat", cursor="hand2", bd=0,
        padx=padx, pady=pady,
        activebackground=T.ACCENT_SOFT,
        activeforeground=T.ACCENT,
        **kw,
    )


def make_lbl(parent, text: str = "",
             font=None,
             fg: str | None = None,
             bg: str | None = None,
             **kw):
    """Label con defaults del tema activo."""
    import tkinter as tk
    return tk.Label(
        parent,
        text=text,
        font=font if font is not None else F.BODY,
        fg=fg   if fg   is not None else T.TEXT,
        bg=bg   if bg   is not None else T.SURFACE,
        **kw,
    )


def make_entry(parent, textvariable=None, width: int = 24, **kw):
    """Entry con estilo del tema activo."""
    import tkinter as tk
    return tk.Entry(
        parent,
        textvariable=textvariable,
        width=width,
        font=F.BODY,
        bg=T.SURFACE,
        fg=T.TEXT,
        insertbackground=T.ACCENT,
        selectbackground=T.ACCENT_SOFT,
        selectforeground=T.TEXT,
        relief="flat",
        bd=1,
        highlightthickness=1,
        highlightcolor=T.ACCENT,
        highlightbackground=T.BORDER,
        **kw,
    )


def make_card(parent, **kw):
    """Frame con apariencia de tarjeta (surface + borde)."""
    import tkinter as tk
    return tk.Frame(parent, bg=T.SURFACE, relief="flat", bd=0,
                    highlightthickness=1, highlightbackground=T.BORDER,
                    **kw)


def separator(parent, orient: str = "horizontal", **kw):
    """Línea divisora en el color del tema."""
    import tkinter as tk
    if orient == "horizontal":
        return tk.Frame(parent, bg=T.BORDER, height=1, **kw)
    return tk.Frame(parent, bg=T.BORDER, width=1, **kw)


def tag_colors_treeview(tree, dark: bool = False) -> None:
    """
    Configura los tags de color estándar en un Treeview para filas alternas
    y estados de conciliación.
    Llamar después de crear cada Treeview.
    """
    tree.tag_configure("odd",  background=T.SURFACE)
    tree.tag_configure("even", background=T.SURFACE_ALT)


# ── Alias de compatibilidad ────────────────────────────────────────────────────
# Mantienen la API anterior durante la migración gradual.
# Los módulos que importaban "from ConvertidorSantander import C_BG" pueden
# importar desde aquí sin cambios.

# Colores
@property
def _lazy(name): return property(lambda self: getattr(T, name))

# Se usan como módulo, así que asignamos directamente como funciones lambda
def _c(name): return lambda: getattr(T, name)

# Para uso como constantes de módulo (from ui_theme import C_BG):
# Estos son getters dinámicos — siempre reflejan el tema activo.
class _Compat:
    """Namespace de compatibilidad — mismo acceso que las constantes antiguas."""
    @property
    def C_BG(self):      return T.BG
    @property
    def C_WHITE(self):   return T.SURFACE
    @property
    def C_ACCENT(self):  return T.ACCENT
    @property
    def C_TEXT(self):    return T.TEXT
    @property
    def C_GRAY(self):    return T.BORDER_STRONG
    @property
    def C_GRAY2(self):   return T.TEXT_MUTED
    @property
    def C_GREEN(self):   return T.SUCCESS
    @property
    def C_RED(self):     return T.DANGER
    @property
    def C_RED_ERR(self): return T.DANGER
    @property
    def C_BLUE(self):    return T.BLUE
    @property
    def C_WARN(self):    return T.WARN
    @property
    def C_DARK(self):    return T.SURFACE_ALT
    @property
    def C_HDR_FG(self):  return T.ACCENT_FG
    @property
    def C_COI(self):     return T.COI

    # Fuentes (acceso como constante)
    @property
    def FONT_TITLE(self):  return F.TITLE
    @property
    def FONT_HEAD(self):   return F.HEAD
    @property
    def FONT_NORMAL(self): return F.BODY
    @property
    def FONT_SMALL(self):  return F.SMALL
    @property
    def FONT_MONO(self):   return F.MONO
    @property
    def FONT_BTN(self):    return F.BTN
    # Alias cortos usados en subventanas
    @property
    def FONT_N(self): return F.BODY
    @property
    def FONT_B(self): return F.BTN_SM
    @property
    def FONT_T(self): return F.TITLE
    @property
    def FONT_S(self): return F.SMALL


_compat = _Compat()

# Exportar los atributos directamente al nivel de módulo como propiedades
# no es posible en Python, pero los módulos que hacen:
#   from ui_theme import C_BG
# obtendrán el valor EN EL MOMENTO de la importación.
# Para que siempre reflejen el tema activo, los módulos deben hacer:
#   import ui_theme; ui_theme.T.BG
# o usar las funciones make_* que siempre leen T.* en tiempo de ejecución.
#
# Para compatibilidad inmediata (los módulos se migran poco a poco),
# exportamos los valores del tema por defecto (light) como constantes.
# Después de apply_theme(), los make_* usarán el tema correcto.

C_BG     = T.BG
C_WHITE  = T.SURFACE
C_ACCENT = T.ACCENT
C_TEXT   = T.TEXT
C_GRAY   = T.BORDER_STRONG
C_GRAY2  = T.TEXT_MUTED
C_GREEN  = T.SUCCESS
C_RED    = T.DANGER
C_RED_ERR= T.DANGER
C_BLUE   = T.BLUE
C_WARN   = T.WARN
C_DARK   = T.SURFACE_ALT
C_HDR_FG = T.ACCENT_FG
C_COI    = T.COI

FONT_TITLE  = F.TITLE
FONT_HEAD   = F.HEAD
FONT_NORMAL = F.BODY
FONT_SMALL  = F.SMALL
FONT_MONO   = F.MONO
FONT_BTN    = F.BTN
FONT_N      = F.BODY
FONT_B      = F.BTN_SM
FONT_T      = F.TITLE
FONT_S      = F.SMALL


def _refresh_module_compat() -> None:
    """
    Actualiza las constantes de módulo después de apply_theme().
    Llamar al final de apply_theme() si se necesitan las constantes como
    valores de módulo (no solo T.*).
    """
    import ui_theme as _m
    _m.C_BG     = T.BG
    _m.C_WHITE  = T.SURFACE
    _m.C_ACCENT = T.ACCENT
    _m.C_TEXT   = T.TEXT
    _m.C_GRAY   = T.BORDER_STRONG
    _m.C_GRAY2  = T.TEXT_MUTED
    _m.C_GREEN  = T.SUCCESS
    _m.C_RED    = T.DANGER
    _m.C_RED_ERR= T.DANGER
    _m.C_BLUE   = T.BLUE
    _m.C_WARN   = T.WARN
    _m.C_DARK   = T.SURFACE_ALT
    _m.C_HDR_FG = T.ACCENT_FG
    _m.C_COI    = T.COI
    _m.FONT_TITLE  = F.TITLE
    _m.FONT_HEAD   = F.HEAD
    _m.FONT_NORMAL = F.BODY
    _m.FONT_SMALL  = F.SMALL
    _m.FONT_MONO   = F.MONO
    _m.FONT_BTN    = F.BTN
    _m.FONT_N      = F.BODY
    _m.FONT_B      = F.BTN_SM
    _m.FONT_T      = F.TITLE
    _m.FONT_S      = F.SMALL


# Actualizar apply_theme para llamar a _refresh_module_compat
_orig_apply = apply_theme

def apply_theme(root, dark: bool | None = None) -> bool:  # noqa: F811
    result = _orig_apply(root, dark)
    _refresh_module_compat()
    return result
