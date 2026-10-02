"""
ventana_grupos_rfc.py — Diálogo de gestión de Grupos de RFC.

Permite crear, editar y eliminar grupos de RFCs que se usarán en la
Fase 5 de conciliación (multi-RFC).  Los RFCs disponibles se toman de
catalogo_terceros para el RFC de la empresa activa.
"""

from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk

from grupos_conciliacion import (
    GrupoConciliacion,
    agregar_grupo,
    cargar_grupos,
    eliminar_grupo,
    guardar_grupos,
)

try:
    from catalogo_terceros import cargar_terceros
    _TERCEROS_OK = True
except ImportError:
    _TERCEROS_OK = False

# ── Tema centralizado ─────────────────────────────────────────────────────────
from ui_theme import T, F, fit_geometry

C_BG    = T.BG
C_WHITE = T.SURFACE
C_ACCENT= T.ACCENT
C_GREEN = T.SUCCESS
C_RED   = T.DANGER
C_TEXT  = T.TEXT
C_GRAY  = T.BORDER_STRONG
C_GRAY2 = T.TEXT_MUTED
C_PURP  = "#4527A0"   # color de grupos (fijo, identidad propia)
FONT_N  = F.BODY
FONT_B  = F.BTN_SM
FONT_T  = F.TITLE
FONT_S  = F.SMALL


def _refresh_theme():
    global C_BG, C_WHITE, C_ACCENT, C_GREEN, C_RED, C_TEXT, C_GRAY, C_GRAY2
    global FONT_N, FONT_B, FONT_T, FONT_S
    C_BG    = T.BG;     C_WHITE = T.SURFACE;  C_ACCENT= T.ACCENT
    C_GREEN = T.SUCCESS; C_RED   = T.DANGER
    C_TEXT  = T.TEXT;   C_GRAY  = T.BORDER_STRONG; C_GRAY2 = T.TEXT_MUTED
    FONT_N  = F.BODY;   FONT_B  = F.BTN_SM;   FONT_T  = F.TITLE;  FONT_S  = F.SMALL

_COLORES_GRUPO = [
    "4527A0", "1565C0", "2E7D32", "BF360C",
    "00695C", "AD1457", "558B2F", "0277BD",
]


def _lbl(parent, text, font=None, fg=None, bg=None, **kw):
    return tk.Label(parent, text=text,
                    font=font if font is not None else F.BODY,
                    fg=fg   if fg   is not None else T.TEXT,
                    bg=bg   if bg   is not None else T.SURFACE,
                    **kw)


def _btn(parent, text, cmd, bg=None, fg=None, font=None, **kw):
    _bg = bg   if bg   is not None else T.ACCENT
    _fg = fg   if fg   is not None else T.ACCENT_FG
    _f  = font if font is not None else F.BTN_SM
    return tk.Button(parent, text=text, command=cmd,
                     bg=_bg, fg=_fg, font=_f,
                     relief="flat", cursor="hand2",
                     activebackground=T.ACCENT_SOFT, activeforeground=T.ACCENT,
                     bd=0, padx=8, pady=4, **kw)


class VentanaGruposRFC(tk.Toplevel):
    """
    Ventana modal para gestionar los grupos de RFC de conciliación.

    Parámetros
    ----------
    parent       : ventana padre
    rfc_empresa  : RFC de la empresa activa (para cargar terceros y guardar)
    on_close     : callback opcional llamado al cerrar (sin argumentos)
    """

    def __init__(self, parent, rfc_empresa: str, on_close=None):
        super().__init__(parent)
        _refresh_theme()
        self.title("Grupos de RFC — Conciliación Multi-RFC")
        self.configure(bg=C_BG)
        self.resizable(True, True)
        fit_geometry(self, 880, 560)
        self.grab_set()

        self._rfc_empresa = rfc_empresa
        self._on_close    = on_close
        self._grupos:     list[GrupoConciliacion] = cargar_grupos(rfc_empresa)
        self._terceros_r: list[tuple[str, str]] = []   # (rfc, nombre) proveedores
        self._terceros_e: list[tuple[str, str]] = []   # (rfc, nombre) clientes
        self._editando:   str | None = None   # nombre del grupo en edición
        # Estado de orden por tabla: (columna, descendente)
        self._sort = {"prov": ("rfc", False), "cli": ("rfc", False)}

        self._cargar_terceros()
        self._build_ui()
        self._refrescar_lista_grupos()
        self.protocol("WM_DELETE_WINDOW", self._cerrar)

    # ── Carga de terceros ─────────────────────────────────────────────────────

    def _cargar_terceros(self) -> None:
        if not _TERCEROS_OK:
            return
        try:
            ts = cargar_terceros(self._rfc_empresa)

            def _dedup(items):
                vistos, out = set(), []
                for t in items:
                    rfc = (t.rfc or "").strip().upper()
                    if rfc and rfc not in vistos:
                        vistos.add(rfc)
                        out.append((rfc, (t.nombre or "").strip()))
                return out

            self._terceros_r = _dedup([t for t in ts if t.tipo == "proveedor"])
            self._terceros_e = _dedup([t for t in ts if t.tipo == "cliente"])
        except Exception:
            pass

    # ── UI ────────────────────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        # Título
        tk.Label(self, text="Grupos de RFC para Conciliación Multi-RFC",
                 font=FONT_T, bg=C_BG, fg=C_TEXT
                 ).grid(row=0, column=0, sticky="ew", pady=(12, 4), padx=16)
        tk.Label(self,
                 text="Agrupa varios RFCs: la suma de sus facturas se cruzará contra un solo movimiento bancario.",
                 font=FONT_S, bg=C_BG, fg=C_GRAY2
                 ).grid(row=0, column=0, sticky="ew", padx=16)

        # Panel principal (izq: lista + der: formulario)
        main = tk.Frame(self, bg=C_BG)
        main.grid(row=1, column=0, sticky="nsew", padx=16, pady=(0, 8))
        main.columnconfigure(0, weight=1)
        main.columnconfigure(1, weight=2)
        main.rowconfigure(0, weight=1)

        self._build_panel_izq(main)
        self._build_panel_der(main)

        # Barra inferior
        bar = tk.Frame(self, bg=C_BG)
        bar.grid(row=2, column=0, sticky="ew", padx=16, pady=(0, 12))
        _btn(bar, "Cerrar", self._cerrar, bg=C_GRAY, fg=C_TEXT
             ).pack(side="right")
        self._lbl_status = tk.Label(bar, text="", font=FONT_S, bg=C_BG, fg=C_GRAY2)
        self._lbl_status.pack(side="left")

    def _build_panel_izq(self, parent) -> None:
        frm = tk.Frame(parent, bg=C_WHITE, relief="solid", bd=1)
        frm.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        frm.columnconfigure(0, weight=1)
        frm.rowconfigure(1, weight=1)

        _lbl(frm, "Grupos existentes", font=FONT_B, fg=C_PURP
             ).grid(row=0, column=0, columnspan=2, sticky="w", padx=10, pady=(8, 4))

        # Treeview de grupos
        cols = ("nombre", "rfcs", "activo")
        self._tree = ttk.Treeview(frm, columns=cols, show="headings",
                                  height=12, selectmode="browse")
        self._tree.heading("nombre",  text="Nombre")
        self._tree.heading("rfcs",    text="RFCs")
        self._tree.heading("activo",  text="Activo")
        self._tree.column("nombre",  width=130, anchor="w")
        self._tree.column("rfcs",    width=80,  anchor="center")
        self._tree.column("activo",  width=50,  anchor="center")
        self._tree.grid(row=1, column=0, sticky="nsew", padx=(8, 0), pady=(0, 4))

        sb = ttk.Scrollbar(frm, orient="vertical", command=self._tree.yview)
        sb.grid(row=1, column=1, sticky="ns", pady=(0, 4))
        self._tree.config(yscrollcommand=sb.set)
        self._tree.bind("<<TreeviewSelect>>", self._on_select_grupo)

        # Botones acción sobre lista
        btn_f = tk.Frame(frm, bg=C_WHITE)
        btn_f.grid(row=2, column=0, columnspan=2, sticky="ew", padx=8, pady=(0, 8))
        _btn(btn_f, "Editar", self._editar_grupo, bg=C_ACCENT
             ).pack(side="left", padx=(0, 4))
        _btn(btn_f, "Eliminar", self._eliminar_grupo, bg=C_RED
             ).pack(side="left")
        self._btn_toggle = _btn(btn_f, "Desactivar", self._toggle_activo, bg=C_GRAY, fg=C_TEXT)
        self._btn_toggle.pack(side="left", padx=(4, 0))

    def _build_panel_der(self, parent) -> None:
        frm = tk.Frame(parent, bg=C_WHITE, relief="solid", bd=1)
        frm.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        frm.columnconfigure(1, weight=1)
        frm.rowconfigure(3, weight=1)

        self._lbl_form_title = _lbl(frm, "Nuevo grupo", font=FONT_B, fg=C_PURP)
        self._lbl_form_title.grid(row=0, column=0,
                                  sticky="w", padx=10, pady=(8, 4))
        _lbl(frm, f"Empresa: {self._rfc_empresa}", font=FONT_S, fg=C_GRAY2
             ).grid(row=0, column=1, sticky="e", padx=(0, 10), pady=(8, 4))

        # Nombre del grupo
        _lbl(frm, "Nombre:", font=FONT_B).grid(row=1, column=0, sticky="w", padx=10, pady=4)
        self._nombre_var = tk.StringVar()
        tk.Entry(frm, textvariable=self._nombre_var, font=FONT_N,
                 bg=T.SURFACE, fg=T.TEXT, insertbackground=T.TEXT,
                 highlightthickness=1, highlightbackground=T.BORDER,
                 relief="flat",
                 ).grid(row=1, column=1, sticky="ew", padx=(0, 10), pady=4)

        # Tolerancia
        _lbl(frm, "Tolerancia MXN:", font=FONT_B).grid(row=2, column=0, sticky="w", padx=10, pady=4)
        self._tol_var = tk.StringVar(value="1.00")
        tk.Entry(frm, textvariable=self._tol_var, font=FONT_N, width=10,
                 bg=T.SURFACE, fg=T.TEXT, insertbackground=T.TEXT,
                 highlightthickness=1, highlightbackground=T.BORDER,
                 relief="flat",
                 ).grid(row=2, column=1, sticky="w", padx=(0, 10), pady=4)

        # ── Dos paneles de selección: Proveedores | Clientes ─────────────────
        dos_col = tk.Frame(frm, bg=C_WHITE)
        dos_col.grid(row=3, column=0, columnspan=2, sticky="nsew", padx=10, pady=(4, 4))
        dos_col.columnconfigure(0, weight=1)
        dos_col.columnconfigure(1, weight=1)
        dos_col.rowconfigure(2, weight=1)

        # — Proveedores (retiros) ——————————————————————————————————————————————
        _lbl(dos_col, "Proveedores — Retiros (pagamos)",
             font=FONT_B, fg="#1565C0"
             ).grid(row=0, column=0, sticky="w", pady=(0, 2), padx=(0, 4))

        self._filtro_prov_var = tk.StringVar()
        self._filtro_prov_var.trace_add("write", lambda *_: self._filtrar_proveedores())
        tk.Entry(dos_col, textvariable=self._filtro_prov_var, font=FONT_S,
                 bg=T.SURFACE, fg=T.TEXT_MUTED, insertbackground=T.TEXT,
                 highlightthickness=1, highlightbackground=T.BORDER, relief="flat",
                 ).grid(row=1, column=0, sticky="ew", pady=(0, 2), padx=(0, 4))

        prov_f = tk.Frame(dos_col, bg=C_WHITE)
        prov_f.grid(row=2, column=0, sticky="nsew", padx=(0, 4))
        prov_f.columnconfigure(0, weight=1)
        prov_f.rowconfigure(0, weight=1)
        self._tree_prov = self._crear_tree_terceros(prov_f, "prov")

        # — Clientes (depósitos) ———————————————————————————————————————————————
        _lbl(dos_col, "Clientes — Depósitos (cobramos)",
             font=FONT_B, fg="#2E7D32"
             ).grid(row=0, column=1, sticky="w", pady=(0, 2), padx=(4, 0))

        self._filtro_cli_var = tk.StringVar()
        self._filtro_cli_var.trace_add("write", lambda *_: self._filtrar_clientes())
        tk.Entry(dos_col, textvariable=self._filtro_cli_var, font=FONT_S,
                 bg=T.SURFACE, fg=T.TEXT_MUTED, insertbackground=T.TEXT,
                 highlightthickness=1, highlightbackground=T.BORDER, relief="flat",
                 ).grid(row=1, column=1, sticky="ew", pady=(0, 2), padx=(4, 0))

        cli_f = tk.Frame(dos_col, bg=C_WHITE)
        cli_f.grid(row=2, column=1, sticky="nsew", padx=(4, 0))
        cli_f.columnconfigure(0, weight=1)
        cli_f.rowconfigure(0, weight=1)
        self._tree_cli = self._crear_tree_terceros(cli_f, "cli")

        # Poblar ambas tablas
        self._refrescar_rfcs()

        # Botones guardar / limpiar
        btn_f2 = tk.Frame(frm, bg=C_WHITE)
        btn_f2.grid(row=4, column=0, columnspan=2, sticky="ew",
                    padx=10, pady=(0, 10))
        _btn(btn_f2, "Guardar grupo", self._guardar_grupo, bg=C_GREEN
             ).pack(side="left", padx=(0, 6))
        _btn(btn_f2, "Limpiar formulario", self._limpiar_form,
             bg=C_GRAY, fg=C_TEXT).pack(side="left")

    # ── Tablas de RFCs (RFC | Nombre, ordenables) ────────────────────────────

    def _crear_tree_terceros(self, parent, kind: str) -> ttk.Treeview:
        """Crea una tabla de 2 columnas (RFC | Nombre) con orden A-Z al clic."""
        cols = ("rfc", "nombre")
        tv = ttk.Treeview(parent, columns=cols, show="headings",
                          selectmode="extended", height=9)
        tv.heading("rfc",    text="RFC",
                   command=lambda: self._ordenar(kind, "rfc"))
        tv.heading("nombre", text="Nombre o Razón social",
                   command=lambda: self._ordenar(kind, "nombre"))
        tv.column("rfc",    width=115, anchor="w", stretch=False)
        tv.column("nombre", width=210, anchor="w", stretch=True)
        tv.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(parent, orient="vertical", command=tv.yview)
        sb.grid(row=0, column=1, sticky="ns")
        tv.config(yscrollcommand=sb.set)
        return tv

    def _ordenar(self, kind: str, col: str) -> None:
        """Alterna orden ascendente/descendente por la columna clicada."""
        col_actual, desc = self._sort[kind]
        desc = (not desc) if col_actual == col else False
        self._sort[kind] = (col, desc)
        if kind == "prov":
            self._filtrar_proveedores()
        else:
            self._filtrar_clientes()

    def _poblar_tree(self, tree, datos, filtro, kind) -> None:
        """Filtra, ordena y puebla la tabla; conserva la selección (iid=RFC)."""
        sel = set(tree.selection())
        tree.delete(*tree.get_children())

        f = (filtro or "").strip().upper()
        filas = [(rfc, nom) for (rfc, nom) in datos
                 if not f or f in rfc.upper() or f in nom.upper()]

        col, desc = self._sort[kind]
        idx = 0 if col == "rfc" else 1
        filas.sort(key=lambda t: t[idx].upper(), reverse=desc)

        flecha = " ▼" if desc else " ▲"
        tree.heading("rfc",    text="RFC" + (flecha if col == "rfc" else ""))
        tree.heading("nombre", text="Nombre o Razón social" + (flecha if col == "nombre" else ""))

        if not filas:
            tree.insert("", "end", iid="__vacio__",
                        values=("", "(sin registros)"))
            return
        for rfc, nom in filas:
            tree.insert("", "end", iid=rfc, values=(rfc, nom))
            if rfc in sel:
                tree.selection_add(rfc)

    def _refrescar_rfcs(self) -> None:
        """Puebla ambas tablas con proveedores y clientes respectivamente."""
        self._filtrar_proveedores()
        self._filtrar_clientes()

    def _filtrar_proveedores(self) -> None:
        filtro = self._filtro_prov_var.get() if hasattr(self, "_filtro_prov_var") else ""
        self._poblar_tree(self._tree_prov, self._terceros_r, filtro, "prov")

    def _filtrar_clientes(self) -> None:
        filtro = self._filtro_cli_var.get() if hasattr(self, "_filtro_cli_var") else ""
        self._poblar_tree(self._tree_cli, self._terceros_e, filtro, "cli")

    # ── Grupos ────────────────────────────────────────────────────────────────

    def _refrescar_lista_grupos(self) -> None:
        for row in self._tree.get_children():
            self._tree.delete(row)
        for g in self._grupos:
            activo_txt = "✔" if g.activo else "—"
            n_prov = len(g.rfcs_proveedores)
            n_cli  = len(g.rfcs_clientes)
            rfcs_txt = f"P:{n_prov} C:{n_cli}"
            self._tree.insert("", "end", iid=g.nombre,
                              values=(g.nombre, rfcs_txt, activo_txt))
        n = len(self._grupos)
        self._lbl_status.config(
            text=f"{n} grupo{'s' if n != 1 else ''} configurado{'s' if n != 1 else ''}."
        )

    def _on_select_grupo(self, _event=None) -> None:
        sel = self._tree.selection()
        activo = True
        if sel:
            g = next((x for x in self._grupos if x.nombre == sel[0]), None)
            if g:
                activo = g.activo
        lbl = "Activar" if not activo else "Desactivar"
        self._btn_toggle.config(text=lbl)

    def _editar_grupo(self) -> None:
        sel = self._tree.selection()
        if not sel:
            messagebox.showwarning("Selecciona un grupo", "Haz clic en un grupo de la lista.",
                                   parent=self)
            return
        g = next((x for x in self._grupos if x.nombre == sel[0]), None)
        if not g:
            return
        self._editando = g.nombre
        self._lbl_form_title.config(text=f"Editando: {g.nombre}")
        self._nombre_var.set(g.nombre)
        self._tol_var.set(f"{g.tolerancia:.2f}")
        self._refrescar_rfcs()

        # Seleccionar en la tabla de Proveedores (iid = RFC)
        rfcs_prov = {r.strip().upper() for r in g.rfcs_proveedores}
        self._tree_prov.selection_set(
            [iid for iid in self._tree_prov.get_children() if iid.upper() in rfcs_prov])

        # Seleccionar en la tabla de Clientes
        rfcs_cli = {r.strip().upper() for r in g.rfcs_clientes}
        self._tree_cli.selection_set(
            [iid for iid in self._tree_cli.get_children() if iid.upper() in rfcs_cli])

    def _guardar_grupo(self) -> None:
        nombre = self._nombre_var.get().strip()
        if not nombre:
            messagebox.showwarning("Nombre requerido",
                                   "Escribe un nombre para el grupo.", parent=self)
            return

        def _extraer(tree) -> list[str]:
            out: list[str] = []
            for iid in tree.selection():
                rfc = iid.strip().upper()
                if rfc and rfc != "__VACIO__" and rfc not in out:
                    out.append(rfc)
            return out

        rfcs_prov = _extraer(self._tree_prov)
        rfcs_cli  = _extraer(self._tree_cli)

        if not rfcs_prov and not rfcs_cli:
            messagebox.showwarning("Sin RFCs",
                                   "Selecciona al menos un RFC en Proveedores o Clientes.",
                                   parent=self)
            return

        try:
            tol = float(self._tol_var.get().replace(",", "."))
        except ValueError:
            tol = 1.00

        idx_color = len(self._grupos) % len(_COLORES_GRUPO)
        color = _COLORES_GRUPO[idx_color]

        grupo = GrupoConciliacion(
            nombre=nombre,
            rfcs_proveedores=rfcs_prov,
            rfcs_clientes=rfcs_cli,
            color_hex=color,
            activo=True,
            tolerancia=tol,
        )

        if self._editando:
            nombre_viejo = self._editando
            self._grupos = [grupo if g.nombre == nombre_viejo else g
                            for g in self._grupos]
            if not any(g.nombre == nombre_viejo for g in self._grupos):
                self._grupos.append(grupo)
        else:
            self._grupos = [g for g in self._grupos if g.nombre != nombre]
            self._grupos.append(grupo)

        guardar_grupos(self._rfc_empresa, self._grupos)
        self._editando = None
        self._limpiar_form()
        self._refrescar_lista_grupos()

    def _eliminar_grupo(self) -> None:
        sel = self._tree.selection()
        if not sel:
            messagebox.showwarning("Selecciona un grupo", "Haz clic en un grupo primero.",
                                   parent=self)
            return
        nombre = sel[0]
        if not messagebox.askyesno("Confirmar",
                                   f"¿Eliminar el grupo '{nombre}'?", parent=self):
            return
        self._grupos = [g for g in self._grupos if g.nombre != nombre]
        guardar_grupos(self._rfc_empresa, self._grupos)
        self._refrescar_lista_grupos()

    def _toggle_activo(self) -> None:
        sel = self._tree.selection()
        if not sel:
            return
        nombre = sel[0]
        for g in self._grupos:
            if g.nombre == nombre:
                g.activo = not g.activo
                break
        guardar_grupos(self._rfc_empresa, self._grupos)
        self._refrescar_lista_grupos()

    def _limpiar_form(self) -> None:
        self._editando = None
        self._lbl_form_title.config(text="Nuevo grupo")
        self._nombre_var.set("")
        self._tol_var.set("1.00")
        self._tree_prov.selection_remove(*self._tree_prov.selection())
        self._tree_cli.selection_remove(*self._tree_cli.selection())

    # ── Cierre ────────────────────────────────────────────────────────────────

    def _cerrar(self) -> None:
        self.grab_release()
        self.destroy()
        if self._on_close:
            self._on_close()
