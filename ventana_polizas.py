"""
ventana_polizas.py — Ventana de configuración y generación de Pólizas COI

Se abre desde VentanaConciliacion después de conciliar.
Tres pestañas:
  1. Catálogo de Cuentas  — importar/ver el plan de cuentas de COI
  2. Clientes y Proveedores — auto-poblar desde XMLs, asociar cuentas
  3. Generar Pólizas       — configurar cuentas clave y generar el reporte
"""

from __future__ import annotations

import os
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Optional

from catalogo_cuentas  import (CuentaContable, cargar_catalogo, guardar_catalogo,
                                importar_desde_excel, buscar_cuentas)
from catalogo_terceros import (Tercero, cargar_terceros, guardar_terceros,
                                poblar_desde_cfdis, auto_asociar_cuentas,
                                actualizar_cuenta)
from catalogo_prod_serv import (ReglaProdServ, cargar_reglas, guardar_reglas,
                                poblar_desde_cfdis as _poblar_prod_serv)
from polizas_coi       import generar_polizas, generar_polizas_diario

# ── Tema centralizado ─────────────────────────────────────────────────────────
from ui_theme import T, F, fit_geometry

C_BG    = T.BG
C_WHITE = T.SURFACE
C_ACCENT= T.ACCENT
C_GREEN = T.SUCCESS
C_RED   = T.DANGER
C_BLUE  = T.BLUE
C_TEXT  = T.TEXT
C_GRAY  = T.BORDER_STRONG
C_GRAY2 = T.TEXT_MUTED
C_COI   = T.COI
FONT_N  = F.BODY
FONT_B  = F.BTN_SM
FONT_T  = F.TITLE
FONT_S  = F.SMALL


def _refresh_theme():
    global C_BG, C_WHITE, C_ACCENT, C_GREEN, C_RED, C_BLUE, C_TEXT, C_GRAY, C_GRAY2, C_COI
    global FONT_N, FONT_B, FONT_T, FONT_S
    C_BG    = T.BG;     C_WHITE = T.SURFACE;  C_ACCENT= T.ACCENT
    C_GREEN = T.SUCCESS; C_RED   = T.DANGER;   C_BLUE  = T.BLUE
    C_TEXT  = T.TEXT;   C_GRAY  = T.BORDER_STRONG; C_GRAY2 = T.TEXT_MUTED
    C_COI   = T.COI
    FONT_N  = F.BODY;   FONT_B  = F.BTN_SM;   FONT_T  = F.TITLE;  FONT_S  = F.SMALL


def _btn(parent, text, cmd, bg=None, fg=None, font=None,
         padx=10, pady=5, **kw):
    _bg = bg   if bg   is not None else T.ACCENT
    _fg = fg   if fg   is not None else T.ACCENT_FG
    _f  = font if font is not None else F.BTN_SM
    return tk.Button(parent, text=text, command=cmd,
                     bg=_bg, fg=_fg, font=_f,
                     relief="flat", cursor="hand2",
                     activebackground=T.ACCENT_SOFT, activeforeground=T.ACCENT,
                     bd=0, padx=padx, pady=pady, **kw)


def _lbl(parent, text, font=None, fg=None, bg=None, **kw):
    return tk.Label(parent, text=text,
                    font=font if font is not None else F.BODY,
                    fg=fg   if fg   is not None else T.TEXT,
                    bg=bg   if bg   is not None else T.SURFACE,
                    **kw)


def _entry_var(parent, var, width=22, **kw):
    return tk.Entry(parent, textvariable=var, width=width,
                    font=FONT_N, bg=C_WHITE, fg=C_TEXT,
                    relief="solid", bd=1, **kw)


# ─────────────────────────────────────────────────────────────────────────────

class VentanaPolizas(tk.Toplevel):
    """
    Ventana Toplevel para configurar y generar pólizas COI.

    Parámetros
    ----------
    parent            : ventana padre
    rfc_empresa       : RFC de la empresa conciliada
    todos_recibidos   : CFDIs recibidos del lote conciliado
    todos_emitidos    : CFDIs emitidos del lote conciliado
    resultados        : list[ResultadoBanco] del motor de conciliación
    output_excel_path : ruta del Excel de conciliación ya generado
    """

    def __init__(self, parent,
                 rfc_empresa:       str,
                 todos_recibidos:   list,
                 todos_emitidos:    list,
                 resultados:        list,
                 output_excel_path: str = ""):
        super().__init__(parent)
        _refresh_theme()
        self.title(f"Pólizas COI — {rfc_empresa}")
        fit_geometry(self, 900, 680)
        self.configure(bg=C_BG)
        self.resizable(True, True)

        self._rfc          = rfc_empresa.strip().upper()
        self._recibidos    = todos_recibidos
        self._emitidos     = todos_emitidos
        self._resultados   = resultados
        self._output_path  = output_excel_path

        self._cuentas:  list[CuentaContable] = cargar_catalogo(self._rfc)
        self._terceros: list[Tercero]        = cargar_terceros(self._rfc)
        self._reglas:   list[ReglaProdServ]  = cargar_reglas(self._rfc)

        self._build_ui()
        self._refrescar_cuentas()
        self._refrescar_terceros()
        self._refrescar_reglas()

    # ── Construcción de UI ────────────────────────────────────────────────────

    def _build_ui(self):
        # Barra superior COI — identidad diferenciada del módulo de conciliación
        tk.Frame(self, bg=C_COI, height=3).pack(fill="x")
        hdr = tk.Frame(self, bg=C_COI, pady=10)
        hdr.pack(fill="x")
        tk.Label(hdr, text=f"Pólizas Contables COI  ·  {self._rfc}",
                 font=FONT_T, bg=C_COI, fg="#FFFFFF"
                 ).pack(side="left", padx=16)
        tk.Label(hdr, text="Configura catálogos y genera pólizas de ingreso/egreso",
                 font=FONT_S, bg=C_COI, fg="#90A4AE"
                 ).pack(side="left", padx=4)
        tk.Frame(self, bg=T.BORDER, height=1).pack(fill="x")

        # Notebook
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=10, pady=8)

        self._tab_cuentas   = tk.Frame(nb, bg=C_BG)
        self._tab_terceros  = tk.Frame(nb, bg=C_BG)
        self._tab_prod_serv = tk.Frame(nb, bg=C_BG)
        self._tab_generar   = tk.Frame(nb, bg=C_BG)

        nb.add(self._tab_cuentas,   text="  📋  Catálogo de Cuentas  ")
        nb.add(self._tab_terceros,  text="  👥  Clientes y Proveedores  ")
        nb.add(self._tab_prod_serv, text="  🏷  Productos / Servicios  ")
        nb.add(self._tab_generar,   text="  ⚙️  Generar Pólizas  ")

        self._build_tab_cuentas()
        self._build_tab_terceros()
        self._build_tab_prod_serv()
        self._build_tab_generar()

    # ══════════════════════════════════════════════════════════════════════════
    #  TAB 1: Catálogo de Cuentas
    # ══════════════════════════════════════════════════════════════════════════

    def _build_tab_cuentas(self):
        p = self._tab_cuentas
        p.columnconfigure(0, weight=1)

        # Barra de acciones
        bar = tk.Frame(p, bg=C_BG)
        bar.grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 4))
        _btn(bar, "📂  Importar desde Excel COI",
             self._importar_cuentas, bg=C_COI).pack(side="left")
        _btn(bar, "🔄  Recargar",
             self._refrescar_cuentas, bg=C_GRAY, fg=C_TEXT).pack(side="left", padx=(6, 0))

        # Buscador
        bus_f = tk.Frame(p, bg=C_BG)
        bus_f.grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 4))
        tk.Label(bus_f, text="Buscar:", font=FONT_S, bg=C_BG, fg=C_TEXT).pack(side="left")
        self._bus_cta_var = tk.StringVar()
        self._bus_cta_var.trace_add("write", lambda *_: self._filtrar_cuentas())
        tk.Entry(bus_f, textvariable=self._bus_cta_var, width=30,
                 font=FONT_N, bg=C_WHITE, fg=C_TEXT, relief="solid", bd=1
                 ).pack(side="left", padx=6)
        self._lbl_n_ctas = tk.Label(bus_f, text="", font=FONT_S, bg=C_BG, fg=C_GRAY2)
        self._lbl_n_ctas.pack(side="right")

        # Tabla de cuentas
        cols = ("cuenta", "nombre", "tipo", "nivel", "moneda")
        hdrs = ("No. Cuenta", "Nombre", "Tipo", "Nivel", "Moneda")
        widths = (120, 320, 60, 50, 70)
        frm = tk.Frame(p, bg=C_BG)
        frm.grid(row=2, column=0, sticky="nsew", padx=12, pady=(0, 8))
        frm.columnconfigure(0, weight=1)
        frm.rowconfigure(0, weight=1)
        p.rowconfigure(2, weight=1)

        self._tv_cuentas = ttk.Treeview(frm, columns=cols, show="headings",
                                         selectmode="browse")
        for c, h, w in zip(cols, hdrs, widths):
            self._tv_cuentas.heading(c, text=h)
            self._tv_cuentas.column(c, width=w, anchor="w")
        sb = ttk.Scrollbar(frm, orient="vertical",
                            command=self._tv_cuentas.yview)
        self._tv_cuentas.configure(yscrollcommand=sb.set)
        self._tv_cuentas.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")

    def _importar_cuentas(self):
        path = filedialog.askopenfilename(
            title="Seleccionar Excel de Catálogo de Cuentas COI",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")],
        )
        if not path:
            return
        ok, msg, cuentas = importar_desde_excel(self._rfc, path)
        if ok:
            self._cuentas = cuentas
            self._refrescar_cuentas()
            messagebox.showinfo("Catálogo importado", msg)
        else:
            messagebox.showerror("Error al importar", msg)

    def _refrescar_cuentas(self):
        self._cuentas = cargar_catalogo(self._rfc)
        self._filtrar_cuentas()

    def _filtrar_cuentas(self):
        termino = getattr(self, "_bus_cta_var", None)
        t = termino.get().strip() if termino else ""
        filtradas = buscar_cuentas(self._cuentas, t) if t else self._cuentas
        self._tv_cuentas.delete(*self._tv_cuentas.get_children())
        for c in filtradas:
            self._tv_cuentas.insert("", "end",
                                     values=(c.cuenta, c.nombre, c.tipo, c.nivel, c.moneda))
        n = len(filtradas)
        self._lbl_n_ctas.config(text=f"{n} cuenta(s)")

    # ══════════════════════════════════════════════════════════════════════════
    #  TAB 2: Clientes y Proveedores
    # ══════════════════════════════════════════════════════════════════════════

    def _build_tab_terceros(self):
        p = self._tab_terceros
        p.columnconfigure(0, weight=1)

        bar = tk.Frame(p, bg=C_BG)
        bar.grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 4))
        _btn(bar, "🔄  Auto-poblar desde XMLs",
             self._poblar_terceros, bg=C_BLUE).pack(side="left")
        _btn(bar, "🔗  Auto-asociar cuentas",
             self._auto_asociar, bg=C_COI).pack(side="left", padx=(6, 0))
        _btn(bar, "✏️  Editar cuenta",
             self._editar_cuenta_tercero, bg=C_GRAY, fg=C_TEXT).pack(side="left", padx=(6, 0))

        self._lbl_n_terc = tk.Label(bar, text="", font=FONT_S, bg=C_BG, fg=C_GRAY2)
        self._lbl_n_terc.pack(side="right")

        # Tabla de terceros
        cols   = ("rfc", "nombre", "tipo", "cuenta")
        hdrs   = ("RFC", "Nombre / Razón Social", "Tipo", "Cuenta Contable")
        widths = (130, 300, 90, 130)
        frm = tk.Frame(p, bg=C_BG)
        frm.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 8))
        frm.columnconfigure(0, weight=1)
        frm.rowconfigure(0, weight=1)
        p.rowconfigure(1, weight=1)

        self._tv_terceros = ttk.Treeview(frm, columns=cols, show="headings",
                                          selectmode="browse")
        for c, h, w in zip(cols, hdrs, widths):
            self._tv_terceros.heading(c, text=h)
            self._tv_terceros.column(c, width=w, anchor="w")

        # Colores por tipo
        self._tv_terceros.tag_configure("cliente",   foreground="#1565C0")
        self._tv_terceros.tag_configure("proveedor", foreground="#2E7D32")
        self._tv_terceros.tag_configure("sin_cuenta", foreground="#B71C1C")

        sb = ttk.Scrollbar(frm, orient="vertical",
                            command=self._tv_terceros.yview)
        self._tv_terceros.configure(yscrollcommand=sb.set)
        self._tv_terceros.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")

        self._tv_terceros.bind("<Double-1>", lambda e: self._editar_cuenta_tercero())

    def _poblar_terceros(self):
        if not self._recibidos and not self._emitidos:
            messagebox.showwarning(
                "Sin XMLs",
                "No hay CFDIs cargados. Concilia primero para tener XMLs disponibles.",
            )
            return
        nc, np = poblar_desde_cfdis(self._rfc, self._recibidos, self._emitidos)
        self._terceros = cargar_terceros(self._rfc)
        self._refrescar_terceros()
        messagebox.showinfo(
            "Terceros actualizados",
            f"Nuevos clientes: {nc}\nNuevos proveedores: {np}\n"
            f"Total en catálogo: {len(self._terceros)}",
        )

    def _auto_asociar(self):
        if not self._cuentas:
            messagebox.showwarning(
                "Sin catálogo de cuentas",
                "Importa primero el catálogo de cuentas en la pestaña anterior.",
            )
            return
        n = auto_asociar_cuentas(self._rfc, self._cuentas)
        self._terceros = cargar_terceros(self._rfc)
        self._refrescar_terceros()
        messagebox.showinfo(
            "Auto-asociación completada",
            f"Se asociaron cuentas contables a {n} tercero(s) por similitud de nombre.\n"
            "Revisa y ajusta manualmente los que quedaron sin cuenta.",
        )

    def _editar_cuenta_tercero(self):
        sel = self._tv_terceros.selection()
        if not sel:
            messagebox.showwarning("Selecciona un tercero",
                                   "Haz clic en un tercero de la tabla para editarlo.")
            return
        vals = self._tv_terceros.item(sel[0])["values"]
        rfc_t = str(vals[0])

        dlg = _DialogCuenta(self, rfc_t, str(vals[1]),
                             str(vals[3]), self._cuentas)
        self.wait_window(dlg)
        if dlg.resultado:
            actualizar_cuenta(self._rfc, rfc_t, dlg.resultado)
            self._terceros = cargar_terceros(self._rfc)
            self._refrescar_terceros()

    def _refrescar_terceros(self):
        self._tv_terceros.delete(*self._tv_terceros.get_children())
        for t in self._terceros:
            tag = t.tipo if t.cuenta_contable else "sin_cuenta"
            self._tv_terceros.insert("", "end",
                                      values=(t.rfc, t.nombre, t.tipo,
                                              t.cuenta_contable or "(sin cuenta)"),
                                      tags=(tag,))
        n = len(self._terceros)
        sin = sum(1 for t in self._terceros if not t.cuenta_contable)
        self._lbl_n_terc.config(
            text=f"{n} tercero(s)"
                 + (f" — {sin} sin cuenta" if sin else "")
        )

    # ══════════════════════════════════════════════════════════════════════════
    #  TAB 3: Productos / Servicios (ClaveProdServ → cuentas de resultado)
    # ══════════════════════════════════════════════════════════════════════════

    def _build_tab_prod_serv(self):
        p = self._tab_prod_serv
        p.columnconfigure(0, weight=1)
        p.rowconfigure(1, weight=1)

        # Barra superior
        bar = tk.Frame(p, bg=C_BG)
        bar.grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 4))
        _btn(bar, "➕  Nueva regla", self._nueva_regla_prod, bg=C_COI
             ).pack(side="left")
        _btn(bar, "✏️  Editar",      self._editar_regla_prod,
             bg=T.SURFACE, fg=T.TEXT, highlightthickness=1, highlightbackground=T.BORDER
             ).pack(side="left", padx=(6, 0))
        _btn(bar, "🗑  Eliminar selección", self._eliminar_regla_prod,
             bg=T.DANGER, fg=T.ACCENT_FG
             ).pack(side="left", padx=(6, 0))
        _btn(bar, "🧹  Limpiar todo", self._limpiar_todas_reglas,
             bg=T.SURFACE, fg=T.DANGER,
             highlightthickness=1, highlightbackground=T.DANGER
             ).pack(side="left", padx=(6, 0))
        _btn(bar, "📥  Importar desde XMLs", self._importar_claves_desde_xml,
             bg="#5b3a8a", fg=T.ACCENT_FG
             ).pack(side="left", padx=(12, 0))
        tk.Label(bar,
                 text="Ctrl+clic o Shift+clic para selección múltiple  ·  "
                      "Emitidos = Ingresos  ·  Recibidos = Gastos",
                 font=FONT_S, bg=C_BG, fg=C_GRAY2
                 ).pack(side="right")

        # Tabla de reglas (3 columnas: Prefijo, Tipo, Descripción, Cuenta Contable)
        cols   = ("prefijo", "tipo", "descripcion", "cuenta")
        hdrs   = ("Prefijo SAT", "Tipo", "Descripción / Concepto SAT", "Cuenta Contable")
        widths = (110, 90, 280, 200)
        frm = tk.Frame(p, bg=C_BG)
        frm.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 8))
        frm.columnconfigure(0, weight=1)
        frm.rowconfigure(0, weight=1)

        self._tv_reglas = ttk.Treeview(frm, columns=cols, show="headings",
                                        selectmode="extended")
        for c, h, w in zip(cols, hdrs, widths):
            self._tv_reglas.heading(c, text=h)
            self._tv_reglas.column(c, width=w, anchor="w")
        sb = ttk.Scrollbar(frm, orient="vertical", command=self._tv_reglas.yview)
        self._tv_reglas.configure(yscrollcommand=sb.set)
        self._tv_reglas.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        self._tv_reglas.bind("<Double-1>", lambda e: self._editar_regla_prod())

        # Tags de color por tipo
        self._tv_reglas.tag_configure("tag_e",
            background=T.ACCENT_SOFT, foreground=T.TEXT)
        self._tv_reglas.tag_configure("tag_r",
            background=T.DANGER_BG, foreground=T.TEXT)
        self._tv_reglas.tag_configure("tag_both",
            background=T.SURFACE, foreground=T.TEXT)
        self._tv_reglas.tag_configure("tag_pending",
            foreground=T.TEXT_MUTED)

    @staticmethod
    def _tipo_label(t: str) -> str:
        return {"E": "Emitido", "R": "Recibido"}.get(t, "Ambos")

    def _refrescar_reglas(self):
        if not hasattr(self, "_tv_reglas"):
            return
        self._tv_reglas.delete(*self._tv_reglas.get_children())
        for r in self._reglas:
            tipo_lbl = self._tipo_label(r.tipo)
            # Mostrar la cuenta relevante según tipo
            if r.tipo == "E":
                cuenta = r.cuenta_ingreso or "⚠ Sin cuenta (Ingresos)"
                tag = "tag_e" if r.cuenta_ingreso else "tag_pending"
            elif r.tipo == "R":
                cuenta = r.cuenta_gasto or "⚠ Sin cuenta (Gastos)"
                tag = "tag_r" if r.cuenta_gasto else "tag_pending"
            else:  # Ambos
                ing = r.cuenta_ingreso or "?"
                gas = r.cuenta_gasto   or "?"
                cuenta = f"Ing: {ing}  |  Gas: {gas}"
                tag = "tag_both"
            self._tv_reglas.insert("", "end", tags=(tag,),
                                    values=(r.prefijo, tipo_lbl,
                                            r.descripcion, cuenta))

    def _nueva_regla_prod(self):
        dlg = _DialogReglaProd(self, None, self._cuentas)
        self.wait_window(dlg)
        if dlg.resultado:
            self._reglas.append(dlg.resultado)
            guardar_reglas(self._rfc, self._reglas)
            self._refrescar_reglas()

    def _editar_regla_prod(self):
        sel = self._tv_reglas.selection()
        if not sel:
            messagebox.showwarning("Selecciona una regla",
                                   "Haz clic en una fila para editarla.")
            return
        if len(sel) > 1:
            messagebox.showinfo("Editar de una en una",
                                "Para editar selecciona solo una fila.\n"
                                "Para borrar varias usa el botón 'Eliminar selección'.")
            return
        vals        = self._tv_reglas.item(sel[0])["values"]
        prefijo_sel = str(vals[0])
        tipo_sel    = {"Emitido": "E", "Recibido": "R"}.get(str(vals[1]), "")
        regla = next((r for r in self._reglas
                      if r.prefijo == prefijo_sel and r.tipo == tipo_sel), None)
        if not regla:
            return
        dlg = _DialogReglaProd(self, regla, self._cuentas)
        self.wait_window(dlg)
        if dlg.resultado:
            self._reglas = [
                dlg.resultado
                if (r.prefijo == regla.prefijo and r.tipo == regla.tipo)
                else r
                for r in self._reglas
            ]
            guardar_reglas(self._rfc, self._reglas)
            self._refrescar_reglas()

    def _eliminar_regla_prod(self):
        sel = self._tv_reglas.selection()
        if not sel:
            messagebox.showwarning("Sin selección",
                                   "Selecciona una o más filas para eliminar.\n"
                                   "Usa Ctrl+clic o Shift+clic para selección múltiple.")
            return
        # Construir set de (prefijo, tipo) a eliminar
        a_borrar = set()
        for iid in sel:
            vals    = self._tv_reglas.item(iid)["values"]
            prefijo = str(vals[0])
            tipo    = {"Emitido": "E", "Recibido": "R"}.get(str(vals[1]), "")
            a_borrar.add((prefijo, tipo))

        n = len(a_borrar)
        msg = (f"¿Eliminar {n} regla(s) seleccionada(s)?"
               if n > 1 else
               f"¿Eliminar la regla '{next(iter(a_borrar))[0]}'?")
        if not messagebox.askyesno("Confirmar eliminación", msg):
            return

        self._reglas = [r for r in self._reglas
                        if (r.prefijo, r.tipo) not in a_borrar]
        guardar_reglas(self._rfc, self._reglas)
        self._refrescar_reglas()

    def _limpiar_todas_reglas(self):
        if not self._reglas:
            messagebox.showinfo("Catálogo vacío", "No hay reglas para eliminar.")
            return
        n = len(self._reglas)
        if not messagebox.askyesno(
            "Limpiar catálogo completo",
            f"¿Eliminar las {n} regla(s) del catálogo de Productos/Servicios?\n\n"
            "Esta acción no se puede deshacer.",
        ):
            return
        self._reglas = []
        guardar_reglas(self._rfc, self._reglas)
        self._refrescar_reglas()

    def _importar_claves_desde_xml(self):
        if not self._recibidos and not self._emitidos:
            messagebox.showinfo("Sin CFDIs",
                                "No hay CFDIs cargados. Carga los XMLs en la pestaña "
                                "Clientes y Proveedores primero.")
            return
        actualizadas, n_nuevas = _poblar_prod_serv(
            self._rfc,
            recibidos=list(self._recibidos),
            emitidos=list(self._emitidos),
            reglas_existentes=self._reglas,
        )
        if n_nuevas == 0:
            messagebox.showinfo("Sin cambios",
                                "Todos los ClaveProdServ encontrados ya están en el catálogo.")
            return
        self._reglas = actualizadas
        self._refrescar_reglas()
        messagebox.showinfo(
            "Importación completada",
            f"Se importaron {n_nuevas} código(s) SAT nuevos.\n\n"
            "Ahora asigna las cuentas contables a cada código haciendo doble clic "
            "sobre la fila o usando el botón Editar."
        )

    # ══════════════════════════════════════════════════════════════════════════
    #  TAB 4: Generar Pólizas
    # ══════════════════════════════════════════════════════════════════════════

    def _build_tab_generar(self):
        p = self._tab_generar
        p.columnconfigure(1, weight=1)

        def sep(row):
            tk.Frame(p, bg=C_GRAY, height=1).grid(
                row=row, column=0, columnspan=3, sticky="ew", padx=12, pady=6)

        def fila_cta(row, lbl, var, hint, parent=p):
            tk.Label(parent, text=lbl, font=FONT_B, bg=C_BG, fg=C_TEXT, anchor="e"
                     ).grid(row=row, column=0, sticky="e", padx=(16, 6), pady=3)
            ef = tk.Frame(parent, bg=C_BG)
            ef.grid(row=row, column=1, sticky="ew", padx=(0, 6), pady=3)
            ef.columnconfigure(0, weight=1)
            _entry_var(ef, var).grid(row=0, column=0, sticky="ew")
            _btn(ef, "...", lambda v=var: self._sel_cuenta(v),
                 bg=C_GRAY, fg=C_TEXT, padx=6).grid(row=0, column=1, padx=(4, 0))
            tk.Label(parent, text=hint, font=FONT_S, bg=C_BG, fg=C_GRAY2
                     ).grid(row=row, column=2, sticky="w", padx=(0, 12))

        # ── Sección 0: Tipo de pólizas ────────────────────────────────────────
        tk.Label(p, text="Tipo de pólizas a generar",
                 font=FONT_T, bg=C_BG, fg=C_COI
                 ).grid(row=0, column=0, columnspan=3, sticky="w", padx=16, pady=(12, 4))

        self._v_tipo_pol = tk.StringVar(value="bancario")
        tipos = [
            ("bancario", "Solo Flujo Bancario  (Ig Ingreso / Eg Egreso)  — requiere conciliación"),
            ("diario",   "Solo Diario Dr  — reconocimiento contable por CFDI (sin cruce bancario)"),
            ("ambas",    "Ambas  — genera Bancario y Diario en el mismo archivo"),
        ]
        for i, (val, txt) in enumerate(tipos):
            tk.Radiobutton(p, text=txt, variable=self._v_tipo_pol, value=val,
                           font=FONT_N, bg=C_BG, fg=C_TEXT,
                           command=self._toggle_secciones_diario
                           ).grid(row=1+i, column=0, columnspan=3, sticky="w", padx=32, pady=1)

        sep(4)

        # ── Sección 1: Numeración ─────────────────────────────────────────────
        tk.Label(p, text="Numeración de pólizas",
                 font=FONT_T, bg=C_BG, fg=C_COI
                 ).grid(row=5, column=0, columnspan=3, sticky="w", padx=16, pady=(2, 2))

        self._v_num_inicio = tk.StringVar(value="1")
        tk.Label(p, text="Número inicial:", font=FONT_B, bg=C_BG, fg=C_TEXT, anchor="e"
                 ).grid(row=6, column=0, sticky="e", padx=(16, 6), pady=3)
        _entry_var(p, self._v_num_inicio).grid(row=6, column=1, sticky="w", pady=3)
        tk.Label(p, text="La primera póliza tendrá este número",
                 font=FONT_S, bg=C_BG, fg=C_GRAY2
                 ).grid(row=6, column=2, sticky="w", padx=(0, 12))

        sep(7)

        # ── Sección 2: Cuentas Bancario ───────────────────────────────────────
        self._frm_bancario = tk.Frame(p, bg=C_BG)
        self._frm_bancario.grid(row=8, column=0, columnspan=3, sticky="ew")
        self._frm_bancario.columnconfigure(1, weight=1)

        tk.Label(self._frm_bancario, text="Cuentas — Flujo Bancario (Ig/Eg)",
                 font=FONT_T, bg=C_BG, fg="#1565C0"
                 ).grid(row=0, column=0, columnspan=3, sticky="w", padx=16, pady=(2, 2))
        tk.Label(self._frm_bancario,
                 text="Selecciona del catálogo de cuentas con el botón …",
                 font=FONT_S, bg=C_BG, fg=C_GRAY2
                 ).grid(row=1, column=0, columnspan=3, sticky="w", padx=16)

        self._v_cta_banco = tk.StringVar()
        self._v_cta_prov  = tk.StringVar()
        self._v_cta_cli   = tk.StringVar()
        self._v_cta_nom   = tk.StringVar()
        fila_cta(2, "Cuenta Banco:",        self._v_cta_banco,
                 "Ej. 1110-001-000  Bancos",         self._frm_bancario)
        fila_cta(3, "Default Proveedores:", self._v_cta_prov,
                 "Fallback CxP si no se localiza proveedor", self._frm_bancario)
        fila_cta(4, "Default Clientes:",    self._v_cta_cli,
                 "Fallback CxC si no se localiza cliente",   self._frm_bancario)
        fila_cta(5, "Gasto de Nómina:",     self._v_cta_nom,
                 "Para CFDIs tipo N — Ej. 6100-001-000",     self._frm_bancario)

        self._v_sin_cfdi = tk.BooleanVar(value=False)
        tk.Checkbutton(self._frm_bancario,
                       text="Incluir movimientos sin CFDI (póliza Dr con cuenta por definir)",
                       variable=self._v_sin_cfdi,
                       bg=C_BG, fg=C_TEXT, font=FONT_S, anchor="w"
                       ).grid(row=6, column=0, columnspan=3, sticky="w", padx=16, pady=(4, 2))

        # ── Sección 3: Cuentas Diario ─────────────────────────────────────────
        self._frm_diario = tk.Frame(p, bg=C_BG)
        self._frm_diario.grid(row=9, column=0, columnspan=3, sticky="ew")
        self._frm_diario.columnconfigure(1, weight=1)

        tk.Label(self._frm_diario, text="Cuentas — Pólizas de Diario Dr",
                 font=FONT_T, bg=C_BG, fg="#2E7D32"
                 ).grid(row=0, column=0, columnspan=3, sticky="w", padx=16, pady=(6, 2))
        tk.Label(self._frm_diario,
                 text="ClaveProdServ → cuentas de resultado: configura reglas en la pestaña 'Productos/Servicios'",
                 font=FONT_S, bg=C_BG, fg=C_GRAY2
                 ).grid(row=1, column=0, columnspan=3, sticky="w", padx=16)

        self._v_cta_cobrar      = tk.StringVar()
        self._v_cta_pagar       = tk.StringVar()
        self._v_cta_iva_tras    = tk.StringVar()
        self._v_cta_iva_acre    = tk.StringVar()
        self._v_cta_iva_ret     = tk.StringVar()
        self._v_cta_isr_ret     = tk.StringVar()
        self._v_cta_def_ingreso = tk.StringVar()
        self._v_cta_def_gasto   = tk.StringVar()

        fila_cta(2, "Cuentas x Cobrar:",      self._v_cta_cobrar,
                 "Clientes — activo circulante",       self._frm_diario)
        fila_cta(3, "Cuentas x Pagar:",       self._v_cta_pagar,
                 "Proveedores — pasivo circulante",    self._frm_diario)
        fila_cta(4, "IVA Trasladado:",         self._v_cta_iva_tras,
                 "Pasivo — IVA cobrado a clientes",    self._frm_diario)
        fila_cta(5, "IVA Acreditable:",        self._v_cta_iva_acre,
                 "Activo — IVA pagado a proveedores",  self._frm_diario)
        fila_cta(6, "IVA Retenido:",           self._v_cta_iva_ret,
                 "Retención IVA (ambos lados)",        self._frm_diario)
        fila_cta(7, "ISR Retenido:",           self._v_cta_isr_ret,
                 "Retención ISR (ambos lados)",        self._frm_diario)
        fila_cta(8, "Ingresos (fallback):",    self._v_cta_def_ingreso,
                 "Cuando ClaveProdServ no tiene regla", self._frm_diario)
        fila_cta(9, "Gastos (fallback):",      self._v_cta_def_gasto,
                 "Cuando ClaveProdServ no tiene regla", self._frm_diario)

        sep(10)

        # ── Destino del archivo COI ───────────────────────────────────────────
        tk.Label(p, text="Exportar archivo COI independiente (opcional):",
                 font=FONT_B, bg=C_BG, fg=C_TEXT
                 ).grid(row=11, column=0, columnspan=3, sticky="w", padx=16, pady=(2, 2))
        dest_f = tk.Frame(p, bg=C_BG)
        dest_f.grid(row=12, column=0, columnspan=3, sticky="ew", padx=16, pady=(0, 4))
        dest_f.columnconfigure(0, weight=1)
        self._v_destino_coi = tk.StringVar()
        tk.Entry(dest_f, textvariable=self._v_destino_coi,
                 font=FONT_S, bg=C_WHITE, fg=C_TEXT, relief="solid", bd=1
                 ).grid(row=0, column=0, sticky="ew")
        _btn(dest_f, "...", self._sel_destino_coi,
             bg=C_GRAY, fg=C_TEXT, padx=6).grid(row=0, column=1, padx=(4, 0))
        tk.Label(dest_f,
                 text="Si se deja vacío, solo se agrega la hoja al reporte de conciliación",
                 font=FONT_S, bg=C_BG, fg=C_GRAY2
                 ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(2, 0))

        # ── Botón generar ─────────────────────────────────────────────────────
        p.rowconfigure(13, weight=1)
        btn_f = tk.Frame(p, bg=C_BG)
        btn_f.grid(row=13, column=0, columnspan=3, sticky="s", pady=10)

        _btn(btn_f, "⚡  GENERAR PÓLIZAS",
             self._generar, bg=T.SUCCESS, fg=T.ACCENT_FG,
             font=F.BTN, padx=24, pady=10
             ).pack(side="left", padx=8)

        self._lbl_estado = tk.Label(btn_f, text="", font=FONT_S,
                                     bg=C_BG, fg=C_GRAY2)
        self._lbl_estado.pack(side="left")

        # Estado inicial de secciones
        self._toggle_secciones_diario()

    def _toggle_secciones_diario(self):
        """Muestra u oculta las secciones Bancario/Diario según el tipo elegido."""
        tipo = self._v_tipo_pol.get()
        if tipo == "diario":
            self._frm_bancario.grid_remove()
            self._frm_diario.grid()
        elif tipo == "bancario":
            self._frm_bancario.grid()
            self._frm_diario.grid_remove()
        else:  # ambas
            self._frm_bancario.grid()
            self._frm_diario.grid()

    def _sel_cuenta(self, var: tk.StringVar):
        """Abre un diálogo para seleccionar una cuenta del catálogo."""
        if not self._cuentas:
            messagebox.showwarning(
                "Sin catálogo",
                "Importa primero el catálogo de cuentas en la pestaña 'Catálogo de Cuentas'.",
            )
            return
        dlg = _DialogSelCuenta(self, self._cuentas)
        self.wait_window(dlg)
        if dlg.resultado:
            var.set(dlg.resultado)

    def _sel_destino_coi(self):
        from datetime import datetime
        hoy = datetime.now().strftime("%Y%m%d")
        ruta = filedialog.asksaveasfilename(
            title="Guardar archivo COI como...",
            defaultextension=".xlsx",
            initialfile=f"Polizas_COI_{self._rfc}_{hoy}.xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if ruta:
            self._v_destino_coi.set(ruta)

    def _generar(self):
        tipo = self._v_tipo_pol.get()
        try:
            num_inicio = int(self._v_num_inicio.get().strip())
        except ValueError:
            messagebox.showwarning("Número inválido",
                                   "El número inicial de póliza debe ser un entero.")
            return

        if tipo in ("bancario", "ambas"):
            cta_banco = self._v_cta_banco.get().strip()
            if not cta_banco:
                messagebox.showwarning("Cuenta banco requerida",
                                       "Especifica la cuenta contable del banco para pólizas bancarias.")
                return
            if not self._resultados:
                messagebox.showwarning("Sin conciliación",
                                       "No hay resultados de conciliación disponibles.\n"
                                       "Concilia primero desde la ventana principal.")
                return

        if tipo in ("diario", "ambas"):
            if not self._recibidos and not self._emitidos:
                messagebox.showwarning("Sin CFDIs",
                                       "No hay CFDIs cargados para generar pólizas de Diario.")
                return

        self._lbl_estado.config(text="Generando...", fg=C_BLUE)
        self.update_idletasks()

        threading.Thread(target=self._ejecutar_generacion,
                         args=(tipo, num_inicio),
                         daemon=True).start()

    def _ejecutar_generacion(self, tipo: str, num_inicio: int):
        try:
            from conciliacion import MovConciliado

            polizas_banc: list = []
            polizas_diar: list = []

            # ── Pólizas Bancarias (Ig/Eg) ─────────────────────────────────────
            if tipo in ("bancario", "ambas"):
                todos_conc: list[MovConciliado] = []
                for res in self._resultados:
                    todos_conc.extend(res.conciliados)

                polizas_banc = generar_polizas(
                    conciliados         = todos_conc,
                    rfc_empresa         = self._rfc,
                    num_inicio          = num_inicio,
                    cuenta_banco        = self._v_cta_banco.get().strip(),
                    terceros            = self._terceros,
                    cuenta_default_prov = self._v_cta_prov.get().strip(),
                    cuenta_default_cli  = self._v_cta_cli.get().strip(),
                    cuenta_nomina       = self._v_cta_nom.get().strip(),
                    incluir_sin_cfdi    = self._v_sin_cfdi.get(),
                )

            # ── Pólizas Diario (Dr) ───────────────────────────────────────────
            if tipo in ("diario", "ambas"):
                num_diario = num_inicio + len(polizas_banc)
                todos_cfdis = list(self._recibidos) + list(self._emitidos)
                self._reglas = cargar_reglas(self._rfc)

                polizas_diar = generar_polizas_diario(
                    cfdis                = todos_cfdis,
                    rfc_empresa          = self._rfc,
                    num_inicio           = num_diario,
                    terceros             = self._terceros,
                    reglas_prod          = self._reglas,
                    cta_cobrar           = self._v_cta_cobrar.get().strip(),
                    cta_pagar            = self._v_cta_pagar.get().strip(),
                    cta_iva_tras         = self._v_cta_iva_tras.get().strip(),
                    cta_iva_acre         = self._v_cta_iva_acre.get().strip(),
                    cta_iva_ret          = self._v_cta_iva_ret.get().strip(),
                    cta_isr_ret          = self._v_cta_isr_ret.get().strip(),
                    cta_default_ingreso  = self._v_cta_def_ingreso.get().strip(),
                    cta_default_gasto    = self._v_cta_def_gasto.get().strip(),
                )

            polizas = polizas_banc + polizas_diar
            n_pol   = len(polizas)
            n_ok    = sum(1 for p in polizas if p.cuadrada)
            n_desc  = n_pol - n_ok

            # 1. Agregar hoja al Excel de conciliación existente
            if self._output_path:
                import openpyxl
                from exportar_conciliacion import _hoja_polizas_coi
                # El reporte pudo haberse guardado como .xlsm (papel de trabajo con
                # macro): resolver la ruta real y reabrir con keep_vba para NO borrar
                # el vbaProject.bin firmado al regrabar.
                ruta = Path(self._output_path)
                if not ruta.exists() and ruta.with_suffix(".xlsm").exists():
                    ruta = ruta.with_suffix(".xlsm")
                if ruta.exists():
                    es_macro = ruta.suffix.lower() == ".xlsm"
                    wb = openpyxl.load_workbook(ruta, keep_vba=es_macro)
                    # Eliminar hoja anterior si existe
                    if "Pólizas COI" in wb.sheetnames:
                        del wb["Pólizas COI"]
                    _hoja_polizas_coi(wb, polizas)
                    wb.save(ruta)

            # 2. Exportar archivo COI independiente (si se indicó destino)
            destino_coi = self._v_destino_coi.get().strip()
            if destino_coi:
                self._exportar_archivo_coi(polizas, destino_coi)

            msg = (
                f"Pólizas generadas: {n_pol}\n"
                f"  ✔ Cuadradas:      {n_ok}\n"
                + (f"  ⚠ Descuadradas:  {n_desc}\n" if n_desc else "")
                + ("\nHoja 'Pólizas COI' agregada al reporte de conciliación."
                   if self._output_path else "")
                + (f"\nArchivo COI exportado:\n{destino_coi}" if destino_coi else "")
            )

            self.after(0, lambda: self._lbl_estado.config(
                text=f"✔ {n_pol} pólizas generadas", fg="#1B5E20"))
            self.after(0, lambda: messagebox.showinfo("Pólizas generadas", msg))

            if self._output_path and Path(self._output_path).exists():
                self.after(0, lambda: os.startfile(self._output_path))

        except Exception as e:
            import traceback
            tb = traceback.format_exc()
            self.after(0, lambda: self._lbl_estado.config(
                text="Error al generar", fg=C_RED))
            self.after(0, lambda: messagebox.showerror("Error", f"{e}\n\n{tb}"))

    def _exportar_archivo_coi(self, polizas: list, ruta: str):
        """Genera un Excel independiente en formato de importación Aspel COI."""
        import openpyxl
        from openpyxl.styles import Font as OFont, PatternFill as OFill, Alignment as OAlign
        from polizas_coi import polizas_a_filas

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Pólizas"

        # Filas 1-2: encabezado de referencia COI
        ws.cell(row=1, column=1).value = "Tipo Pol"
        ws.cell(row=2, column=2).value = "No. Cuenta"
        # (solo para orientar al usuario — COI las ignora)

        filas = polizas_a_filas(polizas)
        for i, fila in enumerate(filas, start=3):
            for j, val in enumerate(fila, start=1):
                ws.cell(row=i, column=j, value=val)

        # Anchos mínimos
        for col, w in enumerate([8, 16, 55, 30, 10, 16, 16, 20, 40, 8], 1):
            from openpyxl.utils import get_column_letter
            ws.column_dimensions[get_column_letter(col)].width = w

        wb.save(ruta)


# ── Diálogos auxiliares ────────────────────────────────────────────────────────

class _DialogReglaProd(tk.Toplevel):
    """Diálogo para crear o editar una ReglaProdServ (ClaveProdServ → cuentas)."""

    _TIPO_OPTS = [("Emitido — Ingreso/Venta", "E"),
                  ("Recibido — Gasto/Compra",  "R"),
                  ("Ambos",                    "")]

    def __init__(self, parent, regla: Optional[ReglaProdServ],
                 cuentas: list[CuentaContable]):
        super().__init__(parent)
        self.title("Regla de Producto/Servicio")
        fit_geometry(self, 600, 360)
        self.configure(bg=C_BG)
        self.resizable(False, False)
        self.grab_set()

        self._cuentas  = cuentas
        self.resultado: Optional[ReglaProdServ] = None

        self._v_prefijo = tk.StringVar(value=regla.prefijo       if regla else "")
        self._v_desc    = tk.StringVar(value=regla.descripcion    if regla else "")
        self._v_tipo    = tk.StringVar(value=(regla.tipo          if regla else "E"))
        self._v_ingreso = tk.StringVar(value=regla.cuenta_ingreso if regla else "")
        self._v_gasto   = tk.StringVar(value=regla.cuenta_gasto   if regla else "")

        # Stripe de acento
        tk.Frame(self, bg=T.ACCENT, height=3).pack(fill="x")

        p = tk.Frame(self, bg=C_BG, padx=20, pady=14)
        p.pack(fill="both", expand=True)
        p.columnconfigure(1, weight=1)

        # ── Fila 0: Prefijo ───────────────────────────────────────────────────
        tk.Label(p, text="Prefijo SAT:", font=FONT_B, bg=C_BG, fg=C_TEXT, anchor="e"
                 ).grid(row=0, column=0, sticky="e", padx=(0, 8), pady=6)
        ef0 = tk.Frame(p, bg=C_BG)
        ef0.grid(row=0, column=1, sticky="w", pady=6)
        tk.Entry(ef0, textvariable=self._v_prefijo, font=FONT_N, width=14,
                 bg=T.SURFACE, fg=T.TEXT, insertbackground=T.TEXT,
                 highlightthickness=1, highlightbackground=T.BORDER, relief="flat"
                 ).pack(side="left")
        tk.Label(p, text="Ej: 84 · 841 · 84111500", font=FONT_S, bg=C_BG, fg=C_GRAY2
                 ).grid(row=0, column=2, sticky="w", padx=(8, 0))

        # ── Fila 1: Descripción ───────────────────────────────────────────────
        tk.Label(p, text="Descripción:", font=FONT_B, bg=C_BG, fg=C_TEXT, anchor="e"
                 ).grid(row=1, column=0, sticky="e", padx=(0, 8), pady=6)
        tk.Entry(p, textvariable=self._v_desc, font=FONT_N, width=38,
                 bg=T.SURFACE, fg=T.TEXT, insertbackground=T.TEXT,
                 highlightthickness=1, highlightbackground=T.BORDER, relief="flat"
                 ).grid(row=1, column=1, columnspan=2, sticky="ew", pady=6)

        # ── Fila 2: Tipo ──────────────────────────────────────────────────────
        tk.Label(p, text="Tipo:", font=FONT_B, bg=C_BG, fg=C_TEXT, anchor="e"
                 ).grid(row=2, column=0, sticky="e", padx=(0, 8), pady=6)
        tf = tk.Frame(p, bg=C_BG)
        tf.grid(row=2, column=1, columnspan=2, sticky="w", pady=6)
        colores = {"E": T.ACCENT, "R": T.DANGER, "": T.TEXT}
        for lbl_txt, val in self._TIPO_OPTS:
            tk.Radiobutton(
                tf, text=lbl_txt, variable=self._v_tipo, value=val,
                font=FONT_N, bg=C_BG, fg=colores[val],
                activebackground=C_BG, selectcolor=T.SURFACE,
                command=self._on_tipo_change,
            ).pack(side="left", padx=(0, 14))

        # ── Fila 3: Cuenta Ingresos (visible para E / Ambos) ─────────────────
        self._lbl_ing = tk.Label(p, text="Cta. Ingresos  (Emitidas):", font=FONT_B,
                                 bg=C_BG, fg=T.ACCENT, anchor="e")
        self._lbl_ing.grid(row=3, column=0, sticky="e", padx=(0, 8), pady=6)
        self._ef_ing = tk.Frame(p, bg=C_BG)
        self._ef_ing.grid(row=3, column=1, sticky="ew", pady=6)
        self._ef_ing.columnconfigure(0, weight=1)
        tk.Entry(self._ef_ing, textvariable=self._v_ingreso, font=FONT_N,
                 bg=T.SURFACE, fg=T.TEXT, insertbackground=T.TEXT,
                 highlightthickness=1, highlightbackground=T.BORDER, relief="flat"
                 ).grid(row=0, column=0, sticky="ew")
        _btn(self._ef_ing, "...", lambda: self._sel(self._v_ingreso),
             bg=T.SURFACE, fg=T.TEXT, padx=6,
             highlightthickness=1, highlightbackground=T.BORDER
             ).grid(row=0, column=1, padx=(4, 0))
        self._hint_ing = tk.Label(p, text="Cr — reconocimiento de ingreso",
                                  font=FONT_S, bg=C_BG, fg=C_GRAY2)
        self._hint_ing.grid(row=3, column=2, sticky="w", padx=(8, 0))

        # ── Fila 4: Cuenta Gastos (visible para R / Ambos) ───────────────────
        self._lbl_gas = tk.Label(p, text="Cta. Gastos  (Recibidas):", font=FONT_B,
                                 bg=C_BG, fg=T.DANGER, anchor="e")
        self._lbl_gas.grid(row=4, column=0, sticky="e", padx=(0, 8), pady=6)
        self._ef_gas = tk.Frame(p, bg=C_BG)
        self._ef_gas.grid(row=4, column=1, sticky="ew", pady=6)
        self._ef_gas.columnconfigure(0, weight=1)
        tk.Entry(self._ef_gas, textvariable=self._v_gasto, font=FONT_N,
                 bg=T.SURFACE, fg=T.TEXT, insertbackground=T.TEXT,
                 highlightthickness=1, highlightbackground=T.BORDER, relief="flat"
                 ).grid(row=0, column=0, sticky="ew")
        _btn(self._ef_gas, "...", lambda: self._sel(self._v_gasto),
             bg=T.SURFACE, fg=T.TEXT, padx=6,
             highlightthickness=1, highlightbackground=T.BORDER
             ).grid(row=0, column=1, padx=(4, 0))
        self._hint_gas = tk.Label(p, text="Dr — reconocimiento de gasto",
                                  font=FONT_S, bg=C_BG, fg=C_GRAY2)
        self._hint_gas.grid(row=4, column=2, sticky="w", padx=(8, 0))

        # ── Separador y botones ───────────────────────────────────────────────
        tk.Frame(p, bg=T.BORDER, height=1).grid(
            row=5, column=0, columnspan=3, sticky="ew", pady=(8, 6))
        bf = tk.Frame(p, bg=C_BG)
        bf.grid(row=6, column=0, columnspan=3, sticky="e")
        _btn(bf, "✔  Guardar", self._guardar, bg=T.SUCCESS, fg=T.ACCENT_FG
             ).pack(side="left", padx=(0, 8))
        _btn(bf, "Cancelar", self.destroy,
             bg=T.SURFACE, fg=T.TEXT,
             highlightthickness=1, highlightbackground=T.BORDER
             ).pack(side="left")

        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self._on_tipo_change()  # estado inicial

    def _on_tipo_change(self):
        """Muestra/oculta las filas de cuenta según el tipo seleccionado."""
        tipo = self._v_tipo.get()
        show_e = tipo in ("E", "")
        show_r = tipo in ("R", "")
        for w in (self._lbl_ing, self._ef_ing, self._hint_ing):
            if show_e:
                w.grid()
            else:
                w.grid_remove()
        for w in (self._lbl_gas, self._ef_gas, self._hint_gas):
            if show_r:
                w.grid()
            else:
                w.grid_remove()

    def _sel(self, var: tk.StringVar):
        dlg = _DialogSelCuenta(self, self._cuentas)
        self.wait_window(dlg)
        if dlg.resultado:
            var.set(dlg.resultado)

    def _guardar(self):
        prefijo = self._v_prefijo.get().strip()
        if not prefijo or not prefijo.isdigit():
            messagebox.showwarning("Prefijo inválido",
                                   "El prefijo debe ser numérico (2-8 dígitos).",
                                   parent=self)
            return
        tipo = self._v_tipo.get()
        self.resultado = ReglaProdServ(
            prefijo        = prefijo,
            descripcion    = self._v_desc.get().strip(),
            cuenta_ingreso = self._v_ingreso.get().strip() if tipo in ("E", "") else "",
            cuenta_gasto   = self._v_gasto.get().strip()   if tipo in ("R", "") else "",
            tipo           = tipo,
        )
        self.destroy()


class _DialogCuenta(tk.Toplevel):
    """Diálogo para asignar cuenta contable a un tercero."""

    def __init__(self, parent, rfc: str, nombre: str,
                 cuenta_actual: str, cuentas: list[CuentaContable]):
        super().__init__(parent)
        self.title(f"Asignar cuenta — {nombre[:40]}")
        fit_geometry(self, 540, 400)
        self.configure(bg=C_BG)
        self.resizable(False, True)
        self.resultado: Optional[str] = None
        self._cuentas = cuentas

        _lbl(self, f"RFC: {rfc}", font=FONT_B, bg=C_BG
             ).pack(anchor="w", padx=12, pady=(10, 0))
        _lbl(self, nombre, bg=C_BG, fg=C_GRAY2
             ).pack(anchor="w", padx=12, pady=(0, 6))

        tk.Label(self, text="Cuenta actual:", font=FONT_S, bg=C_BG, fg=C_GRAY2
                 ).pack(anchor="w", padx=12)
        self._v_nueva = tk.StringVar(value=cuenta_actual
                                     if cuenta_actual != "(sin cuenta)" else "")

        ef = tk.Frame(self, bg=C_BG)
        ef.pack(fill="x", padx=12, pady=4)
        ef.columnconfigure(0, weight=1)
        _entry_var(ef, self._v_nueva, width=30
                   ).grid(row=0, column=0, sticky="ew")
        _btn(ef, "...", self._sel, bg=C_GRAY, fg=C_TEXT, padx=6
             ).grid(row=0, column=1, padx=(4, 0))

        # Lista de cuentas para selección rápida
        self._bus_var = tk.StringVar()
        self._bus_var.trace_add("write", lambda *_: self._filtrar())
        bus_f = tk.Frame(self, bg=C_BG)
        bus_f.pack(fill="x", padx=12, pady=(4, 0))
        tk.Label(bus_f, text="Buscar cuenta:", font=FONT_S, bg=C_BG
                 ).pack(side="left")
        tk.Entry(bus_f, textvariable=self._bus_var, width=25,
                 font=FONT_S, bg=C_WHITE, relief="solid", bd=1
                 ).pack(side="left", padx=6)

        frm = tk.Frame(self, bg=C_BG)
        frm.pack(fill="both", expand=True, padx=12, pady=4)
        frm.columnconfigure(0, weight=1)
        frm.rowconfigure(0, weight=1)
        self._lb = tk.Listbox(frm, font=("Consolas", 8),
                               bg=C_WHITE, selectbackground=C_ACCENT,
                               relief="flat", bd=0)
        sb = ttk.Scrollbar(frm, orient="vertical", command=self._lb.yview)
        self._lb.configure(yscrollcommand=sb.set)
        self._lb.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        self._lb.bind("<<ListboxSelect>>", self._on_sel)
        self._lb.bind("<Double-1>", lambda _: self._aceptar())
        self._filtrar()

        bf = tk.Frame(self, bg=C_BG)
        bf.pack(fill="x", padx=12, pady=8)
        _btn(bf, "✔  Aceptar", self._aceptar, bg=C_GREEN).pack(side="right")
        _btn(bf, "Cancelar",  self.destroy,   bg=C_GRAY, fg=C_TEXT
             ).pack(side="right", padx=(0, 6))

    def _filtrar(self):
        t = self._bus_var.get().strip()
        filtradas = buscar_cuentas(self._cuentas, t) if t else self._cuentas
        self._lb.delete(0, tk.END)
        for c in filtradas:
            self._lb.insert(tk.END, f"{c.cuenta:<18} {c.nombre}")
        self._filtradas = filtradas

    def _on_sel(self, _event=None):
        sel = self._lb.curselection()
        if sel and hasattr(self, "_filtradas"):
            idx = sel[0]
            if idx < len(self._filtradas):
                self._v_nueva.set(self._filtradas[idx].cuenta)

    def _sel(self):
        self._on_sel()

    def _aceptar(self):
        self.resultado = self._v_nueva.get().strip()
        self.destroy()


class _DialogSelCuenta(tk.Toplevel):
    """Diálogo compacto para seleccionar una cuenta del catálogo."""

    def __init__(self, parent, cuentas: list[CuentaContable]):
        super().__init__(parent)
        self.title("Seleccionar cuenta")
        fit_geometry(self, 520, 380)
        self.configure(bg=C_BG)
        self.resizable(False, True)
        self.resultado: Optional[str] = None
        self._cuentas   = cuentas
        self._filtradas = cuentas

        self._bus_var = tk.StringVar()
        self._bus_var.trace_add("write", lambda *_: self._filtrar())

        bus_f = tk.Frame(self, bg=C_BG)
        bus_f.pack(fill="x", padx=12, pady=10)
        tk.Label(bus_f, text="Buscar:", font=FONT_B, bg=C_BG, fg=C_TEXT
                 ).pack(side="left")
        tk.Entry(bus_f, textvariable=self._bus_var, width=35,
                 font=FONT_N, bg=T.SURFACE, fg=T.TEXT, insertbackground=T.TEXT,
                 highlightthickness=1, highlightbackground=T.BORDER, relief="flat"
                 ).pack(side="left", padx=8)

        frm = tk.Frame(self, bg=C_BG)
        frm.pack(fill="both", expand=True, padx=12)
        frm.columnconfigure(0, weight=1)
        frm.rowconfigure(0, weight=1)
        self._lb = tk.Listbox(frm, font=F.MONO,
                               bg=T.SURFACE, fg=T.TEXT,
                               selectbackground=T.SELECTION,
                               selectforeground=T.SELECTION_FG,
                               relief="flat", bd=0)
        sb = ttk.Scrollbar(frm, orient="vertical", command=self._lb.yview)
        self._lb.configure(yscrollcommand=sb.set)
        self._lb.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        self._lb.bind("<Double-1>", lambda _: self._aceptar())
        self._filtrar()

        bf = tk.Frame(self, bg=C_BG)
        bf.pack(fill="x", padx=12, pady=8)
        _btn(bf, "✔  Seleccionar", self._aceptar,
             bg=T.SUCCESS, fg=T.ACCENT_FG).pack(side="right")
        _btn(bf, "Cancelar", self.destroy,
             bg=T.SURFACE, fg=T.TEXT,
             highlightthickness=1, highlightbackground=T.BORDER
             ).pack(side="right", padx=(0, 6))

    def _filtrar(self):
        t = self._bus_var.get().strip()
        self._filtradas = buscar_cuentas(self._cuentas, t) if t else self._cuentas
        self._lb.delete(0, tk.END)
        for c in self._filtradas:
            self._lb.insert(tk.END, f"{c.cuenta:<18} {c.nombre}")

    def _aceptar(self):
        sel = self._lb.curselection()
        if sel:
            idx = sel[0]
            if idx < len(self._filtradas):
                self.resultado = self._filtradas[idx].cuenta
        self.destroy()
