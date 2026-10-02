"""
ventana_conciliacion.py
Pestaña / ventana flotante de conciliación banco ↔ CFDIs.
Soporta múltiples estados de cuenta en un solo reporte.

Uso:
    from ventana_conciliacion import VentanaConciliacion
    tab = VentanaConciliacion(notebook, movimientos, info_cuenta)
    notebook.add(tab, text="Conciliación CFDI")
"""

from __future__ import annotations

import os
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Optional

from cfdi_parser           import leer_cfdis_carpetas, clasificar_cfdis, CFDI
from conciliacion          import mov_a_banco, MovBanco, conciliar_multibanco
from exportar_conciliacion import exportar_conciliacion_multi
from empresas              import registrar_desde_cif, obtener_empresa
from grupos_conciliacion   import cargar_grupos
from repo_conciliaciones   import (guardar_conciliacion, listar_conciliaciones,
                                    cargar_conciliacion, eliminar_conciliacion,
                                    obtener_ultima_conciliacion)

try:
    from ventana_grupos_rfc import VentanaGruposRFC
    _GRUPOS_OK = True
except ImportError:
    _GRUPOS_OK = False

try:
    from ventana_polizas import VentanaPolizas
    _COI_DISPONIBLE = True
except ImportError:
    _COI_DISPONIBLE = False

# ── Tema centralizado ─────────────────────────────────────────────────────────
from ui_theme import T, F

# Alias locales — se actualizan en _refresh_theme() al crear la ventana.
# De esta forma todo el código que usa C_BG, C_TEXT, etc. sigue funcionando
# y recibe los colores del tema activo (light / dark).
C_BG    = T.BG
C_WHITE = T.SURFACE
C_ACCENT= T.ACCENT
C_GREEN = T.SUCCESS
C_RED   = T.DANGER
C_TEXT  = T.TEXT
C_GRAY  = T.BORDER_STRONG
C_GRAY2 = T.TEXT_MUTED
FONT_N  = F.BODY
FONT_B  = F.BTN_SM
FONT_T  = F.TITLE
FONT_S  = F.SMALL


def _refresh_theme():
    """Sincroniza los alias de módulo con el tema activo. Llamar al inicio de _build_ui."""
    global C_BG, C_WHITE, C_ACCENT, C_GREEN, C_RED, C_TEXT, C_GRAY, C_GRAY2
    global FONT_N, FONT_B, FONT_T, FONT_S
    C_BG    = T.BG
    C_WHITE = T.SURFACE
    C_ACCENT= T.ACCENT
    C_GREEN = T.SUCCESS
    C_RED   = T.DANGER
    C_TEXT  = T.TEXT
    C_GRAY  = T.BORDER_STRONG
    C_GRAY2 = T.TEXT_MUTED
    FONT_N  = F.BODY
    FONT_B  = F.BTN_SM
    FONT_T  = F.TITLE
    FONT_S  = F.SMALL

# Colores por defecto para chips de banco (hex sin #) — saturados, no el accent genérico
_COLORES_BANCO = [
    "1565C0", "2E7D32", "6A1B9A", "BF360C",
    "00695C", "4527A0", "558B2F", "AD1457",
]


def _card(parent, **kw) -> tk.Frame:
    return tk.Frame(parent, bg=T.SURFACE,
                    highlightthickness=1, highlightbackground=T.BORDER,
                    **kw)


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
                     bd=0, padx=10, pady=5, **kw)


# ─────────────────────────────────────────────────────────────────────────────

class VentanaConciliacion(tk.Frame):
    """
    Pestaña de conciliación banco ↔ CFDIs con soporte multi-banco.

    Parámetros
    ----------
    parent       : widget padre (Notebook o Frame)
    movimientos  : lista de MovBanco o bancos.base.Movimiento (primer banco)
    info_cuenta  : bancos.base.InfoCuenta (primer banco)
    """

    def __init__(self, parent, movimientos=None, info_cuenta=None, **kw):
        super().__init__(parent, bg=C_BG, **kw)
        self.info_cuenta = info_cuenta

        # Lista de bancos: [(nombre, color_hex, [MovBanco, ...])]
        self._bancos: list[tuple[str, str, list[MovBanco]]] = []

        if movimientos:
            nombre = (getattr(info_cuenta, "banco", None) or "Banco 1") if info_cuenta else "Banco 1"
            color  = (getattr(info_cuenta, "color_hex", None) or _COLORES_BANCO[0]) if info_cuenta else _COLORES_BANCO[0]
            movs   = self._normalizar_movs(movimientos)
            self._bancos.append((nombre, color, movs))

        self._xml_items: list[str] = []   # puede contener carpetas o archivos .xml
        self._build_ui()
        self._refrescar_lista_bancos()
        self.after(100, self._refrescar_status_catalogos)

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _normalizar_movs(self, movs) -> list[MovBanco]:
        if not movs:
            return []
        if isinstance(movs[0], MovBanco):
            return movs
        return [mov_a_banco(m) for m in movs]

    # ── UI ────────────────────────────────────────────────────────────────────

    def _build_ui(self):
        # Sincronizar alias de color con el tema activo (light/dark)
        _refresh_theme()
        self.configure(bg=C_BG)
        self.columnconfigure(0, weight=1)

        # Título
        tk.Label(self, text="Conciliación Bancaria — CFDI",
                 font=FONT_T, bg=C_BG, fg=C_TEXT
                 ).grid(row=0, column=0, sticky="ew", pady=(12, 2), padx=16)
        tk.Label(self,
                 text="Cruza uno o varios estados de cuenta con CFDIs XML del SAT",
                 font=FONT_S, bg=C_BG, fg=C_GRAY2
                 ).grid(row=1, column=0, sticky="ew", padx=16)

        # ── Tarjeta 1: RFC + CSF ──────────────────────────────────────────────
        c1 = _card(self)
        c1.grid(row=2, column=0, sticky="ew", padx=16, pady=(10, 4))
        c1.columnconfigure(1, weight=1)

        _lbl(c1, "RFC de tu empresa:", font=FONT_B
             ).grid(row=0, column=0, sticky="w", padx=12, pady=8)
        self._rfc_var = tk.StringVar(
            value=getattr(self.info_cuenta, "rfc", "") if self.info_cuenta else ""
        )
        tk.Entry(c1, textvariable=self._rfc_var, width=20,
                 font=FONT_N, bg=C_WHITE, fg=C_TEXT, relief="solid", bd=1
                 ).grid(row=0, column=1, sticky="ew", padx=(0, 8), pady=8)
        _btn(c1, "Subir CSF / CIF", self._subir_cif, bg="#5C6BC0"
             ).grid(row=0, column=2, padx=(0, 12), pady=8)

        # Fila 1: nombre / razón social de la empresa (verde cuando CSF cargada)
        self._lbl_empresa = _lbl(c1, "", font=FONT_B, fg="#1B5E20")
        self._lbl_empresa.grid(row=1, column=0, columnspan=3, sticky="w",
                               padx=14, pady=(0, 1))

        # Fila 2: régimen(es) SAT
        self._lbl_regimen = _lbl(c1, "", font=FONT_S, fg=C_GRAY2)
        self._lbl_regimen.grid(row=2, column=0, columnspan=3, sticky="w",
                               padx=14, pady=(0, 8))

        self._rfc_var.trace_add("write", lambda *_: self._actualizar_lbl_regimen())
        self._actualizar_lbl_regimen()

        # ── Tarjeta 2: Archivos/Carpetas XML ─────────────────────────────────
        c2 = _card(self)
        c2.grid(row=3, column=0, sticky="ew", padx=16, pady=4)
        c2.columnconfigure(0, weight=1)

        hdr2 = tk.Frame(c2, bg=C_WHITE)
        hdr2.grid(row=0, column=0, sticky="ew", padx=12, pady=(8, 4))
        hdr2.columnconfigure(0, weight=1)
        _lbl(hdr2, "Archivos XML / Carpetas de CFDIs:", font=FONT_B
             ).grid(row=0, column=0, sticky="w")
        _btn(hdr2, "+ XMLs", self._agregar_xmls, bg=C_ACCENT
             ).grid(row=0, column=1, padx=(8, 0))
        _btn(hdr2, "+ Carpeta", self._agregar_carpeta, bg=C_ACCENT
             ).grid(row=0, column=2, padx=(4, 0))
        _btn(hdr2, "Quitar", self._quitar_xml, bg=C_GRAY, fg=C_TEXT
             ).grid(row=0, column=3, padx=(4, 0))
        _btn(hdr2, "Limpiar todo", self._limpiar_carpetas, bg=C_GRAY, fg=C_TEXT
             ).grid(row=0, column=4, padx=(4, 0))

        frm_lst2 = tk.Frame(c2, bg=C_WHITE)
        frm_lst2.grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 4))
        frm_lst2.columnconfigure(0, weight=1)
        self._lst_carpetas = tk.Listbox(
            frm_lst2, height=4, font=FONT_S,
            bg=T.SURFACE, fg=T.TEXT,
            selectbackground=T.SELECTION, selectforeground=T.SELECTION_FG,
            relief="flat", bd=0,
        )
        self._lst_carpetas.grid(row=0, column=0, sticky="ew")
        _sb_xml = ttk.Scrollbar(frm_lst2, orient="vertical",
                                command=self._lst_carpetas.yview)
        _sb_xml.grid(row=0, column=1, sticky="ns")
        self._lst_carpetas.config(yscrollcommand=_sb_xml.set)

        self._lbl_carpetas_count = _lbl(c2, "0 elementos seleccionados", font=FONT_S, fg=C_GRAY2)
        self._lbl_carpetas_count.grid(row=2, column=0, sticky="e", padx=12, pady=(0, 8))

        # ── Tarjeta 3: Bancos ─────────────────────────────────────────────────
        c3 = _card(self)
        c3.grid(row=4, column=0, sticky="ew", padx=16, pady=4)
        c3.columnconfigure(0, weight=1)

        hdr3 = tk.Frame(c3, bg=C_WHITE)
        hdr3.grid(row=0, column=0, sticky="ew", padx=12, pady=(8, 4))
        hdr3.columnconfigure(0, weight=1)
        _lbl(hdr3, "Estados de cuenta bancarios:", font=FONT_B
             ).grid(row=0, column=0, sticky="w")
        _btn(hdr3, "+ Agregar banco (Excel)", self._agregar_banco_excel, bg=C_ACCENT
             ).grid(row=0, column=1, padx=(8, 0))
        _btn(hdr3, "Quitar seleccionado", self._quitar_banco, bg=C_GRAY, fg=C_TEXT
             ).grid(row=0, column=2, padx=(4, 0))

        self._lst_bancos = tk.Listbox(
            c3, height=4, font=FONT_S,
            bg=T.SURFACE, fg=T.TEXT,
            selectbackground=T.SELECTION, selectforeground=T.SELECTION_FG,
            relief="flat", bd=0,
        )
        self._lst_bancos.grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 4))
        self._lbl_bancos_count = _lbl(c3, "0 bancos cargados", font=FONT_S, fg=C_GRAY2)
        self._lbl_bancos_count.grid(row=2, column=0, sticky="e", padx=12, pady=(0, 8))

        # ── Tarjeta 4: Destino ────────────────────────────────────────────────
        c4 = _card(self)
        c4.grid(row=5, column=0, sticky="ew", padx=16, pady=4)
        c4.columnconfigure(1, weight=1)

        _lbl(c4, "Guardar resultado en:", font=FONT_B
             ).grid(row=0, column=0, sticky="w", padx=12, pady=8)
        self._destino_var = tk.StringVar()
        tk.Entry(c4, textvariable=self._destino_var, font=FONT_S,
                 bg=C_WHITE, fg=C_TEXT, relief="solid", bd=1
                 ).grid(row=0, column=1, sticky="ew", padx=(0, 4), pady=8)
        _btn(c4, "...", self._elegir_destino, bg=C_GRAY, fg=C_TEXT
             ).grid(row=0, column=2, padx=(0, 12), pady=8)

        # ── Tarjeta 5: Catálogos COI ──────────────────────────────────────────
        c5 = _card(self)
        c5.grid(row=6, column=0, sticky="ew", padx=16, pady=4)
        c5.columnconfigure(0, weight=1)

        # Encabezado de la tarjeta
        hdr5 = tk.Frame(c5, bg=C_WHITE)
        hdr5.grid(row=0, column=0, sticky="ew", padx=12, pady=(8, 4))
        hdr5.columnconfigure(0, weight=1)
        _lbl(hdr5, "Catalogos COI (para generacion de polizas):", font=FONT_B
             ).grid(row=0, column=0, sticky="w")

        # ── Fila: Catálogo de Cuentas ─────────────────────────────────────────
        row_cta = tk.Frame(c5, bg=C_WHITE)
        row_cta.grid(row=1, column=0, sticky="ew", padx=12, pady=(2, 2))
        row_cta.columnconfigure(1, weight=1)

        _lbl(row_cta, "Plan de cuentas:", font=FONT_B, fg="#37474F"
             ).grid(row=0, column=0, sticky="w", padx=(0, 8))
        self._lbl_cuentas_status = _lbl(
            row_cta, "No cargado",
            font=FONT_S, fg=C_RED,
        )
        self._lbl_cuentas_status.grid(row=0, column=1, sticky="w")
        _btn(row_cta, "Importar Excel COI",
             self._importar_cuentas_catalogo, bg=T.COI, fg=T.ACCENT_FG
             ).grid(row=0, column=2, padx=(8, 0))

        # ── Fila: Clientes / Proveedores ──────────────────────────────────────
        row_terc = tk.Frame(c5, bg=C_WHITE)
        row_terc.grid(row=2, column=0, sticky="ew", padx=12, pady=(2, 8))
        row_terc.columnconfigure(1, weight=1)

        _lbl(row_terc, "Clientes / Proveedores:", font=FONT_B, fg=T.COI
             ).grid(row=0, column=0, sticky="w", padx=(0, 8))
        self._lbl_terceros_status = _lbl(
            row_terc, "No cargado",
            font=FONT_S, fg=C_RED,
        )
        self._lbl_terceros_status.grid(row=0, column=1, sticky="w")
        _btn(row_terc, "Auto-poblar desde XMLs",
             self._poblar_terceros_catalogo, bg=T.COI, fg=T.ACCENT_FG
             ).grid(row=0, column=2, padx=(8, 0))
        _btn(row_terc, "Gestionar",
             self._gestionar_catalogos, bg=T.SURFACE, fg=T.TEXT,
             highlightthickness=1, highlightbackground=T.BORDER,
             ).grid(row=0, column=3, padx=(4, 0))

        # ── Fila: Grupos de RFC ───────────────────────────────────────────────
        row_grp = tk.Frame(c5, bg=C_WHITE)
        row_grp.grid(row=3, column=0, sticky="ew", padx=12, pady=(0, 8))
        row_grp.columnconfigure(1, weight=1)

        _lbl(row_grp, "Grupos de RFC:", font=FONT_B, fg=T.COI
             ).grid(row=0, column=0, sticky="w", padx=(0, 8))
        self._lbl_grupos = _lbl(
            row_grp, "0 grupos configurados",
            font=FONT_S, fg=C_GRAY2,
        )
        self._lbl_grupos.grid(row=0, column=1, sticky="w")
        _btn(row_grp, "Gestionar grupos",
             self._abrir_grupos_rfc, bg="#4527A0", fg=T.ACCENT_FG
             ).grid(row=0, column=2, padx=(8, 0))
        self._rfc_var.trace_add("write", lambda *_: self._refrescar_lbl_grupos())

        # ── Barra de acciones principal (misma estructura que el tab Convertidor) ──
        action_bar = tk.Frame(self, bg=T.BG, pady=8)
        action_bar.grid(row=7, column=0, sticky="ew", padx=12)

        self._btn_conciliar = tk.Button(
            action_bar, text="▶  CONCILIAR AHORA",
            command=self._conciliar,
            bg=T.SUCCESS, fg=T.ACCENT_FG, font=F.BTN,
            relief="flat", padx=22, pady=7, cursor="hand2", bd=0,
            activebackground=T.BLUE, activeforeground=T.ACCENT_FG,
        )
        self._btn_conciliar.pack(side="right")

        self._btn_polizas = tk.Button(
            action_bar,
            text="⚙  Pólizas COI  (conciliar primero)",
            command=self._abrir_polizas,
            bg=T.COI, fg="#90A4AE", font=F.BTN_SM,
            relief="flat", padx=14, pady=7, cursor="hand2", bd=0,
            state="disabled",
            activebackground=T.COI_SOFT, activeforeground=T.ACCENT_FG,
        )
        self._btn_polizas.pack(side="right", padx=(0, 8))

        tk.Button(
            action_bar,
            text="Buscar CFDIs que sumen un importe",
            command=self._abrir_buscador_suma,
            bg=T.SURFACE, fg=T.TEXT_MUTED, font=F.BODY,
            relief="flat", padx=12, pady=7, cursor="hand2", bd=0,
            highlightthickness=1, highlightbackground=T.BORDER,
            activebackground=T.ACCENT_SOFT, activeforeground=T.ACCENT,
        ).pack(side="left")

        tk.Button(
            action_bar,
            text="📂  Abrir conciliación guardada",
            command=self._abrir_historial,
            bg=T.SURFACE, fg=T.TEXT_MUTED, font=F.BODY,
            relief="flat", padx=12, pady=7, cursor="hand2", bd=0,
            highlightthickness=1, highlightbackground=T.BORDER,
            activebackground=T.ACCENT_SOFT, activeforeground=T.ACCENT,
        ).pack(side="left", padx=(8, 0))

        self._var_fusionar = tk.BooleanVar(value=True)
        tk.Checkbutton(
            action_bar,
            text="🔗 Fusionar con la última conciliación guardada",
            variable=self._var_fusionar,
            bg=T.BG, fg=T.TEXT_MUTED, font=F.SMALL,
            selectcolor=T.SURFACE, activebackground=T.BG,
            activeforeground=T.TEXT, cursor="hand2",
        ).pack(side="left", padx=(12, 0))

        self._lbl_coi_estado = tk.Label(
            action_bar, text="",
            font=F.SMALL, bg=T.BG, fg=T.TEXT_MUTED, anchor="e",
        )
        self._lbl_coi_estado.pack(side="right", padx=(0, 8))

        # Estado de última conciliación (rellenado tras conciliar con éxito)
        self._ultimo_resultado: dict | None = None

        # ── Separador ─────────────────────────────────────────────────────────
        tk.Frame(self, bg=T.BORDER, height=1).grid(row=8, column=0, sticky="ew", padx=12)

        # ── Barra de progreso ─────────────────────────────────────────────────
        self._prog = ttk.Progressbar(self, mode="indeterminate")
        self._prog.grid(row=9, column=0, sticky="ew", padx=16, pady=(6, 0))

        # ── Log estilo terminal (consistente con el tab Convertidor) ──────────
        log_wrap = tk.Frame(self, bg=T.BG)
        log_wrap.grid(row=10, column=0, sticky="nsew", padx=16, pady=(4, 12))
        log_wrap.columnconfigure(0, weight=1)
        log_wrap.rowconfigure(0, weight=1)

        self._log = tk.Text(
            log_wrap, height=8, font=F.MONO,
            bg="#0D1117", fg="#94A3B8",
            relief="flat", state="disabled", wrap="word",
            highlightthickness=1, highlightbackground=T.BORDER,
        )
        self._log.grid(row=0, column=0, sticky="nsew")
        _sb_log = ttk.Scrollbar(log_wrap, orient="vertical", command=self._log.yview)
        _sb_log.grid(row=0, column=1, sticky="ns")
        self._log.configure(yscrollcommand=_sb_log.set)
        self.rowconfigure(10, weight=1)

        # Actualizar estado inicial de catálogos según RFC actual
        self._rfc_var.trace_add("write", lambda *_: self._refrescar_status_catalogos())

    # ── CSF / Régimen ─────────────────────────────────────────────────────────

    def _actualizar_lbl_regimen(self):
        rfc = self._rfc_var.get().strip().upper()

        if not rfc:
            self._lbl_empresa.config(text="")
            self._lbl_regimen.config(
                text="  Ingresa el RFC de tu empresa o sube la CSF",
                fg=C_GRAY2,
            )
            return

        empresa = obtener_empresa(rfc)

        if not empresa:
            self._lbl_empresa.config(text="")
            self._lbl_regimen.config(
                text="  Sin CSF cargada — sube tu Constancia de Situación Fiscal para validar régimen",
                fg="#BF360C",
            )
            return

        nombre  = empresa.get("nombre", rfc)
        codigos = empresa.get("regimenes", [])
        nombres = empresa.get("regimenes_nombres", [])

        self._lbl_empresa.config(text=f"  ✔  {nombre}", fg="#1B5E20")

        if codigos:
            if nombres:
                partes = [
                    f"{n} ({c})" if c else n
                    for n, c in zip(nombres, codigos)
                ]
            else:
                partes = codigos
            self._lbl_regimen.config(
                text=f"  Régimen(es) SAT: {',  '.join(partes)}",
                fg="#2E7D32",
            )
        else:
            self._lbl_regimen.config(
                text="  Empresa registrada sin régimen — sube la CSF para completar",
                fg="#E65100",
            )

    def _subir_cif(self):
        pdf = filedialog.askopenfilename(
            title="Seleccionar Constancia de Situación Fiscal (PDF)",
            filetypes=[("PDF", "*.pdf"), ("Todos", "*.*")],
        )
        if not pdf:
            return
        try:
            ok, msg, datos = registrar_desde_cif(pdf)
            if ok and datos:
                self._rfc_var.set(datos.rfc)
                self._actualizar_lbl_regimen()
                regs_txt = ",\n  ".join(
                    f"{n} ({c})" if c else n
                    for n, c in zip(datos.regimenes_nombres, datos.regimenes_codigos)
                ) or "(régimen no identificado)"
                messagebox.showinfo(
                    "CSF cargada correctamente",
                    f"RFC:    {datos.rfc}\n"
                    f"Nombre: {datos.razon_social or '(sin nombre)'}\n\n"
                    f"Régimen(es):\n  {regs_txt}",
                )
            else:
                messagebox.showerror("Error al leer CSF", msg)
        except Exception as e:
            messagebox.showerror("Error al procesar CSF", str(e))

    # ── Gestión de bancos ─────────────────────────────────────────────────────

    def _refrescar_lista_bancos(self):
        self._lst_bancos.delete(0, tk.END)
        for nombre, _, movs in self._bancos:
            self._lst_bancos.insert(tk.END, f"  {nombre}  ({len(movs)} movimientos)")
        total = sum(len(m) for _, _, m in self._bancos)
        n = len(self._bancos)
        self._lbl_bancos_count.config(
            text=f"{n} banco(s) — {total} movimientos en total"
        )

    def _agregar_banco_excel(self):
        archivo = filedialog.askopenfilename(
            title="Seleccionar Excel de banco",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")],
        )
        if not archivo:
            return
        try:
            movs = self._leer_excel_banco(archivo)
            nombre_default = Path(archivo).stem
            nombre = self._pedir_nombre_banco(nombre_default)
            if not nombre:
                return
            idx = len(self._bancos)
            color = _COLORES_BANCO[idx % len(_COLORES_BANCO)]
            self._bancos.append((nombre, color, movs))
            self._refrescar_lista_bancos()
            self._log_write(f"+ Banco '{nombre}': {len(movs)} movimientos cargados\n")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo leer el Excel:\n{e}")

    def _pedir_nombre_banco(self, default: str) -> str:
        dlg = tk.Toplevel(self)
        dlg.title("Nombre del banco")
        dlg.resizable(False, False)
        dlg.grab_set()
        tk.Label(dlg, text="Nombre del banco / periodo:", font=FONT_N,
                 padx=12, pady=8).pack()
        var = tk.StringVar(value=default)
        entry = tk.Entry(dlg, textvariable=var, width=30, font=FONT_N)
        entry.pack(padx=12)
        entry.select_range(0, tk.END)
        entry.focus()
        resultado = []

        def ok():
            resultado.append(var.get().strip() or default)
            dlg.destroy()

        def cancelar():
            dlg.destroy()

        btns = tk.Frame(dlg)
        btns.pack(pady=8)
        tk.Button(btns, text="Aceptar",  command=ok,      font=FONT_N, width=10).pack(side="left", padx=4)
        tk.Button(btns, text="Cancelar", command=cancelar, font=FONT_N, width=10).pack(side="left", padx=4)
        entry.bind("<Return>", lambda _: ok())
        dlg.wait_window()
        return resultado[0] if resultado else ""

    def _quitar_banco(self):
        sel = self._lst_bancos.curselection()
        if not sel:
            return
        idx = sel[0]
        nombre = self._bancos[idx][0]
        if messagebox.askyesno("Quitar banco", f"Quitar '{nombre}'?"):
            del self._bancos[idx]
            self._refrescar_lista_bancos()

    def _leer_excel_banco(self, archivo: str) -> list[MovBanco]:
        import openpyxl
        wb = openpyxl.load_workbook(archivo, data_only=True)
        ws = wb.active
        hdrs = [str(ws.cell(row=1, column=c).value or "").upper()
                for c in range(1, ws.max_column + 1)]

        def col(name):
            return next((i + 1 for i, h in enumerate(hdrs) if name in h), None)

        c_fecha  = col("FECHA")
        c_folio  = col("FOLIO")
        c_desc   = col("DESCRIPCI")
        c_retiro = col("RETIRO") or col("CARGO")
        c_dep    = col("DEPOSITO") or col("ABONO")
        c_saldo  = col("SALDO")

        movs = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row[0]:
                continue
            movs.append(MovBanco(
                fecha       = str(row[(c_fecha  or 1) - 1] or ""),
                folio       = str(row[(c_folio  or 2) - 1] or ""),
                descripcion = str(row[(c_desc   or 3) - 1] or ""),
                retiro      = float(row[(c_retiro or 4) - 1] or 0),
                deposito    = float(row[(c_dep   or 5) - 1] or 0),
                saldo       = float(row[(c_saldo or 6) - 1] or 0),
            ))
        return movs

    # ── Carpetas / Archivos XML ───────────────────────────────────────────────

    def _actualizar_count_xml(self):
        n = len(self._xml_items)
        self._lbl_carpetas_count.config(text=f"{n} elemento(s) seleccionado(s)")

    def _agregar_xmls(self):
        """Seleccionar uno o varios archivos .xml directamente."""
        archivos = filedialog.askopenfilenames(
            title="Seleccionar archivos XML (CFDIs)",
            filetypes=[("Archivos XML", "*.xml"), ("Todos", "*.*")],
        )
        agregados = 0
        for a in archivos:
            if a not in self._xml_items:
                self._xml_items.append(a)
                self._lst_carpetas.insert(tk.END, Path(a).name)
                agregados += 1
        if agregados:
            self._actualizar_count_xml()

    def _agregar_carpeta(self):
        carpeta = filedialog.askdirectory(title="Seleccionar carpeta con XMLs")
        if carpeta and carpeta not in self._xml_items:
            self._xml_items.append(carpeta)
            self._lst_carpetas.insert(tk.END, carpeta)
            self._actualizar_count_xml()

    def _quitar_xml(self):
        sel = self._lst_carpetas.curselection()
        if not sel:
            return
        idx = sel[0]
        del self._xml_items[idx]
        self._lst_carpetas.delete(idx)
        self._actualizar_count_xml()

    def _limpiar_carpetas(self):
        self._xml_items.clear()
        self._lst_carpetas.delete(0, tk.END)
        self._lbl_carpetas_count.config(text="0 elementos seleccionados")

    # ── Destino ───────────────────────────────────────────────────────────────

    def _elegir_destino(self):
        hoy = datetime.now().strftime("%Y%m%d")
        nombre_default = f"Conciliacion_MultiBanco_{hoy}.xlsx"
        archivo = filedialog.asksaveasfilename(
            title="Guardar conciliacion como...",
            defaultextension=".xlsx",
            initialfile=nombre_default,
            filetypes=[("Excel", "*.xlsx")],
        )
        if archivo:
            self._destino_var.set(archivo)

    # ── Buscador de suma ──────────────────────────────────────────────────────

    def _abrir_buscador_suma(self):
        """
        Abre un diálogo donde el usuario ingresa un importe y se buscan
        combinaciones de CFDIs (recibidos o emitidos) que sumen exactamente
        ese valor.  Usa DP subset-sum igual que el nivel 4 automático.
        """
        dlg = tk.Toplevel(self)
        dlg.title("Buscar CFDIs que sumen un importe")
        dlg.geometry("820x560")
        dlg.configure(bg=C_BG)
        dlg.grab_set()
        dlg.resizable(True, True)

        # ── Fila de controles ────────────────────────────────────────────────
        frm_ctrl = tk.Frame(dlg, bg=C_BG)
        frm_ctrl.pack(fill="x", padx=16, pady=(14, 4))

        tk.Label(frm_ctrl, text="Importe:", font=FONT_B,
                 bg=C_BG, fg=C_TEXT).grid(row=0, column=0, sticky="w", padx=(0, 6))
        importe_var = tk.StringVar()
        ent_importe = tk.Entry(frm_ctrl, textvariable=importe_var, width=16,
                               font=FONT_N, bg=C_WHITE, fg=C_TEXT,
                               relief="solid", bd=1)
        ent_importe.grid(row=0, column=1, padx=(0, 14))
        ent_importe.focus_set()

        tipo_var = tk.StringVar(value="retiro")
        tk.Radiobutton(frm_ctrl, text="Retiro (CFDIs recibidos)",
                       variable=tipo_var, value="retiro",
                       bg=C_BG, fg=C_TEXT, font=FONT_N,
                       activebackground=C_BG).grid(row=0, column=2, padx=(0, 6))
        tk.Radiobutton(frm_ctrl, text="Depósito (CFDIs emitidos)",
                       variable=tipo_var, value="deposito",
                       bg=C_BG, fg=C_TEXT, font=FONT_N,
                       activebackground=C_BG).grid(row=0, column=3, padx=(0, 14))

        btn_buscar = _btn(frm_ctrl, "Buscar", lambda: _buscar(),
                          bg=T.SUCCESS, fg=T.ACCENT_FG)
        btn_buscar.grid(row=0, column=4)

        # ── Etiqueta de estado ───────────────────────────────────────────────
        lbl_status = tk.Label(dlg, text="Ingresa el importe y presiona Buscar.",
                              font=FONT_S, bg=C_BG, fg=C_GRAY2, anchor="w")
        lbl_status.pack(fill="x", padx=16, pady=(0, 4))

        # ── Treeview ─────────────────────────────────────────────────────────
        cols = ("grupo", "uuid", "nombre", "rfc", "importe")
        tree = ttk.Treeview(dlg, columns=cols, show="headings", height=16)
        tree.heading("grupo",   text="#")
        tree.heading("uuid",    text="UUID / CFDI")
        tree.heading("nombre",  text="Nombre")
        tree.heading("rfc",     text="RFC")
        tree.heading("importe", text="Importe")
        tree.column("grupo",   width=28,  minwidth=24,  anchor="center", stretch=False)
        tree.column("uuid",    width=290, minwidth=180)
        tree.column("nombre",  width=210, minwidth=100)
        tree.column("rfc",     width=130, minwidth=80)
        tree.column("importe", width=100, minwidth=70,  anchor="e", stretch=False)

        frm_tree = tk.Frame(dlg, bg=C_BG)
        frm_tree.pack(fill="both", expand=True, padx=16, pady=(0, 4))
        sb_v = ttk.Scrollbar(frm_tree, orient="vertical",   command=tree.yview)
        sb_h = ttk.Scrollbar(frm_tree, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=sb_v.set, xscrollcommand=sb_h.set)
        tree.grid(row=0, column=0, sticky="nsew")
        sb_v.grid(row=0, column=1, sticky="ns")
        sb_h.grid(row=1, column=0, sticky="ew")
        frm_tree.rowconfigure(0, weight=1)
        frm_tree.columnconfigure(0, weight=1)

        tree.tag_configure("subtotal", background=T.ACCENT_SOFT,
                           foreground=T.TEXT, font=F.BTN_SM)
        tree.tag_configure("cfdi",     background=T.SURFACE)

        # ── Pie ──────────────────────────────────────────────────────────────
        lbl_total = tk.Label(dlg, text="", font=FONT_B, bg=C_BG,
                             fg=C_TEXT, anchor="e")
        lbl_total.pack(fill="x", padx=16, pady=(0, 2))
        _btn(dlg, "Cerrar", dlg.destroy, bg=C_GRAY, fg=C_TEXT).pack(pady=(0, 12))

        # ── Lógica DP subset-sum ─────────────────────────────────────────────
        def _dp_find_combos(pool, target_c: int,
                            max_combos: int = 10,
                            max_cfdi:   int = 5) -> list:
            """
            Itera hasta max_combos veces: cada vez corre DP sobre los CFDIs
            que aún no han sido usados por una combinación anterior.
            Devuelve lista de listas[CFDI].
            """
            resultados = []
            usados: set[str] = set()

            for _ in range(max_combos):
                avail = [c for c in pool
                         if c.uuid and c.uuid.lower() not in usados
                         and 0 < int(round(c.total * 100)) < target_c]
                if len(avail) < 1:
                    break

                vals = [int(round(c.total * 100)) for c in avail]
                if sum(vals) < target_c:
                    break

                # DP subset-sum
                dp: dict[int, list[int]] = {0: []}
                combo_idx: list[int] | None = None
                for idx, v in enumerate(vals):
                    for s, comb in list(dp.items()):
                        ns = s + v
                        if ns > target_c or ns in dp:
                            continue
                        nc = comb + [idx]
                        if len(nc) > max_cfdi:
                            continue
                        dp[ns] = nc
                        if ns == target_c:
                            combo_idx = nc
                            break
                    if combo_idx is not None:
                        break

                if combo_idx is None:
                    break

                combo = [avail[i] for i in combo_idx]
                resultados.append(combo)
                for c in combo:
                    usados.add(c.uuid.lower())

            return resultados

        def _buscar():
            raw = importe_var.get().strip().replace(",", ".")
            try:
                target_f = round(float(raw), 2)
            except ValueError:
                messagebox.showwarning("Importe inválido",
                    "Escribe el importe numérico, p. ej.  12 345.67",
                    parent=dlg)
                return
            if target_f <= 0:
                return

            # Obtener pool de CFDIs
            if self._ultimo_resultado:
                recibidos = self._ultimo_resultado["recibidos"]
                emitidos  = self._ultimo_resultado["emitidos"]
            else:
                rfc = self._rfc_var.get().strip().upper()
                if not rfc or not self._xml_items:
                    messagebox.showwarning(
                        "Sin datos",
                        "Carga los XMLs y el RFC, o ejecuta la conciliación primero.",
                        parent=dlg)
                    return
                lbl_status.config(text="Leyendo XMLs...", fg=C_GRAY2)
                dlg.update()
                try:
                    cfdis = leer_cfdis_carpetas(*self._xml_items)
                    recibidos, emitidos, _ = clasificar_cfdis(cfdis, rfc)
                except Exception as exc:
                    messagebox.showerror("Error al leer XMLs", str(exc), parent=dlg)
                    return

            pool = recibidos if tipo_var.get() == "retiro" else emitidos

            lbl_status.config(
                text=f"Buscando combinaciones que sumen  ${target_f:,.2f}  "
                     f"en {len(pool)} CFDIs...",
                fg=C_GRAY2)
            dlg.update()

            target_c = int(round(target_f * 100))
            combos = _dp_find_combos(pool, target_c)

            # Limpiar árbol
            for iid in tree.get_children():
                tree.delete(iid)

            if not combos:
                lbl_status.config(
                    text=f"No se encontró ninguna combinación que sume  ${target_f:,.2f}.",
                    fg=C_RED)
                lbl_total.config(text="")
                return

            lbl_status.config(
                text=f"{len(combos)} combinación(es) encontrada(s) "
                     f"que suman  ${target_f:,.2f}:",
                fg="#1B5E20")

            for idx, combo in enumerate(combos, 1):
                subtotal = sum(c.total for c in combo)
                grp_id = tree.insert(
                    "", tk.END,
                    values=(str(idx),
                            f"── Combinación {idx}  ({len(combo)} CFDI{'s' if len(combo)>1 else ''}) ──",
                            "",
                            "",
                            f"${subtotal:,.2f}"),
                    tags=("subtotal",),
                    open=True,
                )
                for c in combo:
                    nombre = (c.nombre_emisor or c.nombre_receptor or "")[:45]
                    rfc_c  = c.rfc_emisor  or c.rfc_receptor  or ""
                    tree.insert(grp_id, tk.END,
                                values=("", c.uuid, nombre, rfc_c,
                                        f"${c.total:,.2f}"),
                                tags=("cfdi",))

            lbl_total.config(text=f"Importe buscado:  ${target_f:,.2f}")

        # Permitir buscar con Enter desde el campo de importe
        ent_importe.bind("<Return>", lambda _e: _buscar())

    # ── Grupos RFC ────────────────────────────────────────────────────────────

    def _refrescar_lbl_grupos(self) -> None:
        rfc = self._rfc_var.get().strip().upper()
        if not rfc:
            self._lbl_grupos.config(text="0 grupos configurados")
            return
        try:
            gs = cargar_grupos(rfc)
            n_act = sum(1 for g in gs if g.activo)
            total = len(gs)
            if total == 0:
                txt = "0 grupos configurados"
            elif n_act == total:
                txt = f"{total} grupo{'s' if total != 1 else ''} activo{'s' if total != 1 else ''}"
            else:
                txt = f"{n_act}/{total} grupos activos"
            self._lbl_grupos.config(text=txt)
        except Exception:
            self._lbl_grupos.config(text="—")

    def _abrir_grupos_rfc(self) -> None:
        rfc = self._rfc_var.get().strip().upper()
        if not rfc:
            from tkinter import messagebox
            messagebox.showwarning("RFC requerido",
                                   "Ingresa el RFC de tu empresa primero.")
            return
        if not _GRUPOS_OK:
            from tkinter import messagebox
            messagebox.showerror("Módulo no disponible",
                                 "ventana_grupos_rfc.py no encontrado.")
            return
        VentanaGruposRFC(self, rfc_empresa=rfc,
                         on_close=self._refrescar_lbl_grupos)

    # ── Conciliar ─────────────────────────────────────────────────────────────

    def _conciliar(self):
        rfc = self._rfc_var.get().strip().upper()
        if not rfc:
            messagebox.showwarning("RFC requerido", "Ingresa el RFC de tu empresa.")
            return
        if not self._xml_items:
            messagebox.showwarning("XMLs requeridos",
                                   "Selecciona al menos un archivo XML o una carpeta con CFDIs.")
            return
        if not self._bancos:
            messagebox.showwarning("Sin bancos",
                                   "Agrega al menos un estado de cuenta bancario.")
            return
        destino = self._destino_var.get().strip()
        if not destino:
            messagebox.showwarning("Destino",
                                   "Selecciona donde guardar el archivo resultado.")
            return

        self._btn_conciliar.config(state="disabled", text="Procesando...")
        self._prog.start(10)
        self._log_clear()

        threading.Thread(
            target=self._ejecutar_conciliacion,
            args=(rfc, destino),
            daemon=True,
        ).start()

    def _fusionar_con_historial(self, anterior: dict,
                                 recibidos: list[CFDI], emitidos: list[CFDI]):
        """
        Combina los CFDIs recién cargados con los de la última conciliación
        guardada (dedup por UUID — clave única real del SAT; el recién
        cargado gana en caso de conflicto por ser el más fresco).

        También extrae los movimientos bancarios previos agrupados por
        nombre de banco, para que _armar_bancos_input los fusione con los
        estados de cuenta cargados en esta sesión.

        Retorna (recibidos_fusionados, emitidos_fusionados,
                 bancos_previos_por_nombre, n_r_nuevos, n_e_nuevos, n_mov_previos).
        """
        def _fusionar_cfdis(nuevos: list[CFDI], previos: list[CFDI]) -> tuple[list[CFDI], int]:
            por_uuid = {c.uuid.lower(): c for c in previos if c.uuid}
            n_nuevos_reales = sum(
                1 for c in nuevos if c.uuid and c.uuid.lower() not in por_uuid
            )
            for c in nuevos:
                if c.uuid:
                    por_uuid[c.uuid.lower()] = c   # el recién cargado gana
            return list(por_uuid.values()), n_nuevos_reales

        recibidos_f, n_r = _fusionar_cfdis(recibidos, anterior.get("recibidos", []))
        emitidos_f,  n_e = _fusionar_cfdis(emitidos,  anterior.get("emitidos", []))

        bancos_previos: dict[str, list[MovBanco]] = {}
        for res in anterior.get("resultados", []):
            movs = [mc.banco for mc in res.conciliados] + res.sin_cfdi
            bancos_previos.setdefault(res.nombre, []).extend(movs)
        n_mov_previos = sum(len(v) for v in bancos_previos.values())

        return recibidos_f, emitidos_f, bancos_previos, n_r, n_e, n_mov_previos

    def _armar_bancos_input(self, bancos_previos: dict[str, list[MovBanco]]):
        """
        Combina los bancos cargados en esta sesión (self._bancos) con los
        movimientos previos del historial (mismo nombre de banco = mismo
        origen). Deduplica por (fecha, folio, descripcion, retiro, deposito,
        saldo) — un estado de cuenta no trae un ID de transacción real, así
        que esta es la clave natural más cercana a uno.

        Bancos que solo existen en el historial (no recargados esta sesión)
        se conservan tal cual, para no perder ese periodo.
        """
        def _clave(mb: MovBanco) -> tuple:
            return (mb.fecha, mb.folio, mb.descripcion, mb.retiro, mb.deposito, mb.saldo)

        nombres_sesion = {nombre for nombre, _, _ in self._bancos}
        bancos_input: list[tuple[str, list[MovBanco]]] = []
        colores: dict[str, str] = {}

        for nombre, color, movs in self._bancos:
            colores[nombre] = color
            vistos = {_clave(m) for m in movs}
            combinados = list(movs)
            for prev in bancos_previos.get(nombre, []):
                if _clave(prev) not in vistos:
                    combinados.append(prev)
                    vistos.add(_clave(prev))
            bancos_input.append((nombre, combinados))

        for nombre, movs in bancos_previos.items():
            if nombre not in nombres_sesion:
                bancos_input.append((nombre, movs))

        return bancos_input, colores

    def _ejecutar_conciliacion(self, rfc: str, destino: str):
        try:
            # 1. Leer CFDIs
            self._log_write(f"Leyendo XMLs de {len(self._xml_items)} elemento(s)...\n")
            cfdis = leer_cfdis_carpetas(*self._xml_items)
            self._log_write(f"   -> {len(cfdis)} CFDIs encontrados\n")

            recibidos, emitidos, descartados = clasificar_cfdis(cfdis, rfc)
            nom_count = sum(1 for c in recibidos if c.tipo == "N")
            self._log_write(
                f"   -> {len(recibidos)} recibidos ({nom_count} nomina)  |  {len(emitidos)} emitidos"
                + (f"  |  {len(descartados)} descartados (RFC no coincide con {rfc})" if descartados else "")
                + "\n"
            )

            # 1b. Fusionar con la última conciliación guardada de este RFC —
            # así no se pierde lo ya conciliado en sesiones anteriores.
            bancos_previos: dict[str, list[MovBanco]] = {}
            if self._var_fusionar.get():
                anterior = obtener_ultima_conciliacion(rfc)
                if anterior:
                    recibidos, emitidos, bancos_previos, n_r, n_e, n_m = \
                        self._fusionar_con_historial(anterior, recibidos, emitidos)
                    self._log_write(
                        f"Fusionado con conciliación anterior del historial: "
                        f"+{n_r} CFDI(s) recibido(s), +{n_e} emitido(s), "
                        f"+{n_m} movimiento(s) bancario(s) previos.\n"
                    )
                else:
                    self._log_write("Sin conciliación previa para este RFC — corrida nueva.\n")

            # 2. Armar lista de bancos [(nombre, movimientos)], fusionando con
            #    los movimientos previos del mismo nombre de banco (dedup)
            bancos_input, colores = self._armar_bancos_input(bancos_previos)

            self._log_write(f"Bancos a conciliar: {len(bancos_input)}\n")
            for nombre, movs in bancos_input:
                self._log_write(f"   {nombre}: {len(movs)} movimientos\n")

            # 3. Conciliar multi-banco (con Grupos RFC si están configurados)
            grupos = cargar_grupos(rfc)
            if grupos:
                activos = [g for g in grupos if g.activo]
                self._log_write(
                    f"Grupos RFC: {len(activos)} activo(s) — "
                    + ", ".join(f"'{g.nombre}' ({len(g.rfcs)} RFCs)" for g in activos)
                    + "\n"
                )
            else:
                grupos = None

            self._log_write("Conciliando...\n")
            resultados = conciliar_multibanco(bancos_input, recibidos, emitidos,
                                              grupos=grupos)

            total_conc  = sum(len(r.conciliados) for r in resultados)
            total_pend  = sum(len(r.sin_cfdi)    for r in resultados)
            total_rev   = sum(sum(1 for mc in r.conciliados if mc.revisar) for r in resultados)

            for res in resultados:
                n_conc = len(res.conciliados)
                n_pend = len(res.sin_cfdi)
                self._log_write(f"   {res.nombre}: {n_conc} conciliados, {n_pend} sin CFDI\n")
            self._log_write(
                f"   TOTAL: {total_conc} conciliados, {total_pend} sin CFDI"
                f"{', ' + str(total_rev) + ' requieren revision' if total_rev else ''}\n"
            )

            # 4. Exportar
            self._log_write("Generando Excel...\n")
            exportar_conciliacion_multi(
                resultados      = resultados,
                todos_recibidos = recibidos,
                todos_emitidos  = emitidos,
                output_path     = destino,
                colores_banco   = colores,
                rfc_empresa     = rfc,
            )

            # Guardar estado para módulo de pólizas
            self._ultimo_resultado = {
                "resultados": resultados,
                "recibidos":  recibidos,
                "emitidos":   emitidos,
                "destino":    destino,
                "rfc":        rfc,
            }

            # Guardar en el historial persistente (no debe interrumpir el flujo
            # si falla — el Excel ya generado sigue siendo válido)
            try:
                guardar_conciliacion(rfc, resultados, recibidos, emitidos, destino)
                self._log_write("Guardado en historial de conciliaciones.\n")
            except Exception as e:
                self._log_write(f"Aviso: no se pudo guardar en el historial ({e}).\n")

            self._log_write(f"\nArchivo generado:\n   {destino}\n")

            self.after(0, lambda: self._habilitar_boton_polizas(total_conc))
            self.after(0, lambda: messagebox.showinfo(
                "Conciliacion completa",
                f"Archivo generado exitosamente.\n\n"
                f"Bancos procesados: {len(resultados)}\n"
                f"Conciliados: {total_conc}\n"
                f"Sin CFDI: {total_pend}\n"
                + (f"Requieren revision: {total_rev}\n" if total_rev else "")
                + f"\n{destino}"
            ))
            self.after(0, lambda: os.startfile(destino))

        except Exception as e:
            import traceback
            msg = traceback.format_exc()
            self._log_write(f"\nError:\n{msg}\n")
            self.after(0, lambda: messagebox.showerror("Error", str(e)))

        finally:
            self.after(0, self._prog.stop)
            self.after(0, lambda: self._btn_conciliar.config(
                state="normal", text="CONCILIAR AHORA"
            ))

    # ── Catálogos COI ─────────────────────────────────────────────────────────

    def _rfc_actual(self) -> str:
        return self._rfc_var.get().strip().upper()

    def _refrescar_status_catalogos(self):
        """Actualiza los labels de estado de catálogos según el RFC ingresado."""
        if not _COI_DISPONIBLE:
            return
        try:
            from catalogo_cuentas  import cargar_catalogo
            from catalogo_terceros import cargar_terceros
        except ImportError:
            return

        rfc = self._rfc_actual()
        if not rfc:
            self._lbl_cuentas_status.config(text="Ingresa el RFC primero", fg=C_GRAY2)
            self._lbl_terceros_status.config(text="Ingresa el RFC primero", fg=C_GRAY2)
            return

        cuentas  = cargar_catalogo(rfc)
        terceros = cargar_terceros(rfc)

        if cuentas:
            self._lbl_cuentas_status.config(
                text=f"{len(cuentas)} cuentas cargadas", fg="#1B5E20")
        else:
            self._lbl_cuentas_status.config(text="No cargado", fg=C_RED)

        if terceros:
            sin_cta = sum(1 for t in terceros if not t.cuenta_contable)
            txt = f"{len(terceros)} terceros"
            txt += f"  ({sin_cta} sin cuenta)" if sin_cta else "  (todos con cuenta)"
            col = "#E65100" if sin_cta else "#1B5E20"
            self._lbl_terceros_status.config(text=txt, fg=col)
        else:
            self._lbl_terceros_status.config(text="No cargado", fg=C_RED)

    def _importar_cuentas_catalogo(self):
        """Importa el catálogo de cuentas desde un Excel de COI."""
        if not _COI_DISPONIBLE:
            messagebox.showerror("Modulo no disponible",
                                 "catalogo_cuentas.py no encontrado.")
            return
        rfc = self._rfc_actual()
        if not rfc:
            messagebox.showwarning("RFC requerido",
                                   "Escribe el RFC de la empresa antes de importar.")
            return
        from tkinter import filedialog
        path = filedialog.askopenfilename(
            title="Seleccionar Excel de Catalogo de Cuentas COI",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")],
        )
        if not path:
            return
        try:
            from catalogo_cuentas import importar_desde_excel
            ok, msg, _ = importar_desde_excel(rfc, path)
            if ok:
                messagebox.showinfo("Catalogo importado", msg)
            else:
                messagebox.showerror("Error al importar", msg)
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            self._refrescar_status_catalogos()

    def _poblar_terceros_catalogo(self):
        """Auto-puebla clientes/proveedores desde los XMLs ya cargados."""
        if not _COI_DISPONIBLE:
            return
        rfc = self._rfc_actual()
        if not rfc:
            messagebox.showwarning("RFC requerido",
                                   "Escribe el RFC de la empresa primero.")
            return
        if not self._xml_items:
            messagebox.showwarning("Sin XMLs",
                                   "Agrega carpetas o archivos XML antes de poblar.")
            return
        try:
            from cfdi_parser        import leer_cfdis_carpetas, clasificar_cfdis
            from catalogo_terceros  import poblar_desde_cfdis
            cfdis = leer_cfdis_carpetas(*self._xml_items)
            recibidos, emitidos, _ = clasificar_cfdis(cfdis, rfc)
            nc, np = poblar_desde_cfdis(rfc, recibidos, emitidos)
            messagebox.showinfo(
                "Terceros actualizados",
                f"Nuevos clientes:     {nc}\n"
                f"Nuevos proveedores:  {np}\n\n"
                "Usa 'Gestionar' para revisar y asignar cuentas contables.",
            )
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            self._refrescar_status_catalogos()

    def _gestionar_catalogos(self):
        """Abre VentanaPolizas directo en modo catálogos (sin conciliación previa)."""
        if not _COI_DISPONIBLE:
            messagebox.showerror("Modulo no disponible",
                                 "ventana_polizas.py no encontrado.")
            return
        rfc = self._rfc_actual()
        if not rfc:
            messagebox.showwarning("RFC requerido",
                                   "Escribe el RFC de la empresa primero.")
            return
        # Abre la ventana de pólizas con listas vacías — solo para gestionar catálogos
        u = self._ultimo_resultado or {}
        VentanaPolizas(
            parent            = self,
            rfc_empresa       = rfc,
            todos_recibidos   = u.get("recibidos", []),
            todos_emitidos    = u.get("emitidos", []),
            resultados        = u.get("resultados", []),
            output_excel_path = u.get("destino", ""),
        )
        # Al cerrar refrescar los contadores
        self.after(500, self._refrescar_status_catalogos)

    # ── Módulo Pólizas COI ────────────────────────────────────────────────────

    def _abrir_polizas(self):
        if not _COI_DISPONIBLE:
            messagebox.showerror(
                "Módulo no disponible",
                "No se pudo cargar ventana_polizas.py.\n"
                "Verifica que el archivo exista junto a ventana_conciliacion.py.",
            )
            return
        if not self._ultimo_resultado:
            messagebox.showwarning(
                "Sin conciliación",
                "Ejecuta la conciliación primero antes de generar pólizas.",
            )
            return
        u = self._ultimo_resultado
        VentanaPolizas(
            parent              = self,
            rfc_empresa         = u["rfc"],
            todos_recibidos     = u["recibidos"],
            todos_emitidos      = u["emitidos"],
            resultados          = u["resultados"],
            output_excel_path   = u["destino"],
        )
        self.after(500, self._refrescar_status_catalogos)

    def _habilitar_boton_polizas(self, n_conciliados: int) -> None:
        """Activa el botón de Pólizas COI tras conciliar o cargar del historial."""
        self._btn_polizas.config(
            state="normal",
            text="Polizas COI  ->  Haz clic para generar",
            bg="#37474F", fg="#FFFFFF",
        )
        self._lbl_coi_estado.config(
            text=f"{n_conciliados} mov. listos",
            fg="#1B5E20",
        )

    # ── Historial de conciliaciones ──────────────────────────────────────────

    def _abrir_historial(self):
        """Lista las conciliaciones guardadas para el RFC actual y permite
        reabrir una sin reprocesar XMLs ni estados de cuenta."""
        rfc = self._rfc_var.get().strip().upper()
        if not rfc:
            messagebox.showwarning("RFC requerido",
                                   "Escribe el RFC de la empresa primero.")
            return

        registros = listar_conciliaciones(rfc)

        dlg = tk.Toplevel(self)
        dlg.title(f"Conciliaciones guardadas — {rfc}")
        dlg.geometry("640x420")
        dlg.configure(bg=C_BG)
        dlg.grab_set()
        dlg.resizable(True, True)

        tk.Frame(dlg, bg=T.ACCENT, height=3).pack(fill="x")

        if not registros:
            tk.Label(
                dlg,
                text="No hay conciliaciones guardadas para este RFC todavía.\n"
                     "Se guardan automáticamente cada vez que conciliar termina con éxito.",
                font=FONT_N, bg=C_BG, fg=C_GRAY2, justify="left",
            ).pack(padx=16, pady=20, anchor="w")
            _btn(dlg, "Cerrar", dlg.destroy,
                 bg=T.SURFACE, fg=T.TEXT,
                 highlightthickness=1, highlightbackground=T.BORDER
                 ).pack(pady=(0, 12))
            return

        cols   = ("fecha", "conciliados", "pendientes", "recibidos", "emitidos")
        hdrs   = ("Fecha", "Conciliados", "Sin CFDI", "CFDIs Recibidos", "CFDIs Emitidos")
        widths = (150, 100, 90, 120, 120)
        frm = tk.Frame(dlg, bg=C_BG)
        frm.pack(fill="both", expand=True, padx=16, pady=(12, 8))
        frm.columnconfigure(0, weight=1)
        frm.rowconfigure(0, weight=1)

        tree = ttk.Treeview(frm, columns=cols, show="headings", selectmode="browse")
        for c, h, w in zip(cols, hdrs, widths):
            tree.heading(c, text=h)
            tree.column(c, width=w, anchor="w")
        sb = ttk.Scrollbar(frm, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")

        id_por_fila = {}
        for reg in registros:
            iid = tree.insert("", "end", values=(
                reg["fecha_guardado"], reg["n_conciliados"], reg["n_pendientes"],
                reg["n_recibidos"], reg["n_emitidos"],
            ))
            id_por_fila[iid] = reg["id"]

        def _abrir_seleccionado():
            sel = tree.selection()
            if not sel:
                messagebox.showwarning("Selecciona una fila",
                                       "Elige una conciliación de la lista.",
                                       parent=dlg)
                return
            id_registro = id_por_fila[sel[0]]
            fecha_sel   = tree.item(sel[0])["values"][0]
            datos = cargar_conciliacion(id_registro)
            if not datos:
                messagebox.showerror(
                    "No se pudo abrir",
                    "El registro está dañado o ya no existe.",
                    parent=dlg,
                )
                return
            self._ultimo_resultado = datos
            n_conc = sum(len(r.conciliados) for r in datos["resultados"])
            self._habilitar_boton_polizas(n_conc)
            self._log_write(
                f"\nConciliación del {fecha_sel} "
                f"cargada desde el historial ({n_conc} mov. conciliados).\n"
            )
            dlg.destroy()

        def _eliminar_seleccionado():
            sel = tree.selection()
            if not sel:
                return
            if not messagebox.askyesno("Confirmar", "¿Eliminar esta conciliación del historial?",
                                       parent=dlg):
                return
            eliminar_conciliacion(id_por_fila[sel[0]])
            tree.delete(sel[0])

        bf = tk.Frame(dlg, bg=C_BG)
        bf.pack(fill="x", padx=16, pady=(0, 14))
        _btn(bf, "📂  Abrir", _abrir_seleccionado,
             bg=T.SUCCESS, fg=T.ACCENT_FG).pack(side="left")
        _btn(bf, "🗑  Eliminar", _eliminar_seleccionado,
             bg=T.SURFACE, fg=T.DANGER,
             highlightthickness=1, highlightbackground=T.DANGER
             ).pack(side="left", padx=(8, 0))
        _btn(bf, "Cerrar", dlg.destroy,
             bg=T.SURFACE, fg=T.TEXT,
             highlightthickness=1, highlightbackground=T.BORDER
             ).pack(side="right")

        tree.bind("<Double-1>", lambda e: _abrir_seleccionado())

    # ── Log helpers ───────────────────────────────────────────────────────────

    def _log_write(self, texto: str):
        def _do():
            self._log.config(state="normal")
            self._log.insert(tk.END, texto)
            self._log.see(tk.END)
            self._log.config(state="disabled")
        self.after(0, _do)

    def _log_clear(self):
        self._log.config(state="normal")
        self._log.delete("1.0", tk.END)
        self._log.config(state="disabled")

    # ── API para llamar desde ConvertidorSantander ────────────────────────────

    def actualizar_movimientos(self, movimientos, info_cuenta=None):
        """
        Llamar después de convertir un PDF para agregar/reemplazar el primer banco.
        Si ya existe un banco con el mismo nombre lo reemplaza; si no, lo agrega.
        """
        self.info_cuenta = info_cuenta
        rfc = getattr(info_cuenta, "rfc", "") if info_cuenta else ""
        if rfc:
            self._rfc_var.set(rfc)

        nombre = (getattr(info_cuenta, "banco", None) or "Banco") if info_cuenta else "Banco"
        color  = (getattr(info_cuenta, "color_hex", None) or _COLORES_BANCO[0]) if info_cuenta else _COLORES_BANCO[0]
        movs   = self._normalizar_movs(movimientos or [])

        # Reemplazar si ya existe el mismo nombre, sino agregar al inicio
        for i, (n, c, _) in enumerate(self._bancos):
            if n == nombre:
                self._bancos[i] = (nombre, color, movs)
                self._refrescar_lista_bancos()
                return

        self._bancos.insert(0, (nombre, color, movs))
        self._refrescar_lista_bancos()
