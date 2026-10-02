"""Demo visual de las 3 pantallas actualizadas — sin login."""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))

import tkinter as tk
from tkinter import ttk, messagebox

# ── Paleta (copiada de ConvertidorSantander.py) ──────────────────────────────
C_BG    = "#EFEFEF"; C_WHITE = "#FFFFFF"; C_ACCENT= "#99CCFF"
C_GRAY  = "#C5C5C5"; C_GRAY2 = "#666666"; C_TEXT  = "#1A1A1A"
C_GREEN = "#33CC33"; C_RED   = "#CC2222"; C_BLUE  = "#2B5BB5"
FONT_HEAD   = ("Segoe UI", 10, "bold")
FONT_NORMAL = ("Segoe UI", 9)
FONT_SMALL  = ("Segoe UI", 8)
C_WHITE_FG  = "#FFFFFF"

root = tk.Tk()
root.title("DEMO — Pantallas actualizadas con CSF/Régimen")
root.geometry("900x680")
root.configure(bg=C_BG)

nb = ttk.Notebook(root)
nb.pack(fill="both", expand=True, padx=10, pady=10)

# ══════════════════════════════════════════════════════════════════════════════
#  DEMO 1: Tab Convertidor PDF→Excel  (panel derecho)
# ══════════════════════════════════════════════════════════════════════════════
tab1 = tk.Frame(nb, bg=C_BG); nb.add(tab1, text="  🏦  Convertidor PDF→Excel  ")

action = tk.Frame(tab1, bg=C_ACCENT, pady=6); action.pack(fill="x")
tk.Button(action, text="📂  Agregar PDFs", bg=C_ACCENT, fg=C_TEXT,
          font=("Segoe UI", 11, "bold"), relief="flat", padx=20, pady=8).pack(side="left", padx=(16, 8))
tk.Button(action, text="🗑  Limpiar lista", bg=C_GRAY, fg=C_TEXT,
          font=FONT_NORMAL, relief="flat", padx=12, pady=8).pack(side="left")
tk.Button(action, text="▶  CONVERTIR A EXCEL", bg=C_GREEN, fg=C_WHITE,
          font=("Segoe UI", 11, "bold"), relief="flat", padx=24, pady=8).pack(side="right", padx=16)

body = tk.Frame(tab1, bg=C_BG); body.pack(fill="both", expand=True, padx=12, pady=6)

# Tabla (izquierda)
cols = ("Archivo", "Banco", "Estado", "Movs")
tree = ttk.Treeview(body, columns=cols, show="headings", height=10)
for c, w in zip(cols, [300, 130, 110, 70]):
    tree.heading(c, text=c, anchor="w"); tree.column(c, width=w, anchor="w")
tree.pack(side="left", fill="both", expand=True)

# Panel derecho — CON NUEVO BOTÓN CSF
right = tk.Frame(body, bg=C_BG, width=200); right.pack(side="right", fill="y", padx=(10, 0))
right.pack_propagate(False)

def _btn_r(txt, bg=C_GRAY, fg=C_TEXT):
    tk.Button(right, text=txt, bg=bg, fg=fg, font=FONT_NORMAL,
              relief="flat", padx=8, pady=6, cursor="hand2").pack(fill="x", pady=(0, 4))

_btn_r("🏢 Empresas / RFC",   bg=C_ACCENT, fg=C_TEXT)
_btn_r("📄 Subir CSF / CIF",  bg="#5C6BC0", fg=C_WHITE_FG)   # ← NUEVO
_btn_r("📋 Bitácora",          bg=C_ACCENT, fg=C_TEXT)
tk.Frame(right, bg=C_GRAY, height=1).pack(fill="x", pady=6)
tk.Label(right, text="Carpeta de salida:", bg=C_BG, fg=C_TEXT, font=FONT_HEAD, anchor="w").pack(fill="x")
tk.Entry(right, font=FONT_SMALL, bg=C_GRAY, fg=C_GRAY2, relief="flat", bd=4,
         state="readonly").pack(fill="x", pady=4)
_btn_r("📂 Elegir carpeta salida")
_btn_r("↺ Restablecer carpeta")
tk.Frame(right, bg=C_GRAY, height=1).pack(fill="x", pady=6)
tk.Checkbutton(right, text="Abrir Excel al terminar",
               bg=C_BG, fg=C_TEXT, font=FONT_SMALL, cursor="hand2").pack(anchor="w")

# Destacar el botón nuevo con etiqueta
lbl_new1 = tk.Label(right, text="▲ NUEVO", bg=C_BG, fg="#5C6BC0",
                    font=("Segoe UI", 7, "bold"))
lbl_new1.pack()

# ══════════════════════════════════════════════════════════════════════════════
#  DEMO 2: Catálogo Empresas — con columna Régimen y botón Subir CSF
# ══════════════════════════════════════════════════════════════════════════════
tab2 = tk.Frame(nb, bg=C_BG); nb.add(tab2, text="  🏢  Catálogo Empresas (actualizado)  ")

hdr2 = tk.Frame(tab2, bg=C_ACCENT, pady=12); hdr2.pack(fill="x")
tk.Label(hdr2, text="Catálogo de Empresas", bg=C_ACCENT, fg=C_TEXT,
         font=("Segoe UI", 11, "bold")).pack(side="left", padx=16)
tk.Label(hdr2, text="RFC · Nombre · Régimen fiscal (CSF)",
         bg=C_ACCENT, fg=C_GRAY2, font=FONT_SMALL).pack(side="left")

tf2 = tk.Frame(tab2, bg=C_BG); tf2.pack(fill="both", expand=True, padx=14, pady=10)
cols2 = ("RFC", "Nombre de Empresa", "Régimen(es) SAT")   # ← NUEVA COLUMNA
tree2 = ttk.Treeview(tf2, columns=cols2, show="headings", height=8, selectmode="browse")
for c, w in zip(cols2, [140, 240, 220]):
    tree2.heading(c, text=c, anchor="w"); tree2.column(c, width=w, minwidth=80)
tree2.tag_configure("sin_csf", foreground="#9E9E9E")
sb2 = ttk.Scrollbar(tf2, orient="vertical", command=tree2.yview)
tree2.configure(yscrollcommand=sb2.set)
tree2.pack(side="left", fill="both", expand=True)
sb2.pack(side="right", fill="y")

# Datos demo
tree2.insert("", "end", values=("STE120306P75", "STERIMEX",
             "Régimen Simplificado de Confianza (626)"))
tree2.insert("", "end", values=("GCE210503AB1", "GCEASA",
             "General de Ley Personas Morales (601)"))
tree2.insert("", "end", values=("AAA010101AAA", "EMPRESA SIN CSF",
             "—  sin CSF"), tags=("sin_csf",))

form2 = tk.Frame(tab2, bg=C_WHITE, padx=14, pady=10); form2.pack(fill="x")
for i, (lb, val, fnt) in enumerate([
    ("RFC:",        "STE120306P75",  ("Consolas", 10, "bold")),
    ("Nombre:",     "STERIMEX",      FONT_NORMAL),
]):
    tk.Label(form2, text=lb, bg=C_WHITE, fg=C_TEXT, font=FONT_HEAD).grid(row=i, column=0, sticky="w", pady=3)
    ef = tk.Frame(form2, bg=C_GRAY); ef.grid(row=i, column=1, sticky="ew", padx=(8, 20), pady=3)
    e = tk.Entry(ef, font=fnt, bg=C_WHITE, fg=C_TEXT, relief="flat", bd=5); e.pack(fill="x")
    e.insert(0, val)

tk.Label(form2, text="", bg=C_WHITE, fg=C_RED, font=FONT_SMALL).grid(row=2, column=1, sticky="w", padx=8)
tk.Label(form2, text="Régimen(es):", bg=C_WHITE, fg=C_TEXT,   # ← NUEVO
         font=FONT_HEAD).grid(row=3, column=0, sticky="nw", pady=(6, 3))
reg_lbl = tk.Label(form2, bg=C_WHITE, fg="#2E7D32",
                   font=FONT_SMALL, anchor="w", justify="left",
                   text="  • Régimen Simplificado de Confianza (626)")
reg_lbl.grid(row=3, column=1, sticky="w", padx=8, pady=(6, 3))
form2.columnconfigure(1, weight=1)

bf2 = tk.Frame(tab2, bg=C_BG, padx=14, pady=8); bf2.pack(fill="x")
for txt, bg, fg in [
    ("➕ Nuevo",      C_BLUE,    C_WHITE_FG),
    ("💾 Guardar",    C_GREEN,   C_WHITE_FG),
    ("📄 Subir CSF",  "#5C6BC0", C_WHITE_FG),   # ← NUEVO
    ("🗑 Eliminar",   C_RED,     C_WHITE_FG),
    ("✖ Cerrar",     C_GRAY,    C_TEXT),
]:
    tk.Button(bf2, text=txt, bg=bg, fg=fg, font=FONT_NORMAL,
              relief="flat", padx=12, pady=6, cursor="hand2").pack(side="left", padx=(0, 8))

# ══════════════════════════════════════════════════════════════════════════════
#  DEMO 3: Conciliación — PASO 3 con régimen y botón CSF
# ══════════════════════════════════════════════════════════════════════════════
tab3 = tk.Frame(nb, bg=C_BG); nb.add(tab3, text="  📊  Conciliación CFDI (actualizado)  ")

hdr3 = tk.Frame(tab3, bg=C_ACCENT, pady=8); hdr3.pack(fill="x")
tk.Label(hdr3, text="📊  Conciliación Bancaria ↔ CFDI",
         bg=C_ACCENT, fg=C_TEXT, font=("Segoe UI", 12, "bold")).pack(side="left", padx=16)

tk.Label(tab3, text="[ PASO 1 y PASO 2 sin cambios — se muestran solo las adiciones del PASO 3 ]",
         bg=C_BG, fg=C_GRAY2, font=FONT_SMALL).pack(pady=8)

cc = tk.LabelFrame(tab3, text=" PASO 3 — Configuración y Exportar ",
                   bg=C_WHITE, fg=C_BLUE, font=FONT_HEAD, relief="solid", bd=1)
cc.pack(fill="x", padx=12, pady=(0, 6))
cc.columnconfigure(1, weight=1); cc.columnconfigure(3, weight=2); cc.columnconfigure(4, weight=1)

tk.Label(cc, text="Empresa:", bg=C_WHITE, fg=C_TEXT,
         font=FONT_HEAD).grid(row=0, column=0, sticky="w", padx=(12, 4), pady=10)

combo_var = tk.StringVar(value="STE120306P75  —  STERIMEX")
combo = ttk.Combobox(cc, textvariable=combo_var, state="readonly", width=40, font=FONT_NORMAL,
                     values=["STE120306P75  —  STERIMEX", "GCE210503AB1  —  GCEASA"])
combo.grid(row=0, column=1, sticky="w", pady=10)

# Botones: recargar + CSF (NUEVOS)
btn_f = tk.Frame(cc, bg=C_WHITE); btn_f.grid(row=0, column=2, padx=(4, 0), pady=10)
tk.Button(btn_f, text="🔄", bg=C_GRAY, fg=C_TEXT, font=FONT_SMALL,
          relief="flat", padx=6, pady=4, cursor="hand2").pack(side="left", padx=(0, 4))
tk.Button(btn_f, text="📄 CSF", bg="#5C6BC0", fg=C_WHITE_FG, font=FONT_SMALL,   # ← NUEVO
          relief="flat", padx=8, pady=4, cursor="hand2").pack(side="left")

# RFC label
tk.Label(cc, text="STE120306P75", bg=C_WHITE, fg=C_GRAY2,
         font=("Consolas", 8)).grid(row=1, column=1, sticky="w", pady=(0, 2))

# Régimen label (NUEVO)
tk.Label(cc, text="  ✔  Régimen(es): Régimen Simplificado de Confianza (626)",
         bg=C_WHITE, fg="#2E7D32", font=FONT_SMALL, anchor="w"          # ← NUEVO
         ).grid(row=2, column=1, columnspan=2, sticky="w", pady=(0, 8))

tk.Label(cc, text="Guardar en:", bg=C_WHITE, fg=C_TEXT,
         font=FONT_HEAD).grid(row=0, column=3, sticky="w", padx=(20, 4), pady=10)
tk.Entry(cc, font=FONT_SMALL, bg=C_GRAY, fg=C_TEXT, relief="flat", bd=4
         ).grid(row=0, column=4, sticky="ew", pady=10)
tk.Button(cc, text="…", bg=C_GRAY, fg=C_TEXT, font=FONT_NORMAL,
          relief="flat", padx=8, pady=4, cursor="hand2"
          ).grid(row=0, column=5, padx=(4, 12), pady=10)

tk.Button(tab3,
          text="⚡  LEER PDFs  →  CONCILIAR  →  EXPORTAR EXCEL",
          bg=C_GREEN, fg=C_WHITE, font=("Segoe UI", 12, "bold"),
          relief="flat", pady=10, cursor="hand2").pack(fill="x", padx=12, pady=(0, 4))

root.mainloop()
