"""Lanzador de verificación: abre VentanaConciliacion directamente sin login."""
import sys
import os
sys.path.insert(0, os.path.dirname(__file__))

import tkinter as tk
from ventana_conciliacion import VentanaConciliacion

root = tk.Tk()
root.title("VERIFICACION — VentanaConciliacion")
root.geometry("700x750")
root.configure(bg="#EFEFEF")

frame = VentanaConciliacion(root)
frame.pack(fill="both", expand=True)

# Si hay empresa registrada con CSF, llenar el RFC para verlo
try:
    from empresas import listar_empresas, obtener_empresa
    empresas = listar_empresas()
    if empresas:
        rfc0, nombre0 = empresas[0]
        emp = obtener_empresa(rfc0)
        if emp and emp.get("regimenes"):
            frame._rfc_var.set(rfc0)
            print(f"RFC precargado: {rfc0} | Empresa: {nombre0}")
            print(f"Regimenes: {emp.get('regimenes')} | Nombres: {emp.get('regimenes_nombres')}")
        else:
            print(f"Empresa sin regimenes: {rfc0}")
    else:
        print("Sin empresas registradas — labels en estado vacío")
except Exception as e:
    print(f"Error cargando empresa: {e}")

root.mainloop()
