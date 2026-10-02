"""
construir_plantilla.py  —  paso de BUILD, ejecutar UNA vez en la maquina de
desarrollo. Crea `plantilla_papel_trabajo.xlsm` con la macro `PapelTrabajo`
ya embebida en el motor VBA, desde Python (via Excel COM / pywin32).

  python construir_plantilla.py

REQUISITOS (solo en la maquina de build, NO en la del cliente):
  - Microsoft Excel instalado.
  - pywin32:   pip install pywin32
  - Habilitar "Confiar en el acceso al modelo de objetos de proyectos de VBA"
    (AccessVBOM). Por registro (ajusta 16.0 a tu version de Office):
      reg add "HKCU\\Software\\Microsoft\\Office\\16.0\\Excel\\Security" /v AccessVBOM /t REG_DWORD /d 1 /f
    o manual: Excel > Archivo > Opciones > Centro de confianza > Config. de macros.

DESPUES de correr esto (una vez):
  Firmar el proyecto VBA con tu certificado de firma de codigo (ERSL Consultores):
  abre la plantilla, Alt+F11 > Tools > Digital Signature > elige el certificado.

  OJO: `keep_vba` por si solo NO basta para conservar la firma. openpyxl copia
  del original solo lo que casa con re.match("xl/vba|xl/ctrlProps|customUI|..."),
  asi que `xl/vbaProjectSignature*.bin` si pasa pero `xl/_rels/vbaProject.bin.rels`
  NO — y ese es el que enlaza el proyecto con sus firmas. Sin el, cada reporte
  salia con la macro sin firmar aunque la plantilla estuviera firmada.
  `exportar_conciliacion._restaurar_rels_firma()` lo reinyecta despues de guardar.

NOTA: pywin32 es dependencia SOLO de build. NO agregarla al requirements de
runtime del .exe — el cliente no necesita Excel COM ni compilar VBA; el reporte
solo INSERTA el VBA ya compilado (keep_vba) desde la plantilla empaquetada.
"""
from __future__ import annotations

from pathlib import Path

DIR = Path(__file__).parent
BAS = DIR / "PapelTrabajo.bas"
OUT = DIR / "plantilla_papel_trabajo.xlsm"

XL_MACRO_ENABLED = 52  # xlOpenXMLWorkbookMacroEnabled (.xlsm)


def construir() -> Path:
    try:
        import win32com.client as win32
    except ImportError:
        raise SystemExit("Falta pywin32. Instala con:  pip install pywin32")

    if not BAS.exists():
        raise SystemExit(f"No se encontro {BAS}")

    excel = win32.DispatchEx("Excel.Application")  # instancia dedicada
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Add()
        try:
            # Requiere AccessVBOM=1; si falla aqui, revisa REQUISITOS en la cabecera.
            wb.VBProject.VBComponents.Import(str(BAS))
            # Evento Workbook_Open en el modulo ThisWorkbook: auto-asigna el ID
            # (col A) a las filas nuevas cada vez que se ABRE el reporte.
            tw = wb.VBProject.VBComponents("ThisWorkbook").CodeModule
            tw.AddFromString(
                "Private Sub Workbook_Open()\r\n"
                "    On Error Resume Next\r\n"
                "    InicializarIDs\r\n"
                "End Sub\r\n"
                "\r\n"
                "Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)\r\n"
                "    On Error Resume Next\r\n"
                "    If Not Application.Intersect(Target, Sh.Columns(1)) Is Nothing Then\r\n"
                "        RestaurarIDs\r\n"
                "    End If\r\n"
                "End Sub\r\n"
            )
        except Exception as e:  # noqa: BLE001
            raise SystemExit(
                "No se pudo inyectar el VBA. Falta habilitar AccessVBOM "
                f"(ver cabecera de este archivo). Detalle: {e}"
            )
        if OUT.exists():
            OUT.unlink()
        # Excel COM no guarda de forma fiable en carpetas sincronizadas por
        # OneDrive/SharePoint (intenta la ruta web y truena con 0x800A03EC).
        # Este proyecto vive en OneDrive, así que guardamos en una ruta local
        # temporal y de ahí copiamos al destino con el sistema de archivos.
        import shutil
        import tempfile
        tmp = Path(tempfile.gettempdir()) / "_plantilla_papel_trabajo.xlsm"
        if tmp.exists():
            tmp.unlink()
        wb.SaveAs(str(tmp), FileFormat=XL_MACRO_ENABLED)
        wb.Close(SaveChanges=False)
        shutil.copyfile(tmp, OUT)
        tmp.unlink(missing_ok=True)
    finally:
        excel.Quit()
    return OUT


if __name__ == "__main__":
    ruta = construir()
    print(f"OK - plantilla creada: {ruta}")
    print("Siguiente paso: firma el proyecto VBA (Alt+F11 > Tools > Digital Signature).")
