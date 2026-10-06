import os
import sys

try:
    import win32com.client as win32
except ImportError:
    print("Error: Se requiere la librería pywin32.")
    print("Ejecuta: pip install pywin32")
    input("Presiona ENTER para salir...")
    sys.exit(1)

# Constantes de Excel (XlPaperSize y XlFixedFormatType)
XL_PAPER_LETTER = 1       # Tamaño Carta (8.5 x 11 in)
XL_PAPER_A4 = 9           # Tamaño A4 (210 x 297 mm)
XL_TYPE_PDF = 0           # Formato de exportación PDF
XL_SHEET_VISIBLE = -1     # Hoja visible

# Extensiones de Excel estrictamente permitidas
EXTENSIONES_EXCEL_VALIDAS = ('.xlsx', '.xls', '.xlsm')

def mostrar_banner():
    banner = r"""
        .---.
       / /"\ \          ==================================================
      | | ☢ | |          PND Only BMPI Formatter - AUTO EXCEL TO PDF
       \ \_/ /          ==================================================
     .==`---'==.
    / /       \ \        ☢  DESARROLLADO POR: JAMR  ☢
   / /         \ \       🔗 GitHub: https://github.com/JAMRz
   \ \         / /       📋 Versión: 1.1 (Ajuste Carta / A4 + 1x1)
    \ `-------' /       ==================================================
     `---------'
    """
    print(banner)

def configurar_pagina(ws):
    """Ajusta la hoja a 1x1 página y tamaño Carta (con fallback a A4)."""
    ps = ws.PageSetup
    ps.Zoom = False
    ps.FitToPagesWide = 1
    ps.FitToPagesTall = 1

    # Intentar asignar Tamaño Carta; si falla el driver de impresión, usar A4
    try:
        ps.PaperSize = XL_PAPER_LETTER
    except Exception:
        try:
            ps.PaperSize = XL_PAPER_A4
        except Exception as err_papel:
            print(f" [Aviso papel: {err_papel}]", end="")

def procesar_reportes():
    mostrar_banner()

    print("Ingresa la ruta de la carpeta que contiene los reportes de Excel.")
    print("(Deja en blanco y presiona ENTER para procesar la carpeta actual)")
    ruta_input = input("Ruta: ").strip().strip('"').strip("'")

    if not ruta_input:
        if getattr(sys, 'frozen', False):
            directorio = os.path.dirname(sys.executable)
        else:
            directorio = os.path.dirname(os.path.abspath(__file__))
    else:
        directorio = os.path.abspath(ruta_input)

    if not os.path.isdir(directorio):
        print(f"\n❌ Error: La ruta '{directorio}' no existe.")
        input("\nPresiona ENTER para salir...")
        return

    # Filtro estricto: Solo archivos reales de Excel, omitiendo carpetas y archivos temporales (~$)
    archivos = [
        f for f in os.listdir(directorio)
        if os.path.isfile(os.path.join(directorio, f))
        and f.lower().endswith(EXTENSIONES_EXCEL_VALIDAS)
        and not f.startswith('~$')
    ]

    total = len(archivos)
    if total == 0:
        print(f"\n⚠️ No se encontraron archivos válidos de Excel ({', '.join(EXTENSIONES_EXCEL_VALIDAS)}) en: {directorio}")
        input("\nPresiona ENTER para salir...")
        return

    print(f"\n📁 Carpeta de trabajo: {directorio}")
    print(f"📊 Total de reportes encontrados: {total}\n")

    # Iniciar Excel en segundo plano
    try:
        excel = win32.gencache.EnsureDispatch('Excel.Application')
    except Exception:
        excel = win32.Dispatch('Excel.Application')

    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False

    exitosos = 0
    fallidos = 0

    try:
        for idx, archivo in enumerate(archivos, start=1):
            ruta_excel = os.path.join(directorio, archivo)
            nombre_base = os.path.splitext(archivo)[0]
            ruta_pdf = os.path.join(directorio, f"{nombre_base}.pdf")

            print(f"[{idx}/{total}] Procesando reporte BMPI: {archivo} ...", end="", flush=True)

            wb = None
            try:
                wb = excel.Workbooks.Open(ruta_excel, ReadOnly=True)

                # Configurar cada hoja visible
                for ws in wb.Worksheets:
                    if ws.Visible == XL_SHEET_VISIBLE:
                        configurar_pagina(ws)

                # Exportar a PDF
                wb.ExportAsFixedFormat(XL_TYPE_PDF, ruta_pdf)
                wb.Close(SaveChanges=False)
                wb = None

                print(" -> ✅ PDF OK (Carta/A4 - 1x1)")
                exitosos += 1
            except Exception as e:
                print(f" -> ❌ Error: {e}")
                fallidos += 1
                if wb:
                    try:
                        wb.Close(SaveChanges=False)
                    except Exception:
                        pass
    finally:
        excel.Quit()

    print("\n" + "=" * 54)
    print("☢  PND Only BMPI Formatter - FINALIZADO")
    print(f"   Autor:              JAMR (github.com/JAMRz)")
    print(f"   Reportes listos:    {exitosos}")
    print(f"   Fallidos:           {fallidos}")
    print("=" * 54)
    input("\nPresiona ENTER para salir...")

if __name__ == "__main__":
    procesar_reportes()