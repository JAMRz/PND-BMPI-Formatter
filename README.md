# PND Only BMPI Formatter - Auto Excel to PDF

Herramienta ligera para procesar por lotes hojas de cálculo de Microsoft Excel y exportarlas automáticamente a formato PDF, ajustando cada hoja visible a exactamente una página de alto por una de ancho (1x1) y estableciendo el tamaño de papel en Carta (con respaldo automático a A4).

---

## 🚀 Características

- **Procesamiento por lotes:** Convierte múltiples archivos `.xlsx`, `.xls` y `.xlsm` de una carpeta en una sola ejecución.
- **Ajuste 1x1 automático:** Fuerza el escalado vertical y horizontal a una página por hoja visible.
- **Formato Carta / A4:** Establece tamaño Carta de forma prioritaria con fallback transparente a A4.
- **Seguro y no destructivo:** Abre los archivos en modo de solo lectura (`ReadOnly=True`) y no aplica cambios sobre los archivos originales de Excel.
- **Filtrado estricto:** Ignora archivos temporales de bloqueo de Office (`~$*`) y directorios.

---

## 📥 Descarga directa

Puedes descargar el binario ejecutable sin necesidad de tener Python instalado:

👉 **[Descargar la última versión (.exe)](https://github.com/JAMRz/PND-BMPI-Formatter/releases/latest)**

---

## 🛠️ Requisitos para ejecutar desde código fuente

1. Windows con **Microsoft Excel** instalado.
2. Python 3.8 o superior.
3. Biblioteca `pywin32`:

```bash
pip install pywin32
```

Ejecución directa:

```bash
python PND_BMPI_Formatter_JAMR.py
```

---

## 📦 Compilación a `.exe`

Primero instala PyInstaller:

```bash
pip install pyinstaller
```

### Opción 1: Compilación estándar (Recomendada, sin metadatos)

Genera el ejecutable limpio directamente desde el script:

```bash
pyinstaller --noconfirm --onefile --console --collect-submodules win32com "PND_BMPI_Formatter_JAMR.py"
```

### Opción 2: Compilación con metadatos de versión (Opcional)

Si deseas incrustar la información de autor, versión y detalles de Windows usando `version.txt`:

```bash
pyinstaller --noconfirm --onefile --console --version-file "version.txt" --collect-submodules win32com "PND_BMPI_Formatter_JAMR.py"
```

El archivo final `PND_BMPI_Formatter_JAMR.exe` se generará en la carpeta `dist/`.

---

## 📄 Licencia

Herramienta desarrollada por **JAMR** ([GitHub](https://github.com/JAMRz)) de libre uso y distribución.