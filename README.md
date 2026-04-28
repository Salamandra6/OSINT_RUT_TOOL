# OSINT RUT TOOL Chile

Herramienta local de ciberinteligencia/OSINT en Python para detectar RUT chilenos en archivos PDF digitales o escaneados.

La aplicacion funciona de manera local, no usa servidor web y exporta los resultados a Excel `.xlsx`.

## Funciones principales

- Carga de archivos PDF.
- Lectura digital con `pdfplumber`.
- OCR automatico con `pytesseract` y `pdf2image` cuando el PDF no contiene texto extraible.
- Deteccion de RUT chilenos en distintos formatos:
  - `16567792-7`
  - `16.567.792-7`
  - `165677927`
  - `16567792`
- Validacion de digito verificador.
- Identificacion aproximada de nombres cercanos al RUT.
- Visualizacion de resultados en tabla.
- Exportacion automatica a Excel.
- Uso portable con rutas relativas.

## Requisitos

- Windows 10/11
- Python 3.10 o superior
- Tesseract OCR
- Poppler para Windows

## Instalacion rapida en PowerShell

```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt
python .\osint_rut_gui.py
```

## Estructura portable recomendada

```text
OSINT_RUT_TOOL/
│
├── OSINT_RUT_GUI.exe
├── osint_rut_gui.py
├── requirements.txt
├── resultados/
│
├── tesseract/
│   └── tesseract.exe
│
└── poppler/
    └── Library/
        └── bin/
            ├── pdftoppm.exe
            ├── pdfinfo.exe
            └── otros archivos de Poppler
```

La carpeta `resultados` se crea automaticamente si no existe.

## Generar ejecutable EXE

Desde PowerShell, dentro del proyecto y con el entorno virtual activado:

```powershell
pyinstaller --onefile --windowed --name OSINT_RUT_GUI osint_rut_gui.py
```

El ejecutable quedara en:

```text
dist/OSINT_RUT_GUI.exe
```

Para distribuirlo de forma portable, copiar el `.exe` junto con las carpetas `tesseract` y `poppler`.

## Notas tecnicas

- Si el PDF tiene texto digital, se procesa con `pdfplumber`.
- Si no se detecta texto suficiente, se aplica OCR automaticamente.
- El RUT se valida mediante calculo de digito verificador.
- Los resultados se exportan a la carpeta `resultados`.
- No se recomienda subir a GitHub documentos reales que contengan datos personales.

## Uso responsable

Esta herramienta esta pensada para laboratorios de ciberseguridad, ciberinteligencia, auditoria documental y analisis OSINT autorizado.

Debe utilizarse solo sobre documentos propios, publicos o respecto de los cuales exista autorizacion legal o institucional.
