# CSV-PDF

Conversor recursivo en Python de hojas de cálculo Excel/Calc a CSV y documentos Word/Writer a PDF. El punto de entrada es [`csv-pdf.py`](csv-pdf.py), que recorre una carpeta y sus subcarpetas, procesa los formatos reconocidos y genera reportes de la ejecución.

## Funciones

- Exporta cada hoja de un libro de cálculo a un CSV independiente.
- Convierte documentos a PDF usando LibreOffice, Microsoft Word mediante COM en Windows o una alternativa de extracción de texto en Python.
- Ofrece uso interactivo y argumentos para ejecuciones sin preguntas.
- Guarda resultados junto al original, en una carpeta central o en ambas ubicaciones.
- Permite elegir el delimitador y la codificación CSV, y dispone de una opción para escapar texto que parezca una fórmula.
- Genera un reporte CSV y un resumen de texto; si registra conversiones fallidas, añade un log de errores.
- Cuando hay carpeta central, puede copiar allí los originales cuya conversión haya fallado.

## Formatos y motores

| Entrada | Salida | Lectores o motores contemplados |
| --- | --- | --- |
| `.xlsx`, `.xlsm`, `.xltx`, `.xltm` | Un CSV por hoja | `openpyxl`, lectores de pandas y respaldo con LibreOffice |
| `.xls`, `.xlt` | Un CSV por hoja | `xlrd`, lectores alternativos y LibreOffice |
| `.xlsb` | Un CSV por hoja | `pyxlsb`, lectores alternativos y LibreOffice |
| `.ods`, `.ots`, `.fods` | CSV | Lectores de pandas/ODF y alternativas según el archivo |
| Tablas HTML, texto delimitado o XML de hoja de cálculo bajo extensiones reconocidas | CSV | Parsers incluidos en el script |
| Documentos Word/Writer, RTF, texto y HTML | PDF | LibreOffice, Word COM cuando está disponible, o extracción de texto en Python |

El script también reconoce extensiones adicionales. Reconocer una extensión no garantiza una conversión fiel: depende del contenido y del motor disponible. Por ejemplo, `.xml` es ambiguo y, al elegir ambas conversiones, se intenta primero como hoja de cálculo.

## Requisitos

- Python 3 para ejecutar el script.
- Bibliotecas opcionales para ampliar la lectura de hojas de cálculo: `pandas`, `openpyxl`, `xlrd`, `pyxlsb`, `odfpy` y `python-calamine`.
- LibreOffice, con `soffice` o `libreoffice` disponible para el script, para las rutas de conversión que lo utilizan.
- En Windows, la alternativa Word COM requiere Microsoft Word y `pywin32`.

Puedes preparar las bibliotecas Python en un entorno virtual. El repositorio no incluye un archivo de dependencias:

```bash
python3 -m venv .venv
```

Activa el entorno según tu sistema:

```bash
# Linux/macOS
source .venv/bin/activate
```

```powershell
# Windows PowerShell
.venv\Scripts\Activate.ps1
```

```bash
python -m pip install pandas openpyxl xlrd pyxlsb odfpy python-calamine
```

Para utilizar Word COM en Windows, instala también `pywin32` en el entorno. LibreOffice se instala por separado como aplicación del sistema.

## Uso

Desde la carpeta del repositorio, después de activar el entorno:

```bash
# Consultar las opciones
python csv-pdf.py --help

# Elegir carpeta, conversión y destino de forma interactiva
python csv-pdf.py

# Convertir ambos tipos y centralizar los resultados
python csv-pdf.py ./documentos --convert both --output central --non-interactive

# Exportar solo hojas de cálculo, conservando el árbol de carpetas
python csv-pdf.py ./documentos --convert excel --output central --central-layout tree --non-interactive

# Convertir documentos a PDF junto al archivo original
python csv-pdf.py ./documentos --convert docs --output same --non-interactive
```

Sustituye `./documentos` por una carpeta existente. Con `--non-interactive`, los valores omitidos son la carpeta actual, `--convert both` y `--output both`.

## Opciones principales

| Opción | Comportamiento |
| --- | --- |
| `--convert excel\|docs\|both` | Selecciona los tipos de entrada. |
| `--output same\|central\|both` | Selecciona dónde guardar los resultados. |
| `--central-layout flat\|tree` | Organiza la salida central; `flat` es el valor predeterminado. |
| `--delimiter CARÁCTER` | Cambia el separador CSV; por defecto, coma. |
| `--encoding CODIFICACIÓN` | Cambia la codificación de CSV y reportes; por defecto, `utf-8-sig`. |
| `--safe-csv` | Escapa texto que parece una fórmula al abrir el CSV en Excel. |
| `--overwrite` | Permite sobrescribir salidas; sin esta opción, se generan nombres alternativos. |
| `--prefer-libreoffice-excel` | Intenta normalizar o recalcular el libro con LibreOffice antes de leerlo. |
| `--no-libreoffice` | Desactiva las rutas que utilizan LibreOffice. |
| `--prefer-ms-word` | Da prioridad a Word COM para documentos en Windows. |
| `--timeout SEGUNDOS` | Establece el límite de tiempo para conversiones con LibreOffice; por defecto, `240`. |
| `--pdfa none\|1\|2\|3` | Solicita PDF normal o PDF/A cuando el motor lo permite. |
| `--no-copy-failed` | Desactiva la copia de originales fallidos a la carpeta central. |
| `--non-interactive` | Evita las preguntas interactivas. |

## Archivos de salida

La carpeta central se crea **en el directorio desde el que se ejecuta el comando**, con un nombre `CONVERSIONES_OFFICE_<fecha y hora>`. Puede contener:

- `CSV/`: hojas exportadas.
- `PDF/`: documentos convertidos.
- `REPORTES/`: reporte CSV, resumen TXT y, cuando hay errores, su log.
- `NO_CONVERTIDOS/`: copias de originales fallidos, cuando corresponde.

En el modo `same`, los reportes se guardan en el directorio de ejecución. Las carpetas centrales de conversiones anteriores y determinados archivos temporales se excluyen del recorrido.

El reporte registra el archivo de origen, el método empleado, el estado, las rutas de salida, las advertencias y los errores. El programa devuelve `0` si no registra errores, `1` si registra conversiones fallidas y `2` si la configuración inicial no es válida.

## Limitaciones

- CSV conserva valores tabulares, sin el formato visual de Excel, gráficos o fórmulas activas. Los valores de fórmulas pueden depender de la caché del libro o del recálculo con LibreOffice.
- El PDF generado únicamente con Python aproxima el texto y no conserva necesariamente el diseño, imágenes ni todos los caracteres del documento.
- Un acceso directo `.gdoc` no contiene el documento de Google: la alternativa Python puede producir un PDF informativo. Las entradas `.gsheet` y `.gslides` no están en las listas de descubrimiento de trabajos.
- Archivos cifrados, dañados o formatos que no puede leer el motor disponible no tienen una conversión garantizada. `.pages` puede necesitar exportación previa desde Apple Pages.
- El script puede generar un PDF informativo de emergencia cuando falla la conversión del contenido. Un estado `OK` o la existencia de un PDF no garantizan fidelidad: revisa el método, las advertencias y el resultado.
