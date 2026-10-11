# Luma para Android

Luma adapta el conversor de este repositorio a una app en español que funciona sin conexión. El script de escritorio se conserva y se incluye durante la compilación como `office_core.py`; la adaptación estricta está en `android/app/src/main/python/luma_engine.py`.

## Lo que incluye

- Modo claro, oscuro o del sistema, con elección persistente y barras Android del mismo tema.
- Paneles translúcidos, transiciones cortas y respeto a la preferencia de movimiento reducido.
- Selección de varios archivos o de una carpeta con recorrido de subcarpetas.
- Un CSV por hoja; separador coma, punto y coma o tabulación; UTF-8 con/sin BOM o Windows-1252.
- Protección opcional de textos que parecen fórmulas, activada inicialmente.
- PDF A4 de texto con fuentes Android, Unicode y paginación.
- Historial privado, apertura, guardado individual, ZIP del lote, reporte CSV y compartir mediante Android.
- Archivos de ejemplo para probar antes de seleccionar documentos propios.
- Errores por archivo, cancelación al finalizar el archivo actual y recuperación de resultados terminados si el proceso se interrumpe.

Los originales no se modifican. Los resultados permanecen en el almacenamiento privado hasta eliminarlos o desinstalar la app. Exporta lo que quieras conservar.

## Compatibilidad y límites

Android 7.0 o posterior (API 24); APK universal con ARM de 32/64 bits y x86_64. Luma no necesita permisos de Internet ni acceso general al almacenamiento. Usa el selector de documentos de Android.

Hojas: XLSX, XLSM, XLTX, XLTM, XLS, XLT, XLSB, ODS, OTS, FODS y XML de hoja de cálculo. Documentos: DOCX, DOCM, DOTX, DOTM, ODT, OTT, FODT, SXW, RTF, TXT, TEXT, HTML, HTM y XHTML.

Los PDF de documentos Office/HTML/RTF contienen texto y tablas aproximadas: no conservan imágenes, estilos ni maquetación original. Esto se anuncia antes de convertir y en cada resultado afectado. Los documentos sin texto extraíble o dañados fallan explícitamente; no se genera un PDF de error marcado como éxito. DOC antiguo, Pages, enlaces de Google y documentos cifrados requieren exportación previa desde su aplicación de origen.

CSV conserva los valores guardados de las fórmulas; si no hay valor cacheado, conserva el texto y muestra un aviso. No recalcula fórmulas. Hasta 100 archivos, 32 MB por archivo, 128 MB por lote, 500.000 celdas por hoja, un millón de caracteres y 300 páginas de texto por documento. Archivos ZIP internos: 100 MB expandidos y 5.000 entradas. Las carpetas pueden contener hasta 3.000 elementos y 20 niveles. Android puede restringir algunas carpetas del sistema.

## Compilación y pruebas

El workflow `.github/workflows/luma-apk.yml` ejecuta pruebas del motor, pruebas de interfaz con Chromium, compilación, Android Lint y pruebas instrumentadas en Android 15. Genera un APK de producción firmado y su SHA-256 en el artefacto **Luma-APK**. La firma persistente se conserva fuera del repositorio y en secretos de GitHub para permitir futuras actualizaciones.

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements-android.txt pytest playwright
.venv/bin/python -m pytest -q
.venv/bin/python -m playwright install chromium
.venv/bin/python scripts/browser_smoke.py

# JDK 17, Python 3.11 y Android SDK 35:
cd android
./gradlew assembleRelease lintRelease
```

`assembleRelease` produce inicialmente un APK sin firmar. GitHub Actions lo firma con `LUMA_SIGNING_KEYSTORE` y `LUMA_SIGNING_PASSWORD`. No subas las claves al código. `scripts/setup_signing.py` configura estos secretos para el repositorio del propietario y guarda la clave en `/root/.local/share/luma/signing/`.

Arquitectura: interfaz HTML/CSS local servida mediante [WebViewAssetLoader](https://developer.android.com/reference/androidx/webkit/WebViewAssetLoader), puente limitado a recursos internos, tareas secuenciales en un hilo de trabajo y Python integrado con [Chaquopy](https://chaquo.com/chaquopy/doc/current/android.html). El PDF se escribe con [PdfDocument de Android](https://developer.android.com/reference/android/graphics/pdf/PdfDocument). La selección y exportación usan el [Storage Access Framework](https://developer.android.com/training/data-storage/shared/documents-files).

## Instalar

Descarga `Luma.apk`, ábrelo desde el administrador de archivos y permite la instalación desde esa aplicación si Android lo solicita. Después, toca **Añadir archivos** o **Prueba con un ejemplo**. Los resultados aparecen al terminar y quedan disponibles en **Actividad**.
