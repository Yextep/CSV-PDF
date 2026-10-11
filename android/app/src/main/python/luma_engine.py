"""Adaptación local y estricta del conversor de escritorio para Luma."""
from __future__ import annotations

import importlib.util
import json
import sys
import zipfile
from pathlib import Path
from defusedxml import ElementTree as ET

try:
    import office_core as core
except ImportError:
    # Las pruebas de escritorio usan exactamente el script del repositorio.
    root = next(p for p in Path(__file__).resolve().parents if (p / "csv-pdf.py").is_file())
    spec = importlib.util.spec_from_file_location("office_core", root / "csv-pdf.py")
    core = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = core
    spec.loader.exec_module(core)

core.ET = ET

SHEETS = {".xlsx", ".xlsm", ".xltx", ".xltm", ".xls", ".xlt", ".xlsb", ".ods", ".ots", ".fods", ".xml"}
DOCS = {".docx", ".docm", ".dotx", ".dotm", ".odt", ".ott", ".fodt", ".sxw", ".txt", ".text", ".html", ".htm", ".xhtml", ".rtf"}
MAX_CELLS = 500_000
MAX_TEXT = 1_000_000


def inspect_archive(source: Path) -> None:
    if not zipfile.is_zipfile(source):
        return
    with zipfile.ZipFile(source) as archive:
        entries = archive.infolist()
        if len(entries) > 5000 or sum(x.file_size for x in entries) > 100 * 1024 * 1024:
            raise ValueError("El archivo expandido es demasiado grande para el teléfono.")
        if any(x.flag_bits & 1 for x in entries):
            raise ValueError("El archivo está cifrado. Guarda una copia sin contraseña.")


def local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def attr(node, key, default=""):
    return next((v for k, v in node.attrib.items() if local(k) == key), default)


def odf_root(source):
    if source.suffix.lower() in {".fods", ".fodt"}:
        return ET.fromstring(source.read_bytes())
    with zipfile.ZipFile(source) as archive:
        return ET.fromstring(archive.read("content.xml"))


def odf_text(node):
    parts = []
    def visit(element):
        if element.text:
            parts.append(element.text)
        for child in element:
            name = local(child.tag)
            if name == "s":
                parts.append(" " * min(int(attr(child, "c", "1")), 1000))
            elif name == "tab":
                parts.append("\t")
            elif name == "line-break":
                parts.append("\n")
            else:
                visit(child)
            if child.tail:
                parts.append(child.tail)
    visit(node)
    return "".join(parts)


def odf_sheets(source):
    root = odf_root(source)
    for table in root.iter():
        if local(table.tag) != "table":
            continue
        rows, missing_formula, total, pending = [], 0, 0, 0
        for row in table.iter():
            if local(row.tag) != "table-row":
                continue
            values, blanks = [], 0
            for cell in row:
                if local(cell.tag) not in {"table-cell", "covered-table-cell"}:
                    continue
                repeat = max(1, int(attr(cell, "number-columns-repeated", "1")))
                kind = attr(cell, "value-type")
                if kind in {"float", "percentage", "currency"}:
                    value = attr(cell, "value")
                elif kind == "boolean":
                    value = attr(cell, "boolean-value").upper()
                elif kind in {"date", "time"}:
                    value = attr(cell, kind + "-value")
                else:
                    value = "\n".join(odf_text(p) for p in cell if local(p.tag) == "p")
                if not value and attr(cell, "formula"):
                    value = attr(cell, "formula")
                    missing_formula += repeat
                if not value:
                    blanks += repeat
                    continue
                if len(values) + blanks + repeat > 16384:
                    raise ValueError("La hoja tiene demasiadas columnas para convertirla aquí.")
                values.extend([""] * blanks)
                values.extend([value] * repeat)
                blanks = 0
            repeat_rows = max(1, int(attr(row, "number-rows-repeated", "1")))
            if not values:
                pending += repeat_rows
                continue
            total += pending + len(values) * repeat_rows
            if total > MAX_CELLS or len(rows) + pending + repeat_rows > 100000:
                raise ValueError("La hoja supera el límite de 500.000 celdas.")
            rows.extend([[] for _ in range(pending)])
            rows.extend([values[:] for _ in range(repeat_rows)])
            pending = 0
        warnings = []
        if missing_formula:
            warnings.append(f"{missing_formula} fórmulas sin valor guardado: se conserva su texto.")
        yield core.SheetData(source, attr(table, "name", "Hoja"), rows, "odf-local", warnings)


def check_xlsx_dimensions(source):
    from openpyxl import load_workbook
    book = load_workbook(source, read_only=True, data_only=False)
    try:
        for sheet in book:
            if (sheet.max_row or 0) * (sheet.max_column or 0) > MAX_CELLS:
                raise ValueError("La hoja supera el límite de 500.000 celdas. Divide el libro y vuelve a intentarlo.")
    finally:
        book.close()


def read_sheets(source):
    ext = source.suffix.lower()
    if ext in {".ods", ".ots", ".fods"}:
        return odf_sheets(source)
    if ext in core.OOXML_EXCEL_EXTENSIONS and zipfile.is_zipfile(source):
        check_xlsx_dimensions(source)
    if ext == ".xml":
        # XML se acepta únicamente como hoja: nunca como texto por accidente.
        return core.iter_xml_spreadsheet_sheets(source)
    return core.iter_excel_sheets_with_fallbacks(source, None, 60)


def document_lines(source):
    ext = source.suffix.lower()
    if ext in {".docx", ".docm", ".dotx", ".dotm"}:
        lines, _ = core.extract_docx_lines(source)
    elif ext in {".odt", ".ott", ".fodt", ".sxw"}:
        lines = [odf_text(node) for node in odf_root(source).iter() if local(node.tag) in {"p", "h"}]
    elif ext == ".rtf":
        from striprtf.striprtf import rtf_to_text
        text, _ = core.read_text_lossy(source)
        if not text.lstrip().startswith("{\\rtf"):
            raise ValueError("El archivo no es un documento RTF válido.")
        lines = rtf_to_text(text).splitlines()
    elif ext in {".html", ".htm", ".xhtml"}:
        lines, _ = core.extract_html_lines(source)
    else:
        if not core.is_probably_text_bytes(core.read_prefix(source)):
            raise ValueError("El archivo no contiene texto legible.")
        text, _ = core.read_text_lossy(source)
        lines = text.lstrip("\ufeff").splitlines()
    if not any(line.strip() for line in lines):
        raise ValueError("No hay texto extraíble. Los documentos que solo contienen imágenes necesitan otro motor.")
    if sum(len(line) for line in lines) > MAX_TEXT:
        raise ValueError("El documento supera un millón de caracteres. Divide el archivo.")
    warnings = []
    if ext not in {".txt", ".text"}:
        warnings.append("PDF de texto: diseño aproximado, sin imágenes ni formato original. Revisa el resultado.")
    return lines, warnings


def convert(source_path: str, output_path: str, options_json: str = "{}") -> str:
    """No genera PDF de emergencia ni declara éxito cuando no hay contenido."""
    source, output = Path(source_path), Path(output_path)
    options = json.loads(options_json)
    delimiter = options.get("delimiter", ",")
    encoding = options.get("encoding", "utf-8-sig")
    if delimiter not in {",", ";", "\t"} or encoding not in {"utf-8-sig", "utf-8", "cp1252"}:
        raise ValueError("La configuración CSV no es válida.")
    ext = source.suffix.lower()
    if ext not in SHEETS | DOCS:
        raise ValueError("Formato no compatible en Android. Exporta a XLSX, ODS, DOCX, ODT o TXT.")
    if not source.is_file() or source.stat().st_size == 0:
        raise ValueError("El archivo está vacío o ya no está disponible.")
    if source.stat().st_size > 32 * 1024 * 1024:
        raise ValueError("El archivo supera el límite de 32 MB.")
    inspect_archive(source)
    output.mkdir(parents=True, exist_ok=True)
    if ext in DOCS:
        lines, warnings = document_lines(source)
        return json.dumps({"kind": "pdf", "lines": lines, "warnings": warnings}, ensure_ascii=False)
    files, warnings = [], []
    try:
        sheets = list(read_sheets(source))
        if not sheets:
            raise ValueError("El libro no contiene hojas convertibles.")
        for sheet in sheets:
            name = core.build_excel_output_name(source, sheet.sheet_name, len(sheets))
            target = core.unique_path(output / name)
            temp = target.with_suffix(".part")
            files.append(target)
            def bounded_rows():
                count = 0
                for row in sheet.rows:
                    count += max(1, len(row))
                    if count > MAX_CELLS:
                        raise ValueError("La hoja supera el límite de 500.000 celdas.")
                    yield row
            core.write_rows_to_csv(bounded_rows(), temp, delimiter, encoding, bool(options.get("safe", True)))
            temp.replace(target)
            warnings.extend(sheet.warnings)
        return json.dumps({"kind": "csv", "files": [p.name for p in files], "warnings": list(dict.fromkeys(warnings))}, ensure_ascii=False)
    except Exception:
        for file in files:
            file.unlink(missing_ok=True)
            file.with_suffix(".part").unlink(missing_ok=True)
        raise


def create_demo(directory: str) -> None:
    from openpyxl import Workbook
    target = Path(directory)
    target.mkdir(parents=True, exist_ok=True)
    (target / "Bienvenida.txt").write_text(
        "Hola, soy Luma.\n\nUn poquito de luz para tus archivos.\n"
        "Convierte documentos en PDF y hojas de cálculo en CSV, sin conexión.\n\n"
        "Puedes guardar tus resultados o compartirlos desde la app.\n"
        "Acentos que viajan contigo: á, é, í, ó, ú, ñ.\n", encoding="utf-8")
    book = Workbook()
    sheet = book.active
    sheet.title = "Inventario"
    sheet.append(["Producto", "Cantidad", "Precio"])
    sheet.append(["Café", 4, 12500])
    sheet.append(["Té verde", 2, 8900])
    sheet.append(["Cuaderno", 3, 6000])
    book.save(target / "Pequeño inventario.xlsx")


def smoke(directory: str) -> str:
    target = Path(directory)
    create_demo(directory)
    csv_result = json.loads(convert(str(target / "Pequeño inventario.xlsx"), str(target / "csv")))
    assert "Café" in (target / "csv" / csv_result["files"][0]).read_text(encoding="utf-8-sig")
    pdf_result = json.loads(convert(str(target / "Bienvenida.txt"), str(target / "pdf")))
    assert any("á, é" in line for line in pdf_result["lines"])
    (target / "dañado.docx").write_bytes(b"no es un docx")
    try:
        convert(str(target / "dañado.docx"), str(target / "broken"))
    except Exception:
        pass
    else:
        raise AssertionError("Un DOCX dañado no puede declararse convertido")
    return json.dumps(pdf_result, ensure_ascii=False)
