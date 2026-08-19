from __future__ import annotations

import argparse
import html
import re
from datetime import date, datetime
from html.parser import HTMLParser
from pathlib import Path
from typing import Iterable, List, Sequence, Tuple
from zipfile import ZIP_DEFLATED, ZipFile


VARIABLE_ESCENARIO = "tasmax_ssp585"
OUTPUT_DEFAULT = "00_IglesiaColorada_tasmax_ssp585_compilado.xlsx"

MESES = [
    (1, "ene", "enero"),
    (2, "feb", "febrero"),
    (3, "mar", "marzo"),
    (4, "abr", "abril"),
    (5, "may", "mayo"),
    (6, "jun", "junio"),
    (7, "jul", "julio"),
    (8, "ago", "agosto"),
    (9, "sep", "septiembre"),
    (10, "oct", "octubre"),
    (11, "nov", "noviembre"),
    (12, "dic", "diciembre"),
]


class HTMLTableParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self._in_cell = False
        self._current_cell: List[str] = []
        self._current_row: List[str] = []
        self.rows: List[List[str]] = []

    def handle_starttag(self, tag: str, attrs) -> None:
        tag = tag.lower()
        if tag == "tr":
            self._current_row = []
        elif tag in {"td", "th"}:
            self._in_cell = True
            self._current_cell = []

    def handle_data(self, data: str) -> None:
        if self._in_cell:
            self._current_cell.append(data)

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        if tag in {"td", "th"} and self._in_cell:
            value = " ".join("".join(self._current_cell).split())
            self._current_row.append(value)
            self._in_cell = False
        elif tag == "tr" and self._current_row:
            self.rows.append(self._current_row)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Genera un xlsx compilado con los archivos mensuales "
            f"*_{VARIABLE_ESCENARIO}.xls."
        )
    )
    parser.add_argument(
        "--input-dir",
        default=".",
        help="Carpeta donde estan los xls mensuales. Por defecto: carpeta actual.",
    )
    parser.add_argument(
        "--output",
        default=OUTPUT_DEFAULT,
        help=f"Nombre del xlsx de salida. Por defecto: {OUTPUT_DEFAULT}",
    )
    return parser.parse_args()


def read_month_file(path: Path) -> List[List[str]]:
    text = path.read_text(encoding="utf-8-sig")
    parser = HTMLTableParser()
    parser.feed(text)
    if len(parser.rows) < 2:
        raise ValueError(f"No se encontro una tabla de datos en {path.name}")
    return [normalize_row(row) for row in parser.rows]


def normalize_row(row: Sequence[str]) -> List[str]:
    normalized = []
    for value in row:
        normalized.append(value.strip().replace(",", "."))
    return normalized


def find_month_file(input_dir: Path, month_num: int, month_short: str) -> Path:
    pattern = f"{month_num:02d}_*_{month_short}_{VARIABLE_ESCENARIO}.xls"
    matches = sorted(input_dir.glob(pattern))
    if not matches:
        raise FileNotFoundError(f"No se encontro archivo con patron: {pattern}")
    if len(matches) > 1:
        names = ", ".join(path.name for path in matches)
        raise ValueError(f"Hay mas de un archivo para {month_num:02d} {month_short}: {names}")
    return matches[0]


def build_compiled_rows(month_tables: Sequence[Tuple[int, str, List[List[str]]]]) -> List[List[str]]:
    header = month_tables[0][2][0]
    compiled = [["Fecha", "Año", "Mes", "Mes_nombre"] + header[1:]]

    for month_num, month_name, rows in month_tables:
        current_header = rows[0]
        if current_header != header:
            raise ValueError(f"Las columnas del mes {month_num} no coinciden con enero.")

        for row in rows[1:]:
            year = row[0]
            date_value = f"{year}-{month_num:02d}-01"
            compiled.append([date_value, year, str(month_num), month_name] + row[1:])

    return compiled


def column_name(index: int) -> str:
    name = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name


def cell_ref(row_idx: int, col_idx: int) -> str:
    return f"{column_name(col_idx)}{row_idx}"


def looks_numeric(value: str) -> bool:
    return bool(re.fullmatch(r"-?\d+(\.\d+)?", value))


def excel_date_serial(value: str) -> int:
    parsed = datetime.strptime(value, "%Y-%m-%d").date()
    return (parsed - date(1899, 12, 30)).days


def looks_iso_date(value: str) -> bool:
    return bool(re.fullmatch(r"\d{4}-\d{2}-\d{2}", value))


def cell_xml(row_idx: int, col_idx: int, value: str, header: bool) -> str:
    ref = cell_ref(row_idx, col_idx)
    style = ' s="1"' if header else ""

    if not header and looks_iso_date(value):
        return f'<c r="{ref}" s="2"><v>{excel_date_serial(value)}</v></c>'

    if not header and looks_numeric(value):
        return f'<c r="{ref}"><v>{value}</v></c>'

    safe_value = html.escape(str(value), quote=False)
    return f'<c r="{ref}" t="inlineStr"{style}><is><t>{safe_value}</t></is></c>'


def sheet_xml(rows: Sequence[Sequence[str]]) -> str:
    max_cols = max(len(row) for row in rows)
    max_rows = len(rows)
    dimension = f"A1:{column_name(max_cols)}{max_rows}"

    col_defs = []
    for col_idx in range(1, max_cols + 1):
        width = 12 if col_idx <= 4 else 24
        col_defs.append(f'<col min="{col_idx}" max="{col_idx}" width="{width}" customWidth="1"/>')

    row_xml = []
    for row_idx, row in enumerate(rows, start=1):
        cells = [
            cell_xml(row_idx, col_idx, value, header=(row_idx == 1))
            for col_idx, value in enumerate(row, start=1)
        ]
        row_xml.append(f'<row r="{row_idx}">{"".join(cells)}</row>')

    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="{dimension}"/>
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
      <selection pane="bottomLeft"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols>{"".join(col_defs)}</cols>
  <sheetData>{"".join(row_xml)}</sheetData>
  <autoFilter ref="{dimension}"/>
</worksheet>'''


def content_types_xml(sheet_count: int) -> str:
    sheets = "".join(
        f'<Override PartName="/xl/worksheets/sheet{i}.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        for i in range(1, sheet_count + 1)
    )
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  {sheets}
</Types>'''


def root_rels_xml() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''


def workbook_xml(sheet_names: Sequence[str]) -> str:
    sheets = []
    for idx, name in enumerate(sheet_names, start=1):
        safe_name = html.escape(name, quote=True)
        sheets.append(f'<sheet name="{safe_name}" sheetId="{idx}" r:id="rId{idx}"/>')

    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>{"".join(sheets)}</sheets>
</workbook>'''


def workbook_rels_xml(sheet_count: int) -> str:
    relationships = []
    for idx in range(1, sheet_count + 1):
        relationships.append(
            f'<Relationship Id="rId{idx}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{idx}.xml"/>'
        )
    relationships.append(
        f'<Relationship Id="rId{sheet_count + 1}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/>'
    )
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  {"".join(relationships)}
</Relationships>'''


def styles_xml() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts>
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="3">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''


def core_xml() -> str:
    now = datetime.utcnow().replace(microsecond=0).isoformat() + "Z"
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>generar_compilatorio_tasmax_ssp245.py</dc:creator>
  <cp:lastModifiedBy>generar_compilatorio_tasmax_ssp245.py</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>
</cp:coreProperties>'''


def app_xml(sheet_names: Sequence[str]) -> str:
    sheet_count = len(sheet_names)
    titles = "".join(f"<vt:lpstr>{html.escape(name)}</vt:lpstr>" for name in sheet_names)
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Python</Application>
  <TitlesOfParts><vt:vector size="{sheet_count}" baseType="lpstr">{titles}</vt:vector></TitlesOfParts>
</Properties>'''


def write_xlsx(output_path: Path, sheets: Sequence[Tuple[str, Sequence[Sequence[str]]]]) -> None:
    sheet_names = [name for name, _ in sheets]
    with ZipFile(output_path, "w", ZIP_DEFLATED) as xlsx:
        xlsx.writestr("[Content_Types].xml", content_types_xml(len(sheets)))
        xlsx.writestr("_rels/.rels", root_rels_xml())
        xlsx.writestr("docProps/core.xml", core_xml())
        xlsx.writestr("docProps/app.xml", app_xml(sheet_names))
        xlsx.writestr("xl/workbook.xml", workbook_xml(sheet_names))
        xlsx.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml(len(sheets)))
        xlsx.writestr("xl/styles.xml", styles_xml())
        for idx, (_name, rows) in enumerate(sheets, start=1):
            xlsx.writestr(f"xl/worksheets/sheet{idx}.xml", sheet_xml(rows))


def main() -> None:
    args = parse_args()
    input_dir = Path(args.input_dir).resolve()
    output_path = Path(args.output)
    if not output_path.is_absolute():
        output_path = input_dir / output_path

    month_tables = []
    sheets = []
    for month_num, month_short, month_name in MESES:
        source = find_month_file(input_dir, month_num, month_short)
        rows = read_month_file(source)
        sheet_name = str(month_num)
        sheets.append((sheet_name, rows))
        month_tables.append((month_num, month_name, rows))

    sheets.append(("Compilado", build_compiled_rows(month_tables)))
    write_xlsx(output_path, sheets)

    print(f"Compilatorio generado: {output_path}")
    print(f"Hojas mensuales: {len(MESES)}")
    print(f"Hoja final: Compilado")


if __name__ == "__main__":
    main()
