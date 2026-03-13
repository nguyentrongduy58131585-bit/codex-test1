from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

OUTPUT_FILE = Path("equipment_management.xlsx")
HEADERS = [
    "Equipment ID",
    "Equipment Name",
    "Location",
    "Status",
    "Maintenance Date",
    "Notes",
]


def col_name(index: int) -> str:
    """Convert a 1-based column index to Excel column letters."""
    letters = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def build_sheet_xml(headers: list[str]) -> str:
    cols_xml = "".join(
        f'<c r="{col_name(i)}1" t="inlineStr"><is><t>{value}</t></is></c>'
        for i, value in enumerate(headers, start=1)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">\n'
        f'  <dimension ref="A1:{col_name(len(headers))}1"/>\n'
        '  <sheetData>\n'
        f'    <row r="1">{cols_xml}</row>\n'
        '  </sheetData>\n'
        '</worksheet>\n'
    )


def create_excel_file(output_file: Path, headers: list[str]) -> None:
    with ZipFile(output_file, "w", ZIP_DEFLATED) as workbook:
        workbook.writestr(
            "[Content_Types].xml",
            '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>
''',
        )
        workbook.writestr(
            "_rels/.rels",
            '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
''',
        )
        workbook.writestr(
            "xl/workbook.xml",
            '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Equipment" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
''',
        )
        workbook.writestr(
            "xl/_rels/workbook.xml.rels",
            '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>
''',
        )
        workbook.writestr("xl/worksheets/sheet1.xml", build_sheet_xml(headers))


def main() -> None:
    create_excel_file(OUTPUT_FILE, HEADERS)
    print(f"Created {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
