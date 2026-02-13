from __future__ import annotations

"""Generate sample startlist Excel file without external dependencies.

Usage:
    python data/create_sample_startlist.py
"""

from pathlib import Path
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZipFile

OUTPUT = Path(__file__).resolve().parent / "competition-15-startlist.xlsx"


def _col_name(index: int) -> str:
    result = ""
    while index:
        index, rem = divmod(index - 1, 26)
        result = chr(65 + rem) + result
    return result


def _sheet_xml(rows: list[list[object]]) -> str:
    lines = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>',
    ]

    for r_idx, row in enumerate(rows, start=1):
        lines.append(f'<row r="{r_idx}">')
        for c_idx, value in enumerate(row, start=1):
            ref = f"{_col_name(c_idx)}{r_idx}"
            if isinstance(value, (int, float)):
                lines.append(f'<c r="{ref}"><v>{value}</v></c>')
            else:
                text = escape(str(value))
                lines.append(f'<c r="{ref}" t="inlineStr"><is><t>{text}</t></is></c>')
        lines.append("</row>")

    lines.append("</sheetData></worksheet>")
    return "".join(lines)


def generate(path: Path = OUTPUT) -> Path:
    rows1 = [
        ["Ф. И.", "Год рождения", "Команда", "Заявочное время", "Заплыв/дорожка"],
        ["Тестов1 Участник1", 2011, "Load Team 1", "0.58.88", "1/3"],
        ["Тестов2 Участник2", 2014, "Load Team 1", "0.59.21", "1/4"],
        ["Тестов3 Участник3", 2012, "Load Team 2", "1.00.43", "1/2"],
        ["Тестов4 Участник4", 2011, "Load Team 3", "1.00.71", "1/5"],
    ]
    rows2 = [
        ["Ф. И.", "Год рождения", "Команда", "Заявочное время", "Заплыв/дорожка"],
        ["Тестов11 Участник11", 2010, "Load Team 3", "1.10.22", "1/1"],
        ["Тестов12 Участник12", 2011, "Load Team 2", "1.11.45", "1/2"],
        ["Тестов13 Участник13", 2012, "Load Team 2", "1.12.09", "1/3"],
    ]

    content_types = """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">
<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>
<Default Extension=\"xml\" ContentType=\"application/xml\"/>
<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>
<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>
<Override PartName=\"/xl/worksheets/sheet2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>
</Types>"""
    rels = """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>
</Relationships>"""
    workbook = """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">
<sheets>
<sheet name=\"1 100 батт\" sheetId=\"1\" r:id=\"rId1\"/>
<sheet name=\"2 100 брасс\" sheetId=\"2\" r:id=\"rId2\"/>
</sheets>
</workbook>"""
    workbook_rels = """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>
<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet2.xml\"/>
</Relationships>"""

    path.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("_rels/.rels", rels)
        archive.writestr("xl/workbook.xml", workbook)
        archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        archive.writestr("xl/worksheets/sheet1.xml", _sheet_xml(rows1))
        archive.writestr("xl/worksheets/sheet2.xml", _sheet_xml(rows2))

    return path


if __name__ == "__main__":
    out = generate()
    print(f"Created sample file: {out}")
