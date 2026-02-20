from pathlib import Path

import pytest
from openpyxl import Workbook

from core.excel_importer import ExcelImportError, import_excel


def test_import_excel_rejects_legacy_xls(tmp_path: Path):
    legacy_file = tmp_path / "startlist.xls"
    legacy_file.write_text("legacy excel")

    with pytest.raises(ExcelImportError, match=".xls"):
        import_excel(legacy_file)


def test_import_excel_parses_birth_year_as_text(tmp_path: Path):
    file_path = tmp_path / "startlist.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "50m"
    ws.append(["ФИО", "Год рождения", "Заплыв/дорожка"])
    ws.append(["Иванов Иван", "2012 г.", "6/3"])
    wb.save(file_path)

    data = import_excel(file_path)
    swimmer = data["50m"][0]
    assert swimmer["birth_year"] == 2012
