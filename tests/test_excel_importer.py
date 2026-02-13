from pathlib import Path

import pytest

from core.excel_importer import ExcelImportError, import_excel


def test_import_excel_rejects_legacy_xls(tmp_path: Path):
    legacy_file = tmp_path / "startlist.xls"
    legacy_file.write_text("legacy excel")

    with pytest.raises(ExcelImportError, match=".xls"):
        import_excel(legacy_file)
