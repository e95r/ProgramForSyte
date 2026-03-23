from pathlib import Path

from openpyxl import Workbook

from core.excel_importer import import_excel


def test_import_excel_uses_single_cell_title_before_header_even_on_protocol_sheet(tmp_path: Path):
    file_path = tmp_path / "protocol-titles.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Протокол"
    ws.append(["Открытый турнир по плаванию «АКВАДОН»"])
    ws.append([])
    ws.append(["100 брасс, женщины все"])
    ws.append(["Заплыв", "Дорожка", "Ф. И.", "Год рождения", "Команда", "Заявочное время"])
    ws.append([1, 1, "Козлов Андрей", 2008, "Команда для теста", "3.43.54"])
    ws.append([None, 2, "Белова Екатерина", 2010, "Команда для теста", "1.38.47"])
    ws.append([])
    ws.append(["100 брасс, мужчины все"])
    ws.append(["Заплыв", "Дорожка", "Ф. И.", "Год рождения", "Команда", "Заявочное время"])
    ws.append([2, 1, "Морозов Михаил", 2012, "Команда для теста", "30.34.53"])
    ws.append([None, 2, "Смирнов Алексей", 2013, "Команда для теста", "3.21.50"])
    wb.save(file_path)

    data = import_excel(file_path)

    assert list(data) == ["100 брасс, женщины все", "100 брасс, мужчины все"]
    assert [swimmer["full_name"] for swimmer in data["100 брасс, женщины все"]] == [
        "Козлов Андрей",
        "Белова Екатерина",
    ]
    assert [swimmer["full_name"] for swimmer in data["100 брасс, мужчины все"]] == [
        "Морозов Михаил",
        "Смирнов Алексей",
    ]
