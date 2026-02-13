from __future__ import annotations

from zipfile import BadZipFile
from pathlib import Path

from core.time_utils import parse_seed_time_to_cs

HEADER_ALIASES = {
    "name": {"фи", "ф. и.", "фио", "участник", "name"},
    "year": {"год рождения", "год", "birth", "year"},
    "team": {"команда", "team"},
    "seed": {"заявочное время", "время", "seed", "entry time"},
    "heat_lane": {"заплыв/дорожка", "заплыв/дорожка ", "заплыв", "heat/lane"},
}


class ExcelImportError(ValueError):
    """Raised when a startlist file cannot be parsed as supported Excel."""


def _normalize(value: object) -> str:
    return str(value or "").strip().lower()


def _find_columns(header_row: list[object]) -> dict[str, int]:
    mapping: dict[str, int] = {}
    for idx, raw in enumerate(header_row):
        h = _normalize(raw)
        for key, aliases in HEADER_ALIASES.items():
            if h in aliases and key not in mapping:
                mapping[key] = idx
    return mapping


def _parse_heat_lane(text: object) -> tuple[int | None, int | None]:
    raw = str(text or "").strip().replace(" ", "")
    if not raw or "/" not in raw:
        return None, None
    a, b = raw.split("/", 1)
    if a.isdigit() and b.isdigit():
        return int(a), int(b)
    return None, None


def import_excel(path: Path) -> dict[str, list[dict]]:
    def _file_debug_message(file_path: Path) -> str:
        exists = file_path.exists()
        size = file_path.stat().st_size if exists else 0
        suffix = file_path.suffix.lower()
        return (
            f"Selected: {file_path}\n"
            f"Exists: {exists}\n"
            f"Size: {size}\n"
            f"Suffix: {suffix}"
        )

    def _validate_input_file(file_path: Path) -> None:
        if not file_path.exists():
            raise ExcelImportError("Выбранный файл не существует.")
        size = file_path.stat().st_size
        suffix = file_path.suffix.lower()
        if size == 0:
            raise ExcelImportError("Выбранный файл пустой (0 байт). Выберите корректный Excel-файл.")
        if suffix not in {".xlsx", ".xlsm"}:
            raise ExcelImportError("Поддерживаются только файлы .xlsx и .xlsm.")

    initial_debug_message = _file_debug_message(path)
    print(initial_debug_message)
    _validate_input_file(path)

    if path.suffix.lower() == ".xls":
        raise ExcelImportError("Формат .xls не поддерживается. Сохраните файл как .xlsx и попробуйте снова.")

    try:
        from openpyxl import load_workbook
        from openpyxl.utils.exceptions import InvalidFileException
    except ModuleNotFoundError as exc:
        raise ExcelImportError("Не установлен пакет openpyxl. Установите зависимости приложения.") from exc

    try:
        pre_load_debug_message = _file_debug_message(path)
        print(pre_load_debug_message)
        _validate_input_file(path)
        wb = load_workbook(path, data_only=True)
    except (BadZipFile, InvalidFileException) as exc:
        raise ExcelImportError(
            "Не удалось открыть Excel-файл. Проверьте, что это корректный .xlsx/.xlsm файл."
        ) from exc

    result: dict[str, list[dict]] = {}

    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue

        header_idx = 0
        for i, row in enumerate(rows[:10]):
            joined = " ".join(_normalize(c) for c in row)
            if "заплыв" in joined or "фи" in joined or "name" in joined:
                header_idx = i
                break

        header = list(rows[header_idx])
        cols = _find_columns(header)
        if "name" not in cols:
            continue

        swimmers: list[dict] = []
        for row in rows[header_idx + 1 :]:
            name = str(row[cols["name"]] or "").strip()
            if not name:
                continue
            year = row[cols["year"]] if "year" in cols else None
            team = row[cols["team"]] if "team" in cols else None
            seed_raw = row[cols["seed"]] if "seed" in cols else None
            seed_raw_text = str(seed_raw).strip() if seed_raw is not None else None
            heat, lane = (None, None)
            if "heat_lane" in cols:
                heat, lane = _parse_heat_lane(row[cols["heat_lane"]])

            swimmers.append(
                {
                    "full_name": name,
                    "birth_year": int(year) if isinstance(year, (int, float)) else None,
                    "team": str(team).strip() if team is not None else None,
                    "seed_time_raw": seed_raw_text,
                    "seed_time_cs": parse_seed_time_to_cs(seed_raw_text),
                    "heat": heat,
                    "lane": lane,
                    "status": "OK",
                }
            )

        if swimmers:
            result[ws.title] = swimmers

    return result
