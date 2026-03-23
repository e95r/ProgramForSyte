from __future__ import annotations

import hashlib
import html
import json
import re
import secrets
import shutil
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from core.db import MeetRepository
from core.models import Secretary
from core.reseeding import compress_lanes_within_heats, full_reseed
from core.time_utils import parse_seed_time_to_cs


EVENT_NAME_PARTS_RE = re.compile(
    r"^\s*(?P<base>.+?)(?:\s*,\s*(?P<gender>женщины|девушки|девочки|мужчины|юноши|мальчики|все))?(?:\s+(?P<age>все))?\s*$",
    re.IGNORECASE,
)
TRAILING_ALL_RE = re.compile(r"^(?P<title>.+?)(?:\s*,)?\s+все\s*$", re.IGNORECASE)


class MeetService:
    def __init__(self, root: Path):
        self.root = root
        self.meet_dir = self.root / "meet"
        self.backup_dir = self.meet_dir / "backups"
        self.db_path = self.meet_dir / "meet.db"
        self.backup_dir.mkdir(parents=True, exist_ok=True)
        if self.db_path.exists():
            self.create_backup(reason="startup")
        self.repo = MeetRepository(self.db_path)

    def close(self) -> None:
        self.repo.close()

    def create_backup(self, reason: str = "manual") -> Path | None:
        if not self.db_path.exists():
            return None
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        backup_path = self.backup_dir / f"meet-{reason}-{stamp}.db"
        shutil.copy2(self.db_path, backup_path)
        return backup_path

    def secretary_count(self) -> int:
        return self.repo.secretary_count()

    def _hash_password(self, password: str, salt: str | None = None) -> str:
        salt = salt or secrets.token_hex(16)
        digest = hashlib.sha256(f"{salt}:{password}".encode("utf-8")).hexdigest()
        return f"{salt}${digest}"

    def _verify_password(self, password: str, stored_hash: str) -> bool:
        try:
            salt, digest = stored_hash.split("$", 1)
        except ValueError:
            return False
        return self._hash_password(password, salt) == f"{salt}${digest}"

    def register_secretary(
        self,
        username: str,
        password: str,
        password_hint: str,
        display_name: str = "",
    ) -> Secretary:
        username = username.strip()
        display_name = display_name.strip() or username
        password_hint = password_hint.strip()
        if not username:
            raise ValueError("Укажите логин секретаря")
        if len(password) < 4:
            raise ValueError("Пароль должен содержать минимум 4 символа")
        if not password_hint:
            raise ValueError("Укажите подсказку для восстановления пароля")
        if self.repo.get_secretary_by_username(username) is not None:
            raise ValueError("Секретарь с таким логином уже зарегистрирован")
        secretary_id = self.repo.create_secretary(
            username=username,
            display_name=display_name,
            password_hash=self._hash_password(password),
            password_hint=password_hint,
        )
        self.repo.log("register_secretary", f"username={username}")
        return Secretary(secretary_id, username, display_name, password_hint)

    def authenticate_secretary(self, username: str, password: str) -> Secretary | None:
        row = self.repo.get_secretary_auth_row(username.strip())
        if row is None or not self._verify_password(password, row["password_hash"]):
            return None
        secretary = Secretary(
            id=int(row["id"]),
            username=row["username"],
            display_name=row["display_name"],
            password_hint=row["password_hint"],
        )
        self.repo.log("authenticate_secretary", f"username={secretary.username}")
        return secretary

    def get_secretary_password_hint(self, username: str) -> str | None:
        secretary = self.repo.get_secretary_by_username(username.strip())
        if secretary is None:
            return None
        self.repo.log("password_hint_requested", f"username={secretary.username}")
        return secretary.password_hint

    def import_startlist(self, excel_path: Path) -> None:
        self.repo.clear_all()
        from core.excel_importer import extract_meet_metadata, import_excel

        imported = import_excel(excel_path)
        metadata = extract_meet_metadata(excel_path)
        self.repo.set_meta("competition_title", metadata.get("competition_title") or self._derive_competition_title(excel_path))
        self.repo.set_meta("competition_date", metadata.get("competition_date") or "")
        self.repo.set_meta("competition_place", metadata.get("competition_place") or "")
        self.repo.set_meta("age_groups", json.dumps(metadata.get("age_groups") or [], ensure_ascii=False))
        self.repo.set_meta("relay_age_groups", json.dumps(metadata.get("relay_age_groups") or [], ensure_ascii=False))
        for event_name, swimmers in imported.items():
            normalized_event_name = self._normalize_imported_event_name(event_name)
            lanes_count = self._infer_imported_lanes_count(swimmers)
            event_id = self.repo.upsert_event(normalized_event_name, lanes_count=lanes_count)
            swimmers = self._normalize_imported_start_protocol(swimmers, lanes_count=lanes_count)
            self.repo.add_swimmers(event_id, swimmers)
        self.repo.log("import_excel", str(excel_path))

    def _derive_competition_title(self, excel_path: Path) -> str:
        raw_title = excel_path.stem.replace("_", " ").replace("-", " ")
        normalized = " ".join(raw_title.split())
        return normalized or "Итоговый протокол соревнований"

    def _normalize_imported_event_name(self, event_name: str) -> str:
        normalized = " ".join(event_name.split()).strip(" ,")
        match = TRAILING_ALL_RE.match(normalized)
        if match:
            normalized = match.group("title").strip(" ,")
        return normalized or event_name.strip()


    def _infer_imported_lanes_count(self, swimmers: list[dict], default: int = 8) -> int:
        lanes = [int(lane) for lane in (s.get("lane") for s in swimmers) if isinstance(lane, int) and lane > 0]
        if lanes:
            return max(lanes)

        heats: dict[int, int] = {}
        for swimmer in swimmers:
            heat = swimmer.get("heat")
            if isinstance(heat, int) and heat > 0:
                heats[heat] = heats.get(heat, 0) + 1
        if heats:
            return max(heats.values())

        return default

    def _normalize_imported_start_protocol(self, swimmers: list[dict], lanes_count: int = 8) -> list[dict]:
        active = [dict(s) for s in swimmers if s.get("status") != "DNS"]
        if active and all(s.get("heat") is not None and s.get("lane") is not None for s in active):
            ordered_active = sorted(
                active,
                key=lambda s: (
                    s.get("heat") is None,
                    s.get("heat") or 10**12,
                    s.get("lane") is None,
                    s.get("lane") or 10**12,
                    (s.get("full_name") or "").lower(),
                ),
            )
            dns = [dict(s) for s in swimmers if s.get("status") == "DNS"]
            for swimmer in dns:
                swimmer["heat"] = None
                swimmer["lane"] = None
            return ordered_active + dns
        return self._rebuild_start_protocol(swimmers, lanes_count=lanes_count)

    def _rebuild_start_protocol(self, swimmers: list[dict], lanes_count: int = 8) -> list[dict]:
        active = [dict(s) for s in swimmers if s.get("status") != "DNS"]
        dns = [dict(s) for s in swimmers if s.get("status") == "DNS"]
        active.sort(
            key=lambda s: (
                s.get("seed_time_cs") is None,
                s.get("seed_time_cs") or 10**12,
                (s.get("full_name") or "").lower(),
            )
        )
        for idx, swimmer in enumerate(active):
            swimmer["heat"] = idx // lanes_count + 1
            swimmer["lane"] = idx % lanes_count + 1
        for swimmer in dns:
            swimmer["heat"] = None
            swimmer["lane"] = None
        return active + dns

    def mark_dns(self, event_id: int, swimmer_ids: list[int]) -> None:
        self.repo.set_dns(swimmer_ids)
        self.repo.log("mark_dns", f"event={event_id}; ids={swimmer_ids}")

    def restore_swimmers(self, event_id: int, swimmer_ids: list[int], mode: str = "soft") -> None:
        self.repo.restore_swimmers(swimmer_ids)
        self.reseed_event(event_id, mode=mode)
        self.repo.log("restore_swimmers", f"event={event_id}; ids={swimmer_ids}; mode={mode}")

    def reseed_event(self, event_id: int, mode: str = "soft") -> None:
        event = next(e for e in self.repo.list_events() if e.id == event_id)
        swimmers = self.repo.list_swimmers(event_id)
        if mode == "full":
            updated = full_reseed(swimmers, lanes_count=event.lanes_count)
        else:
            updated = compress_lanes_within_heats(swimmers, lanes_count=event.lanes_count)
        self.repo.update_swimmer_positions(updated)
        self.repo.log("reseed_event", f"event={event_id}; mode={mode}")

    def save_event_results(self, event_id: int, results: list[dict[str, str]]) -> None:
        payload: list[tuple[int, str | None, int | None, str | None]] = []
        for row in results:
            swimmer_id = int(row["swimmer_id"])
            raw = row.get("result_time_raw", "").strip() or None
            mark = row.get("result_mark", "").strip() or None
            cs = parse_seed_time_to_cs(raw) if raw else None
            payload.append((swimmer_id, raw, cs, mark))
        self.repo.save_results(payload)
        self.repo.log("save_event_results", f"event={event_id}; rows={len(payload)}")

    def build_event_protocol(
        self,
        event_id: int,
        grouped: bool = True,
        sort_by: str = "place",
        sort_desc: bool = False,
        group_by: str = "heat",
    ) -> str:
        event = next(e for e in self.repo.list_events() if e.id == event_id)
        swimmers = self.repo.list_swimmers(event_id)
        results_mode = any(self._swimmer_has_result(swimmer) for swimmer in swimmers)
        return self._build_protocol_document(
            page_title=self.repo.get_meta("competition_title") or event.name,
            events=[(event.name, swimmers)],
            final_mode=results_mode,
            sort_by=sort_by,
            sort_desc=sort_desc,
            group_by=group_by,
            grouped=grouped,
        )

    def build_final_protocol(
        self,
        grouped: bool = True,
        sort_by: str = "place",
        sort_desc: bool = False,
        group_by: str = "heat",
    ) -> str:
        events = []
        for event in self.repo.list_events():
            swimmers = self._filter_final_protocol_swimmers(self.repo.list_swimmers(event.id))
            events.append((event.name, swimmers))
        return self._build_protocol_document(
            page_title=f"Итоговый протокол {self.repo.get_meta('competition_title') or 'соревнований'}",
            events=events,
            final_mode=True,
            sort_by=sort_by,
            sort_desc=sort_desc,
            group_by=group_by,
            grouped=grouped,
        )

    def export_event_protocol_xlsx(
        self,
        path: Path,
        event_id: int,
        grouped: bool = True,
        sort_by: str = "place",
        sort_desc: bool = False,
        group_by: str = "heat",
    ) -> None:
        event = next(e for e in self.repo.list_events() if e.id == event_id)
        swimmers = self.repo.list_swimmers(event_id)
        results_mode = any(self._swimmer_has_result(swimmer) for swimmer in swimmers)
        self._export_protocol_workbook(
            path=path,
            page_title=self.repo.get_meta("competition_title") or event.name,
            events=[(event.name, swimmers)],
            final_mode=results_mode,
            sort_by=sort_by,
            sort_desc=sort_desc,
            group_by=group_by,
            grouped=grouped,
        )

    def export_final_protocol_xlsx(
        self,
        path: Path,
        grouped: bool = True,
        sort_by: str = "place",
        sort_desc: bool = False,
        group_by: str = "heat",
    ) -> None:
        events = []
        for event in self.repo.list_events():
            swimmers = self._filter_final_protocol_swimmers(self.repo.list_swimmers(event.id))
            events.append((event.name, swimmers))
        self._export_protocol_workbook(
            path=path,
            page_title=f"Итоговый протокол {self.repo.get_meta('competition_title') or 'соревнований'}",
            events=events,
            final_mode=True,
            sort_by=sort_by,
            sort_desc=sort_desc,
            group_by=group_by,
            grouped=grouped,
        )

    def _build_protocol_document(
        self,
        page_title: str,
        events: list[tuple[str, list]],
        final_mode: bool,
        sort_by: str,
        sort_desc: bool,
        group_by: str,
        grouped: bool,
    ) -> str:
        title = self.repo.get_meta("competition_title") or "Итоговый протокол соревнований"
        date = self.repo.get_meta("competition_date") or ""
        place = self.repo.get_meta("competition_place") or ""
        body: list[str] = [self._protocol_styles()]
        doc_title = title if not final_mode else page_title
        body.append(f"<div class='doc-title'>{html.escape(doc_title)}</div>")
        body.append("<table class='meta-table'>")
        body.append(f"<tr><td>{html.escape(date)}</td></tr>")
        body.append(f"<tr><td>{html.escape(place)}</td></tr>")
        body.append("</table>")

        for table in self._build_protocol_tables(events, final_mode, sort_by, sort_desc, group_by, grouped):
            body.append(self._render_protocol_table_html(table))
        return "\n".join(body)

    def _protocol_styles(self) -> str:
        return (
            "<style>"
            "@page { size: A4 portrait; margin: 12mm; }"
            "body { font-family: Arial, sans-serif; color: #000; }"
            ".doc-title { text-align: center; font-size: 24px; font-weight: 700; margin: 0 0 18px 0; }"
            ".meet-line { font-size: 16px; margin: 0 0 8px 0; }"
            ".meta-table { width: 100%; border-collapse: collapse; margin: 0 0 16px 0; }"
            ".meta-table td { border: 1px solid #cfcfcf; padding: 8px 10px; text-align: center; font-size: 16px; }"
            ".protocol-table { width: 100%; border-collapse: collapse; margin: 0 0 18px 0; font-size: 13px; table-layout: fixed; }"
            ".protocol-table col.col-num, .protocol-table col.col-heat, .protocol-table col.col-lane, .protocol-table col.col-year, .protocol-table col.col-time, .protocol-table col.col-place { width: 12%; }"
            ".protocol-table col.col-name { width: 28%; }"
            ".protocol-table col.col-team { width: 24%; }"
            ".protocol-table th, .protocol-table td { border: 1px solid #cfcfcf; padding: 6px 8px; vertical-align: middle; overflow-wrap: anywhere; word-break: break-word; }"
            ".protocol-table th { text-align: left; }"
            ".category-title { color: #fff; text-align: center; font-weight: 700; font-size: 18px; text-transform: uppercase; padding: 12px 8px; }"
            ".category-title.boys { background: #5b9bd5; }"
            ".category-title.girls { background: #e06666; }"
            ".category-title.mixed { background: #808080; }"
            ".num, .place, .year, .heat, .lane, .time { text-align: center; }"
            ".name { text-align: left; }"
            ".heat-label td { background: #f1f1f1; font-weight: 700; text-align: center; }"
            ".heat-merged { vertical-align: middle; font-weight: 700; }"
            "</style>"
        )

    def _split_event_into_age_groups(self, event_name: str, swimmers: list) -> list[dict]:
        age_groups = self._load_age_groups(event_name)
        gender_label, gender_color = self._detect_gender(event_name)
        base_title = self._base_event_name(event_name)
        if not age_groups:
            return [{
                "base_title": base_title,
                "age_label": "Все возраста",
                "gender_label": gender_label,
                "gender_color": gender_color,
                "swimmers": list(swimmers),
            }]

        result: list[dict] = []
        for group in age_groups:
            filtered = [s for s in swimmers if self._matches_age_group(s.birth_year, group)]
            if filtered:
                result.append(
                    {
                        "base_title": base_title,
                        "age_label": group["label"],
                        "gender_label": gender_label,
                        "gender_color": gender_color,
                        "swimmers": filtered,
                    }
                )
        if result:
            return result
        return [{
            "base_title": base_title,
            "age_label": "Все возраста",
            "gender_label": gender_label,
            "gender_color": gender_color,
            "swimmers": list(swimmers),
        }]

    def _load_age_groups(self, event_name: str) -> list[dict]:
        raw = self.repo.get_meta("relay_age_groups" if self._is_relay_event(event_name) else "age_groups") or "[]"
        try:
            data = json.loads(raw)
        except json.JSONDecodeError:
            return []
        return [group for group in data if isinstance(group, dict) and group.get("label")]

    def _is_relay_event(self, event_name: str) -> bool:
        lowered = event_name.lower()
        return "эстаф" in lowered or "x" in lowered or "х" in lowered or "×" in lowered

    def _matches_age_group(self, birth_year: int | None, group: dict) -> bool:
        if birth_year is None:
            return False
        min_year = group.get("min_year")
        max_year = group.get("max_year")
        if min_year is not None and birth_year < int(min_year):
            return False
        if max_year is not None and birth_year > int(max_year):
            return False
        return True

    def _parse_event_name(self, event_name: str) -> tuple[str, str]:
        normalized = " ".join(event_name.split())
        match = EVENT_NAME_PARTS_RE.match(normalized)
        if not match:
            return event_name.strip(), "Все"

        base_name = (match.group("base") or "").strip(" ,") or event_name.strip()
        gender_token = (match.group("gender") or "").lower()
        age_token = (match.group("age") or "").lower()

        if gender_token in {"женщины", "девушки", "девочки"}:
            return base_name, "Женщины"
        if gender_token in {"мужчины", "юноши", "мальчики"}:
            return base_name, "Мужчины"
        if gender_token == "все" or age_token == "все":
            return base_name, "Все"
        return base_name, "Все"

    def _detect_gender(self, event_name: str) -> tuple[str, str]:
        _, gender_label = self._parse_event_name(event_name)
        if gender_label == "Женщины":
            return gender_label, "girls"
        if gender_label == "Мужчины":
            return gender_label, "boys"
        return gender_label, "mixed"

    def _base_event_name(self, event_name: str) -> str:
        base_name, _ = self._parse_event_name(event_name)
        return base_name

    def _build_protocol_tables(
        self,
        events: list[tuple[str, list]],
        final_mode: bool,
        sort_by: str,
        sort_desc: bool,
        group_by: str,
        grouped: bool,
    ) -> list[dict]:
        tables: list[dict] = []
        for event_name, swimmers in events:
            groups = self._split_event_into_age_groups(event_name, swimmers)
            for group in groups:
                tables.append(
                    self._build_age_group_table_data(
                        event_name=event_name,
                        swimmers=group["swimmers"],
                        age_label=group["age_label"],
                        gender_label=group["gender_label"],
                        gender_color=group["gender_color"],
                        final_mode=final_mode,
                        sort_by=sort_by,
                        sort_desc=sort_desc,
                        group_by=group_by,
                        grouped=grouped,
                    )
                )
        return tables

    def _build_age_group_table_data(
        self,
        event_name: str,
        swimmers: list,
        age_label: str,
        gender_label: str,
        gender_color: str,
        final_mode: bool,
        sort_by: str,
        sort_desc: bool,
        group_by: str,
        grouped: bool,
    ) -> str:
        ranked, places = self._rank_swimmers(swimmers)
        title_parts = [self._base_event_name(event_name), gender_label]
        if gender_label == "Все":
            title_parts.append(age_label)
        elif age_label != "Все возраста":
            title_parts.append(age_label)
        title = ", ".join(title_parts).upper()
        if final_mode:
            headers = [
                {"label": "№", "class": "num"},
                {"label": "Фамилия Имя", "class": "name"},
                {"label": "Год рождения", "class": "year"},
                {"label": "Команда", "class": "team"},
                {"label": "Время", "class": "time"},
                {"label": "Место", "class": "place"},
            ]
            ordered = self._sort_protocol_rows(swimmers, places, sort_by, sort_desc, final_mode=True)
            rows = [
                [
                    {"value": index, "class": "num"},
                    {"value": swimmer.full_name, "class": "name"},
                    {"value": swimmer.birth_year or "", "class": "year"},
                    {"value": swimmer.team or "", "class": "team"},
                    {"value": self._display_swimmer_time(swimmer), "class": "time"},
                    {"value": places.get(swimmer.id, ""), "class": "place"},
                ]
                for index, swimmer in enumerate(ordered, start=1)
            ]
        else:
            headers = [
                {"label": "Заплыв", "class": "heat"},
                {"label": "Дорожка", "class": "lane"},
                {"label": "Ф. И.", "class": "name"},
                {"label": "Год рождения", "class": "year"},
                {"label": "Команда", "class": "team"},
                {"label": "Заявочное время", "class": "time"},
            ]
            rows = []
            if grouped and group_by == "heat":
                ordered = sorted(swimmers, key=lambda s: (s.heat is None, s.heat or 999, s.lane is None, s.lane or 999, s.full_name.lower()))
                heat_sizes: dict[int | None, int] = {}
                for swimmer in ordered:
                    heat_sizes[swimmer.heat] = heat_sizes.get(swimmer.heat, 0) + 1
                rendered_heats: set[int | None] = set()
                for swimmer in ordered:
                    row = []
                    if swimmer.heat not in rendered_heats:
                        rendered_heats.add(swimmer.heat)
                        row.append(
                            {"value": swimmer.heat or "", "class": "heat heat-merged", "rowspan": heat_sizes.get(swimmer.heat, 1)}
                        )
                    row.extend(
                        [
                            {"value": swimmer.lane or "", "class": "lane"},
                            {"value": swimmer.full_name, "class": "name"},
                            {"value": swimmer.birth_year or "", "class": "year"},
                            {"value": swimmer.team or "", "class": "team"},
                            {"value": swimmer.seed_time_raw or "", "class": "time"},
                        ]
                    )
                    rows.append(row)
            else:
                ordered = self._sort_protocol_rows(swimmers, places, sort_by, sort_desc, final_mode=False)
                for swimmer in ordered:
                    rows.append(
                        [
                            {"value": swimmer.heat or "", "class": "heat"},
                            {"value": swimmer.lane or "", "class": "lane"},
                            {"value": swimmer.full_name, "class": "name"},
                            {"value": swimmer.birth_year or "", "class": "year"},
                            {"value": swimmer.team or "", "class": "team"},
                            {"value": swimmer.seed_time_raw or "", "class": "time"},
                        ]
                    )
        return {"title": title, "gender_color": gender_color, "headers": headers, "rows": rows}

    def _render_protocol_table_html(self, table: dict) -> str:
        parts = ["<table class='protocol-table'>", "<colgroup>"]
        for header in table["headers"]:
            parts.append(f"<col class='col-{header['class'].split()[0]}'>")
        parts.append("</colgroup>")
        parts.append(
            f"<tr><td class='category-title {table['gender_color']}' colspan='{len(table['headers'])}'>{html.escape(table['title'])}</td></tr>"
        )
        parts.append("<tr>")
        for header in table["headers"]:
            parts.append(f"<th class='{header['class']}'>{html.escape(str(header['label']))}</th>")
        parts.append("</tr>")
        for row in table["rows"]:
            parts.append("<tr>")
            for cell in row:
                rowspan = f" rowspan='{cell['rowspan']}'" if cell.get("rowspan") else ""
                parts.append(f"<td class='{cell['class']}'{rowspan}>{html.escape(str(cell['value']))}</td>")
            parts.append("</tr>")
        parts.append("</table>")
        return "".join(parts)

    def _export_protocol_workbook(
        self,
        path: Path,
        page_title: str,
        events: list[tuple[str, list]],
        final_mode: bool,
        sort_by: str,
        sort_desc: bool,
        group_by: str,
        grouped: bool,
    ) -> None:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Протокол"

        title = self.repo.get_meta("competition_title") or "Итоговый протокол соревнований"
        doc_title = title if not final_mode else page_title
        date = self.repo.get_meta("competition_date") or ""
        place = self.repo.get_meta("competition_place") or ""
        tables = self._build_protocol_tables(events, final_mode, sort_by, sort_desc, group_by, grouped)

        border = Border(
            left=Side(style="thin", color="CFCFCF"),
            right=Side(style="thin", color="CFCFCF"),
            top=Side(style="thin", color="CFCFCF"),
            bottom=Side(style="thin", color="CFCFCF"),
        )
        centered = Alignment(horizontal="center", vertical="center", wrap_text=True)
        left_wrapped = Alignment(horizontal="left", vertical="center", wrap_text=True)
        fills = {"boys": "5B9BD5", "girls": "E06666", "mixed": "808080"}
        column_widths = [12, 12, 28, 14, 24, 14]
        for idx, width in enumerate(column_widths, start=1):
            sheet.column_dimensions[chr(64 + idx)].width = width

        row = 1
        sheet.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
        title_cell = sheet.cell(row=row, column=1, value=doc_title)
        title_cell.font = Font(size=16, bold=True)
        title_cell.alignment = centered
        row += 1
        for meta_value in (date, place):
            sheet.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
            cell = sheet.cell(row=row, column=1, value=meta_value)
            cell.border = border
            cell.alignment = centered
            row += 1
        row += 1

        for table in tables:
            sheet.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
            cell = sheet.cell(row=row, column=1, value=table["title"])
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(fill_type="solid", fgColor=fills[table["gender_color"]])
            cell.alignment = centered
            cell.border = border
            sheet.row_dimensions[row].height = 24
            row += 1

            for column, header in enumerate(table["headers"], start=1):
                cell = sheet.cell(row=row, column=column, value=header["label"])
                cell.font = Font(bold=True)
                cell.border = border
                cell.alignment = centered if header["class"] not in {"name", "team"} else left_wrapped
            sheet.row_dimensions[row].height = 24
            row += 1

            for data_row in table["rows"]:
                column = len(table["headers"]) - len(data_row) + 1
                for data_cell in data_row:
                    rowspan = int(data_cell.get("rowspan", 1))
                    if rowspan > 1:
                        sheet.merge_cells(start_row=row, start_column=column, end_row=row + rowspan - 1, end_column=column)
                    cell = sheet.cell(row=row, column=column, value=data_cell["value"])
                    cell.border = border
                    cell.alignment = centered if data_cell["class"].split()[0] in {"num", "heat", "lane", "year", "time", "place"} else left_wrapped
                    column += 1
                sheet.row_dimensions[row].height = 30
                row += 1
            row += 1

        workbook.save(path)

    def _rank_swimmers(self, swimmers: list) -> tuple[list, dict[int, int]]:
        active = [s for s in swimmers if s.status != "DNS"]
        ranked = sorted(active, key=lambda s: (self._ranking_time_cs(s) is None, self._ranking_time_cs(s) or 99999999, s.full_name.lower()))
        places: dict[int, int] = {}
        last_time_cs = None
        for idx, swimmer in enumerate(ranked, start=1):
            current_time_cs = self._ranking_time_cs(swimmer)
            if current_time_cs is None:
                continue
            if current_time_cs != last_time_cs:
                places[swimmer.id] = idx
                last_time_cs = current_time_cs
                continue
            places[swimmer.id] = places[ranked[idx - 2].id]
        return ranked, places

    def _ranking_time_cs(self, swimmer) -> int | None:
        return swimmer.result_time_cs if swimmer.result_time_cs is not None else swimmer.seed_time_cs

    def _display_swimmer_time(self, swimmer) -> str:
        return swimmer.result_time_raw or swimmer.seed_time_raw or swimmer.result_mark or ""

    def _swimmer_has_result(self, swimmer) -> bool:
        return swimmer.result_time_cs is not None or bool((swimmer.result_mark or "").strip())

    def _sort_protocol_rows(self, swimmers: list, places: dict[int, int], sort_by: str, sort_desc: bool, final_mode: bool) -> list:
        def sort_key(s):
            if sort_by == "id":
                return (s.id,)
            if sort_by == "team":
                team = (s.team or "").strip()
                return (team == "", team.lower(), s.full_name.lower())
            if sort_by == "birth_year":
                return (s.birth_year is None, s.birth_year or 0, s.full_name.lower())
            if sort_by == "seed_time":
                return (s.seed_time_cs is None, s.seed_time_cs or 99999999, s.full_name.lower())
            if sort_by == "result_time":
                return (s.result_time_cs is None, s.result_time_cs or 99999999, s.full_name.lower())
            if sort_by == "heat":
                return (s.heat is None, s.heat or 999, s.lane is None, s.lane or 999, s.full_name.lower())
            if sort_by == "lane":
                return (s.lane is None, s.lane or 999, s.heat is None, s.heat or 999, s.full_name.lower())
            if sort_by == "status":
                return ((s.status or "").lower(), s.full_name.lower())
            if sort_by == "mark":
                mark = (s.result_mark or "").strip()
                return (mark == "", mark, s.full_name.lower())
            if sort_by == "full_name":
                return (s.full_name.lower(),)
            if final_mode:
                return (places.get(s.id, 10**9), s.full_name.lower())
            return (s.heat is None, s.heat or 999, s.lane is None, s.lane or 999, s.full_name.lower())

        return sorted(swimmers, key=sort_key, reverse=sort_desc)

    def _filter_final_protocol_swimmers(self, swimmers: list) -> list:
        disallowed_marks = {"DNS", "DQ"}
        return [
            s
            for s in swimmers
            if s.status != "DNS" and ((s.result_mark or "").strip().upper() not in disallowed_marks)
        ]
