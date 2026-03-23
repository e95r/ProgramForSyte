from __future__ import annotations

import hashlib
import html
import json
import re
import secrets
import shutil
from datetime import datetime
from pathlib import Path

from core.db import MeetRepository
from core.models import Secretary
from core.reseeding import compress_lanes_within_heats, full_reseed
from core.time_utils import parse_seed_time_to_cs


EVENT_NAME_PARTS_RE = re.compile(
    r"^\s*(?P<base>.+?)(?:\s*,\s*(?P<gender>женщины|девушки|девочки|мужчины|юноши|мальчики|все))?(?:\s+(?P<age>все))?\s*$",
    re.IGNORECASE,
)


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
            lanes_count = self._infer_imported_lanes_count(swimmers)
            event_id = self.repo.upsert_event(event_name, lanes_count=lanes_count)
            swimmers = self._normalize_imported_start_protocol(swimmers, lanes_count=lanes_count)
            self.repo.add_swimmers(event_id, swimmers)
        self.repo.log("import_excel", str(excel_path))

    def _derive_competition_title(self, excel_path: Path) -> str:
        raw_title = excel_path.stem.replace("_", " ").replace("-", " ")
        normalized = " ".join(raw_title.split())
        return normalized or "Итоговый протокол соревнований"

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
        return self._build_protocol_document(
            page_title=event.name,
            events=[(event.name, swimmers)],
            final_mode=False,
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
        body.append(f"<div class='doc-title'>{html.escape(page_title)}</div>")
        body.append("<table class='meta-table'>")
        body.append(f"<tr><td>{html.escape(date)}</td></tr>")
        body.append(f"<tr><td>{html.escape(place)}</td></tr>")
        body.append("</table>")
        if not final_mode and title and title != page_title:
            body.append(f"<div class='meet-line'>Соревнование: {html.escape(title)}</div>")

        for event_name, swimmers in events:
            groups = self._split_event_into_age_groups(event_name, swimmers)
            for group in groups:
                body.append(
                    self._build_age_group_table_html(
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
            ".protocol-table { width: 100%; border-collapse: collapse; margin: 0 0 18px 0; font-size: 14px; }"
            ".protocol-table th, .protocol-table td { border: 1px solid #cfcfcf; padding: 6px 8px; }"
            ".protocol-table th { text-align: left; }"
            ".category-title { color: #fff; text-align: center; font-weight: 700; font-size: 18px; text-transform: uppercase; padding: 12px 8px; }"
            ".category-title.boys { background: #5b9bd5; }"
            ".category-title.girls { background: #e06666; }"
            ".category-title.mixed { background: #808080; }"
            ".num, .place, .year, .heat, .lane, .time { text-align: center; }"
            ".name { text-align: left; }"
            ".heat-label td { background: #f1f1f1; font-weight: 700; text-align: center; }"
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

    def _build_age_group_table_html(
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
        parts = ["<table class='protocol-table'>"]
        parts.append(f"<tr><td class='category-title {gender_color}' colspan='6'>{html.escape(title)}</td></tr>")
        if final_mode:
            parts.append("<tr><th class='num'>№</th><th>Фамилия Имя</th><th class='year'>Год рождения</th><th>Команда</th><th class='time'>Время</th><th class='place'>Место</th></tr>")
            ordered = self._sort_protocol_rows(swimmers, places, sort_by, sort_desc, final_mode=True)
            for index, swimmer in enumerate(ordered, start=1):
                parts.append(
                    "<tr>"
                    f"<td class='num'>{index}</td>"
                    f"<td class='name'>{html.escape(swimmer.full_name)}</td>"
                    f"<td class='year'>{swimmer.birth_year or ''}</td>"
                    f"<td>{html.escape(swimmer.team or '')}</td>"
                    f"<td class='time'>{html.escape(swimmer.result_time_raw or swimmer.seed_time_raw or '')}</td>"
                    f"<td class='place'>{places.get(swimmer.id, '')}</td>"
                    "</tr>"
                )
        else:
            parts.append("<tr><th class='heat'>Заплыв</th><th class='lane'>Дорожка</th><th>Ф. И.</th><th class='year'>Год рождения</th><th>Команда</th><th class='time'>Заявочное время</th></tr>")
            if grouped and group_by == "heat":
                ordered = sorted(swimmers, key=lambda s: (s.heat is None, s.heat or 999, s.lane is None, s.lane or 999, s.full_name.lower()))
                current_heat = object()
                for swimmer in ordered:
                    if swimmer.heat != current_heat:
                        current_heat = swimmer.heat
                    parts.append(
                        "<tr>"
                        f"<td class='heat'>{swimmer.heat or ''}</td>"
                        f"<td class='lane'>{swimmer.lane or ''}</td>"
                        f"<td class='name'>{html.escape(swimmer.full_name)}</td>"
                        f"<td class='year'>{swimmer.birth_year or ''}</td>"
                        f"<td>{html.escape(swimmer.team or '')}</td>"
                        f"<td class='time'>{html.escape(swimmer.seed_time_raw or '')}</td>"
                        "</tr>"
                    )
            else:
                ordered = self._sort_protocol_rows(swimmers, places, sort_by, sort_desc, final_mode=False)
                for swimmer in ordered:
                    parts.append(
                        "<tr>"
                        f"<td class='heat'>{swimmer.heat or ''}</td>"
                        f"<td class='lane'>{swimmer.lane or ''}</td>"
                        f"<td class='name'>{html.escape(swimmer.full_name)}</td>"
                        f"<td class='year'>{swimmer.birth_year or ''}</td>"
                        f"<td>{html.escape(swimmer.team or '')}</td>"
                        f"<td class='time'>{html.escape(swimmer.seed_time_raw or '')}</td>"
                        "</tr>"
                    )
        parts.append("</table>")
        return "".join(parts)

    def _rank_swimmers(self, swimmers: list) -> tuple[list, dict[int, int]]:
        active = [s for s in swimmers if s.status != "DNS"]
        ranked = sorted(active, key=lambda s: (s.result_time_cs is None, s.result_time_cs or 99999999, s.full_name.lower()))
        places: dict[int, int] = {}
        last_time_cs = None
        for idx, swimmer in enumerate(ranked, start=1):
            if swimmer.result_time_cs is None:
                continue
            if swimmer.result_time_cs != last_time_cs:
                places[swimmer.id] = idx
                last_time_cs = swimmer.result_time_cs
                continue
            places[swimmer.id] = places[ranked[idx - 2].id]
        return ranked, places

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
