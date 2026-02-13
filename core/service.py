from __future__ import annotations

import shutil
from datetime import datetime
from pathlib import Path
from core.db import MeetRepository
from core.reseeding import compress_lanes_within_heats, full_reseed
from core.time_utils import parse_seed_time_to_cs


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

    def import_startlist(self, excel_path: Path) -> None:
        self.repo.clear_all()
        from core.excel_importer import import_excel

        imported = import_excel(excel_path)
        for event_name, swimmers in imported.items():
            event_id = self.repo.upsert_event(event_name)
            self.repo.add_swimmers(event_id, swimmers)
        self.repo.log("import_excel", str(excel_path))

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
            updated = compress_lanes_within_heats(swimmers)
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
        return self._build_protocol_html(
            event.name,
            swimmers,
            grouped=grouped,
            sort_by=sort_by,
            sort_desc=sort_desc,
            group_by=group_by,
        )

    def build_final_protocol(
        self,
        grouped: bool = True,
        sort_by: str = "place",
        sort_desc: bool = False,
        group_by: str = "heat",
    ) -> str:
        blocks: list[str] = [
            "<style>@page { size: A4 portrait; margin: 12mm; } body { font-family: Arial, sans-serif; }"
            " h1, h2 { margin: 0 0 8px 0; } table { margin-bottom: 16px; font-size: 12px; border-collapse: collapse; }"
            " th, td { border: 1px solid #333; padding: 4px; }</style>",
            "<h1>Итоговый протокол соревнований</h1>",
        ]
        for event in self.repo.list_events():
            swimmers = self.repo.list_swimmers(event.id)
            blocks.append(
                self._build_protocol_html(
                    event.name,
                    swimmers,
                    grouped=grouped,
                    with_title=True,
                    sort_by=sort_by,
                    sort_desc=sort_desc,
                    group_by=group_by,
                )
            )
        return "\n".join(blocks)

    def _build_protocol_html(
        self,
        title: str,
        swimmers: list,
        grouped: bool,
        with_title: bool = True,
        sort_by: str = "place",
        sort_desc: bool = False,
        group_by: str = "heat",
    ) -> str:
        active = [s for s in swimmers if s.status != "DNS"]
        ranked = sorted(active, key=lambda s: (s.result_time_cs is None, s.result_time_cs or 99999999, s.full_name))
        places = {s.id: idx + 1 for idx, s in enumerate(ranked) if s.result_time_cs is not None}

        def row_html(s, place: str) -> str:
            return (
                "<tr>"
                f"<td>{s.heat or '-'} / {s.lane or '-'}</td>"
                f"<td>{s.full_name}</td>"
                f"<td>{s.team or ''}</td>"
                f"<td>{s.seed_time_raw or ''}</td>"
                f"<td>{s.result_time_raw or ''}</td>"
                f"<td>{s.result_mark or ''}</td>"
                f"<td>{place}</td>"
                "</tr>"
            )

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
            return (s.result_time_cs is None, s.result_time_cs or 99999999, s.full_name.lower())

        def group_key(s):
            if group_by == "status":
                status = (s.status or "").strip()
                return (status.lower(),), status or "Статус не указан"
            if group_by == "team":
                label = (s.team or "").strip()
                return (label.lower(),), label or "Без команды"
            if group_by == "birth_year":
                year = s.birth_year
                return (year is None, year or 0), str(year) if year else "Год не указан"
            if group_by == "mark":
                mark = (s.result_mark or "").strip()
                return (mark == "", mark), mark or "Без отметки"
            if group_by == "lane":
                lane = s.lane
                return (lane is None, lane or 0), f"Дорожка {lane}" if lane else "Без дорожки"
            heat_key = s.heat or 999
            return (heat_key,), "Без заплыва" if heat_key == 999 else f"Заплыв {heat_key}"

        rows: list[str] = []
        if grouped:
            groups: dict[tuple, dict[str, object]] = {}
            for s in swimmers:
                key, label = group_key(s)
                if key not in groups:
                    groups[key] = {"label": label, "rows": []}
                groups[key]["rows"].append(s)
            for group in sorted(groups):
                rows.append(f"<tr><td colspan='7'><b>{groups[group]['label']}</b></td></tr>")
                for s in sorted(groups[group]["rows"], key=sort_key, reverse=sort_desc):
                    place = str(places.get(s.id, ""))
                    rows.append(row_html(s, place))
        else:
            for s in sorted(swimmers, key=sort_key, reverse=sort_desc):
                place = str(places.get(s.id, ""))
                rows.append(row_html(s, place))

        heading = f"<h2>{title}</h2>" if with_title else ""
        return (
            f"{heading}"
            "<table border='1' cellspacing='0' cellpadding='4' width='100%'>"
            "<tr><th>Заплыв/дорожка</th><th>ФИО</th><th>Команда</th><th>Заявка</th><th>Результат</th><th>Отм.</th><th>Место</th></tr>"
            + "".join(rows)
            + "</table>"
        )
