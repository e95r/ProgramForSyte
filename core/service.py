from __future__ import annotations

import shutil
from datetime import datetime
from pathlib import Path
from typing import Iterable

from core.db import MeetRepository
from core.reseeding import compress_lanes_within_heats, full_reseed
from core.time_utils import format_cs, parse_seed_time_to_cs


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

    def mark_dns(self, event_id: int, swimmer_ids: list[int], mode: str = "soft") -> None:
        self.repo.set_dns(swimmer_ids)
        event = next(e for e in self.repo.list_events() if e.id == event_id)
        swimmers = self.repo.list_swimmers(event_id)
        if mode == "full":
            updated = full_reseed(swimmers, lanes_count=event.lanes_count)
        else:
            updated = compress_lanes_within_heats(swimmers)
        self.repo.update_swimmer_positions(updated)
        self.repo.log("mark_dns", f"event={event_id}; ids={swimmer_ids}; mode={mode}")

    def save_event_results(self, event_id: int, results: Iterable[dict]) -> None:
        for row in results:
            status = row.get("result_status", "OK")
            raw = (row.get("result_time_raw") or "").strip() or None
            time_cs = parse_seed_time_to_cs(raw) if raw else None
            if status != "OK":
                raw = None
                time_cs = None
            self.repo.set_result(int(row["swimmer_id"]), raw, time_cs, status)
        self.repo.log("save_event_results", f"event={event_id}")

    @staticmethod
    def _group_by_heats(swimmers: list) -> dict[str, list]:
        grouped: dict[str, list] = {}
        for swimmer in swimmers:
            heat_key = f"Заплыв {swimmer.heat}" if swimmer.heat is not None else "Заплыв -"
            grouped.setdefault(heat_key, []).append(swimmer)
        return grouped

    def _build_event_protocol_by_result(self, event_id: int) -> str:
        events = {e.id: e for e in self.repo.list_events()}
        event = events[event_id]
        swimmers = self.repo.list_event_results(event_id)
        lines = [f"Протокол дистанции: {event.name}", "=" * 80]
        lines.append("Место | ФИО | Год | Команда | Заплыв/дорожка | Результат")
        place = 0
        for swimmer in swimmers:
            result = swimmer.result_status if swimmer.result_status != "OK" else (format_cs(swimmer.result_time_cs) or "-")
            if swimmer.result_status == "OK" and swimmer.result_time_cs is not None:
                place += 1
                place_display = str(place)
            else:
                place_display = "-"
            lines.append(
                f"{place_display:>5} | {swimmer.full_name} | {swimmer.birth_year or '-'} | {swimmer.team or '-'} | "
                f"{swimmer.heat or '-'} / {swimmer.lane or '-'} | {result}"
            )
        return "\n".join(lines)

    def _build_event_protocol_by_heat(self, event_id: int) -> str:
        events = {e.id: e for e in self.repo.list_events()}
        event = events[event_id]
        swimmers = self.repo.list_swimmers(event_id)
        lines = [f"Протокол дистанции: {event.name}", "=" * 80]
        groups = self._group_by_heats(swimmers)
        for heat_name, heat_swimmers in groups.items():
            lines.append("")
            lines.append(heat_name)
            lines.append("Дорожка | ФИО | Год | Команда | Статус | Результат")
            for swimmer in heat_swimmers:
                result = swimmer.result_status if swimmer.result_status != "OK" else (format_cs(swimmer.result_time_cs) or "-")
                lines.append(
                    f"{(swimmer.lane or '-'):>7} | {swimmer.full_name} | {swimmer.birth_year or '-'} | "
                    f"{swimmer.team or '-'} | {swimmer.status} | {result}"
                )
        return "\n".join(lines)

    def build_event_protocol_text(self, event_id: int, sort_mode: str = "result") -> str:
        if sort_mode == "heat":
            return self._build_event_protocol_by_heat(event_id)
        return self._build_event_protocol_by_result(event_id)

    def build_start_protocol_text(self, event_id: int) -> str:
        events = {e.id: e for e in self.repo.list_events()}
        event = events[event_id]
        swimmers = self.repo.list_swimmers(event_id)
        lines = [f"Стартовый протокол: {event.name}", "=" * 80]
        groups = self._group_by_heats(swimmers)
        for heat_name, heat_swimmers in groups.items():
            lines.append("")
            lines.append(heat_name)
            lines.append("Дорожка | ФИО | Год | Команда | Заявочное время | Статус")
            for swimmer in heat_swimmers:
                lines.append(
                    f"{(swimmer.lane or '-'):>7} | {swimmer.full_name} | {swimmer.birth_year or '-'} | "
                    f"{swimmer.team or '-'} | {swimmer.seed_time_raw or '-'} | {swimmer.status}"
                )
        return "\n".join(lines)

    def save_event_protocol(self, event_id: int, output: Path | None = None, sort_mode: str = "result") -> Path:
        if output is None:
            output_dir = self.meet_dir / "protocols"
            output_dir.mkdir(parents=True, exist_ok=True)
            output = output_dir / f"event-{event_id}-protocol.txt"
        output.write_text(self.build_event_protocol_text(event_id, sort_mode=sort_mode), encoding="utf-8")
        self.repo.log("save_event_protocol", str(output))
        return output

    def save_start_protocol(self, event_id: int, output: Path | None = None) -> Path:
        if output is None:
            output_dir = self.meet_dir / "protocols"
            output_dir.mkdir(parents=True, exist_ok=True)
            output = output_dir / f"event-{event_id}-start-protocol.txt"
        output.write_text(self.build_start_protocol_text(event_id), encoding="utf-8")
        self.repo.log("save_start_protocol", str(output))
        return output

    def build_final_protocol_text(self, sort_mode: str = "result") -> str:
        lines = ["Итоговый протокол соревнования", "=" * 80]
        for event in self.repo.list_events():
            lines.append("")
            lines.append(self.build_event_protocol_text(event.id, sort_mode=sort_mode))
        return "\n".join(lines)

    def save_final_protocol(self, output: Path | None = None, sort_mode: str = "result") -> Path:
        if output is None:
            output_dir = self.meet_dir / "protocols"
            output_dir.mkdir(parents=True, exist_ok=True)
            stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
            output = output_dir / f"final-protocol-{stamp}.txt"
        output.write_text(self.build_final_protocol_text(sort_mode=sort_mode), encoding="utf-8")
        self.repo.log("save_final_protocol", str(output))
        return output
