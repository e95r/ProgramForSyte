from __future__ import annotations

import shutil
from datetime import datetime
from pathlib import Path

from core.db import MeetRepository
from core.reseeding import compress_lanes_within_heats, full_reseed


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
