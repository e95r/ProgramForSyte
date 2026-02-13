from __future__ import annotations

import sqlite3
from pathlib import Path
from typing import Iterable

from core.models import Event, Swimmer


SCHEMA = """
CREATE TABLE IF NOT EXISTS events (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL UNIQUE,
    lanes_count INTEGER NOT NULL DEFAULT 8
);

CREATE TABLE IF NOT EXISTS swimmers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    event_id INTEGER NOT NULL,
    full_name TEXT NOT NULL,
    birth_year INTEGER,
    team TEXT,
    seed_time_raw TEXT,
    seed_time_cs INTEGER,
    heat INTEGER,
    lane INTEGER,
    status TEXT NOT NULL DEFAULT 'OK',
    FOREIGN KEY(event_id) REFERENCES events(id)
);

CREATE INDEX IF NOT EXISTS idx_swimmers_event ON swimmers(event_id);
CREATE INDEX IF NOT EXISTS idx_swimmers_name ON swimmers(full_name);

CREATE TABLE IF NOT EXISTS audit_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    action TEXT NOT NULL,
    details TEXT
);
"""


class MeetRepository:
    def __init__(self, db_path: Path):
        self.db_path = db_path
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self.conn = sqlite3.connect(self.db_path)
        self.conn.row_factory = sqlite3.Row
        self.conn.executescript(SCHEMA)
        self.conn.commit()

    def close(self) -> None:
        self.conn.close()

    def log(self, action: str, details: str = "") -> None:
        self.conn.execute(
            "INSERT INTO audit_log(action, details) VALUES (?, ?)",
            (action, details),
        )
        self.conn.commit()

    def clear_all(self) -> None:
        self.conn.execute("DELETE FROM swimmers")
        self.conn.execute("DELETE FROM events")
        self.log("clear_all", "wipe imported data")
        self.conn.commit()

    def upsert_event(self, name: str, lanes_count: int = 8) -> int:
        self.conn.execute(
            "INSERT OR IGNORE INTO events(name, lanes_count) VALUES (?, ?)",
            (name, lanes_count),
        )
        row = self.conn.execute("SELECT id FROM events WHERE name=?", (name,)).fetchone()
        assert row
        return int(row["id"])

    def add_swimmers(self, event_id: int, swimmers: Iterable[dict]) -> None:
        self.conn.executemany(
            """
            INSERT INTO swimmers(
                event_id, full_name, birth_year, team, seed_time_raw, seed_time_cs, heat, lane, status
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            [
                (
                    event_id,
                    s["full_name"],
                    s.get("birth_year"),
                    s.get("team"),
                    s.get("seed_time_raw"),
                    s.get("seed_time_cs"),
                    s.get("heat"),
                    s.get("lane"),
                    s.get("status", "OK"),
                )
                for s in swimmers
            ],
        )
        self.conn.commit()

    def list_events(self) -> list[Event]:
        rows = self.conn.execute("SELECT * FROM events ORDER BY id").fetchall()
        return [Event(id=int(r["id"]), name=r["name"], lanes_count=int(r["lanes_count"])) for r in rows]

    def list_swimmers(self, event_id: int, search: str = "") -> list[Swimmer]:
        sql = "SELECT * FROM swimmers WHERE event_id=?"
        params: list[object] = [event_id]
        if search:
            sql += " AND lower(full_name) LIKE ?"
            params.append(f"%{search.lower()}%")
        sql += " ORDER BY CASE WHEN heat IS NULL THEN 999 ELSE heat END, CASE WHEN lane IS NULL THEN 999 ELSE lane END, full_name"
        rows = self.conn.execute(sql, params).fetchall()
        return [
            Swimmer(
                id=int(r["id"]),
                event_id=int(r["event_id"]),
                full_name=r["full_name"],
                birth_year=r["birth_year"],
                team=r["team"],
                seed_time_raw=r["seed_time_raw"],
                seed_time_cs=r["seed_time_cs"],
                heat=r["heat"],
                lane=r["lane"],
                status=r["status"],
            )
            for r in rows
        ]

    def update_swimmer_positions(self, swimmers: Iterable[Swimmer]) -> None:
        self.conn.executemany(
            "UPDATE swimmers SET heat=?, lane=?, status=? WHERE id=?",
            [(s.heat, s.lane, s.status, s.id) for s in swimmers],
        )
        self.conn.commit()

    def set_dns(self, swimmer_ids: list[int]) -> None:
        if not swimmer_ids:
            return
        placeholders = ",".join(["?"] * len(swimmer_ids))
        self.conn.execute(
            f"UPDATE swimmers SET status='DNS', heat=NULL, lane=NULL WHERE id IN ({placeholders})",
            swimmer_ids,
        )
        self.conn.commit()

    def restore_swimmers(self, swimmer_ids: list[int]) -> None:
        if not swimmer_ids:
            return
        placeholders = ",".join(["?"] * len(swimmer_ids))
        self.conn.execute(
            f"UPDATE swimmers SET status='OK' WHERE id IN ({placeholders})",
            swimmer_ids,
        )
        self.conn.commit()
