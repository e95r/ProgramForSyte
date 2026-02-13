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
    result_time_raw TEXT,
    result_time_cs INTEGER,
    result_status TEXT NOT NULL DEFAULT 'OK',
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
        self._migrate_swimmers_table()
        self.conn.commit()

    def _migrate_swimmers_table(self) -> None:
        columns = {
            row["name"]
            for row in self.conn.execute("PRAGMA table_info(swimmers)").fetchall()
        }
        if "result_time_raw" not in columns:
            self.conn.execute("ALTER TABLE swimmers ADD COLUMN result_time_raw TEXT")
        if "result_time_cs" not in columns:
            self.conn.execute("ALTER TABLE swimmers ADD COLUMN result_time_cs INTEGER")
        if "result_status" not in columns:
            self.conn.execute("ALTER TABLE swimmers ADD COLUMN result_status TEXT NOT NULL DEFAULT 'OK'")

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
                event_id, full_name, birth_year, team, seed_time_raw, seed_time_cs, heat, lane, result_time_raw, result_time_cs, result_status, status
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
                    s.get("result_time_raw"),
                    s.get("result_time_cs"),
                    s.get("result_status", "OK"),
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
                result_time_raw=r["result_time_raw"],
                result_time_cs=r["result_time_cs"],
                result_status=r["result_status"],
                status=r["status"],
            )
            for r in rows
        ]

    def set_result(self, swimmer_id: int, result_time_raw: str | None, result_time_cs: int | None, result_status: str) -> None:
        self.conn.execute(
            "UPDATE swimmers SET result_time_raw=?, result_time_cs=?, result_status=? WHERE id=?",
            (result_time_raw, result_time_cs, result_status, swimmer_id),
        )
        self.conn.commit()

    def list_event_results(self, event_id: int) -> list[Swimmer]:
        rows = self.conn.execute(
            """
            SELECT * FROM swimmers
            WHERE event_id=?
            ORDER BY
                CASE WHEN result_status='OK' AND result_time_cs IS NOT NULL THEN 0 ELSE 1 END,
                result_time_cs,
                CASE WHEN heat IS NULL THEN 999 ELSE heat END,
                CASE WHEN lane IS NULL THEN 999 ELSE lane END,
                full_name
            """,
            (event_id,),
        ).fetchall()
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
                result_time_raw=r["result_time_raw"],
                result_time_cs=r["result_time_cs"],
                result_status=r["result_status"],
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
