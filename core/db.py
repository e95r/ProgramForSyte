from __future__ import annotations

import sqlite3
from pathlib import Path
from typing import Iterable

from core.models import Event, Secretary, Swimmer


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
    result_time_raw TEXT,
    result_time_cs INTEGER,
    result_mark TEXT,
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


CREATE TABLE IF NOT EXISTS secretaries (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT NOT NULL UNIQUE,
    display_name TEXT NOT NULL,
    password_hash TEXT NOT NULL,
    password_hint TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS meet_meta (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL DEFAULT ''
);
"""


class MeetRepository:
    def __init__(self, db_path: Path):
        self.db_path = db_path
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self.conn = sqlite3.connect(self.db_path)
        self.conn.row_factory = sqlite3.Row
        self.conn.executescript(SCHEMA)
        self._migrate_schema()
        self.conn.commit()

    def _migrate_schema(self) -> None:
        existing_columns = {
            row["name"] for row in self.conn.execute("PRAGMA table_info(swimmers)").fetchall()
        }
        for column, ddl in (
            ("result_time_raw", "ALTER TABLE swimmers ADD COLUMN result_time_raw TEXT"),
            ("result_time_cs", "ALTER TABLE swimmers ADD COLUMN result_time_cs INTEGER"),
            ("result_mark", "ALTER TABLE swimmers ADD COLUMN result_mark TEXT"),
        ):
            if column not in existing_columns:
                self.conn.execute(ddl)


    def secretary_count(self) -> int:
        row = self.conn.execute("SELECT COUNT(*) AS cnt FROM secretaries").fetchone()
        return int(row["cnt"]) if row else 0

    def create_secretary(
        self,
        username: str,
        display_name: str,
        password_hash: str,
        password_hint: str,
    ) -> int:
        cursor = self.conn.execute(
            """
            INSERT INTO secretaries(username, display_name, password_hash, password_hint)
            VALUES (?, ?, ?, ?)
            """,
            (username, display_name, password_hash, password_hint),
        )
        self.conn.commit()
        return int(cursor.lastrowid)

    def get_secretary_by_username(self, username: str) -> Secretary | None:
        row = self.conn.execute(
            "SELECT id, username, display_name, password_hint FROM secretaries WHERE lower(username)=lower(?)",
            (username,),
        ).fetchone()
        if row is None:
            return None
        return Secretary(
            id=int(row["id"]),
            username=row["username"],
            display_name=row["display_name"],
            password_hint=row["password_hint"],
        )

    def get_secretary_auth_row(self, username: str) -> sqlite3.Row | None:
        return self.conn.execute(
            "SELECT * FROM secretaries WHERE lower(username)=lower(?)",
            (username,),
        ).fetchone()

    def list_secretaries(self) -> list[Secretary]:
        rows = self.conn.execute(
            "SELECT id, username, display_name, password_hint FROM secretaries ORDER BY id"
        ).fetchall()
        return [
            Secretary(
                id=int(r["id"]),
                username=r["username"],
                display_name=r["display_name"],
                password_hint=r["password_hint"],
            )
            for r in rows
        ]

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

    def set_meta(self, key: str, value: str) -> None:
        self.conn.execute(
            "INSERT INTO meet_meta(key, value) VALUES (?, ?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",
            (key, value),
        )
        self.conn.commit()

    def get_meta(self, key: str) -> str | None:
        row = self.conn.execute("SELECT value FROM meet_meta WHERE key=?", (key,)).fetchone()
        return row["value"] if row else None

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
                event_id, full_name, birth_year, team, seed_time_raw, seed_time_cs, heat, lane, status,
                result_time_raw, result_time_cs, result_mark
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
                    s.get("status", "OK"),
                    s.get("result_time_raw"),
                    s.get("result_time_cs"),
                    s.get("result_mark"),
                )
                for s in swimmers
            ],
        )
        self.conn.commit()

    def list_events(self) -> list[Event]:
        rows = self.conn.execute("SELECT * FROM events ORDER BY id").fetchall()
        return [Event(id=int(r["id"]), name=r["name"], lanes_count=int(r["lanes_count"])) for r in rows]

    def list_swimmers(self, event_id: int | None, search: str = "") -> list[Swimmer]:
        sql = """
            SELECT swimmers.*, events.name AS event_name
            FROM swimmers
            JOIN events ON events.id = swimmers.event_id
            WHERE 1=1
        """
        params: list[object] = []
        if event_id is not None:
            sql += " AND swimmers.event_id=?"
            params.append(event_id)
        search_tokens = [token.casefold() for token in search.split() if token.strip()]
        sql += """
            ORDER BY swimmers.event_id,
                CASE WHEN swimmers.heat IS NULL THEN 999 ELSE swimmers.heat END,
                CASE WHEN swimmers.lane IS NULL THEN 999 ELSE swimmers.lane END,
                swimmers.full_name
        """
        rows = self.conn.execute(sql, params).fetchall()
        swimmers = [
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
                result_time_raw=r["result_time_raw"],
                result_time_cs=r["result_time_cs"],
                result_mark=r["result_mark"],
                event_name=r["event_name"],
            )
            for r in rows
        ]
        if not search_tokens:
            return swimmers
        return [
            swimmer
            for swimmer in swimmers
            if all(token in swimmer.full_name.casefold() for token in search_tokens)
        ]

    def update_swimmer_positions(self, swimmers: Iterable[Swimmer]) -> None:
        self.conn.executemany(
            "UPDATE swimmers SET heat=?, lane=?, status=? WHERE id=?",
            [(s.heat, s.lane, s.status, s.id) for s in swimmers],
        )
        self.conn.commit()

    def save_results(self, swimmer_results: list[tuple[int, str | None, int | None, str | None]]) -> None:
        self.conn.executemany(
            "UPDATE swimmers SET result_time_raw=?, result_time_cs=?, result_mark=? WHERE id=?",
            [(raw, cs, mark, swimmer_id) for swimmer_id, raw, cs, mark in swimmer_results],
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
