from __future__ import annotations

from dataclasses import dataclass
from typing import Optional


@dataclass(slots=True)
class Swimmer:
    id: int
    event_id: int
    full_name: str
    birth_year: Optional[int]
    team: Optional[str]
    seed_time_raw: Optional[str]
    seed_time_cs: Optional[int]
    heat: Optional[int]
    lane: Optional[int]
    status: str = "OK"


@dataclass(slots=True)
class Event:
    id: int
    name: str
    lanes_count: int
