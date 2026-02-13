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
    result_time_raw: Optional[str] = None
    result_time_cs: Optional[int] = None
    result_status: str = "OK"
    status: str = "OK"


@dataclass(slots=True)
class Event:
    id: int
    name: str
    lanes_count: int
