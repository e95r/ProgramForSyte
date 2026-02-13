from __future__ import annotations

from collections import defaultdict
from typing import Iterable

from core.models import Swimmer


def compress_lanes_within_heats(swimmers: Iterable[Swimmer]) -> list[Swimmer]:
    grouped: dict[int, list[Swimmer]] = defaultdict(list)
    unchanged: list[Swimmer] = []
    for s in swimmers:
        if s.status == "DNS":
            s.heat = None
            s.lane = None
            unchanged.append(s)
        else:
            grouped[s.heat or 1].append(s)

    result = unchanged[:]
    for heat in sorted(grouped.keys()):
        active = grouped[heat]
        active.sort(key=lambda x: (x.lane or 999, x.seed_time_cs or 10**9, x.full_name))
        for idx, swimmer in enumerate(active, start=1):
            swimmer.heat = heat
            swimmer.lane = idx
        result.extend(active)
    return sorted(result, key=lambda x: (x.heat or 999, x.lane or 999, x.full_name))


def full_reseed(swimmers: Iterable[Swimmer], lanes_count: int) -> list[Swimmer]:
    dns: list[Swimmer] = []
    active: list[Swimmer] = []
    for s in swimmers:
        if s.status == "DNS":
            s.heat = None
            s.lane = None
            dns.append(s)
        else:
            active.append(s)

    active.sort(key=lambda x: (x.seed_time_cs is None, x.seed_time_cs or 10**12, x.full_name))

    for idx, swimmer in enumerate(active):
        swimmer.heat = idx // lanes_count + 1
        swimmer.lane = idx % lanes_count + 1

    return active + dns
