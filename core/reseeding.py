from __future__ import annotations

from collections import defaultdict
from typing import Iterable

from core.models import Swimmer


def lane_order(lanes_count: int) -> list[int]:
    if lanes_count <= 0:
        return []
    if lanes_count == 1:
        return [1]
    if lanes_count % 2 == 0:
        left = lanes_count // 2
        right = left + 1
        order: list[int] = []
        while left >= 1 or right <= lanes_count:
            if left >= 1:
                order.append(left)
                left -= 1
            if right <= lanes_count:
                order.append(right)
                right += 1
        return order

    middle = lanes_count // 2 + 1
    order = [middle]
    offset = 1
    while len(order) < lanes_count:
        right = middle + offset
        left = middle - offset
        if right <= lanes_count:
            order.append(right)
        if left >= 1:
            order.append(left)
        offset += 1
    return order


def _seed_sort_key(swimmer: Swimmer) -> tuple[bool, int, str]:
    return (swimmer.seed_time_cs is None, swimmer.seed_time_cs or 10**12, swimmer.full_name.lower())


def _assign_lanes(swimmers: list[Swimmer], lanes_count: int) -> list[Swimmer]:
    order = lane_order(lanes_count)
    ranked = sorted(swimmers, key=_seed_sort_key)
    for idx, swimmer in enumerate(ranked):
        swimmer.lane = order[idx]
    return ranked


def compress_lanes_within_heats(swimmers: Iterable[Swimmer], lanes_count: int = 6) -> list[Swimmer]:
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
        _assign_lanes(active, lanes_count=max(len(active), lanes_count if lanes_count > 0 else len(active)))
        for swimmer in active:
            swimmer.heat = heat
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

    active.sort(key=lambda x: (x.seed_time_cs is None, -(x.seed_time_cs or -1), x.full_name.lower()))

    reseeded: list[Swimmer] = []
    for heat_idx in range(0, len(active), lanes_count):
        heat_number = heat_idx // lanes_count + 1
        heat_swimmers = active[heat_idx:heat_idx + lanes_count]
        _assign_lanes(heat_swimmers, lanes_count)
        for swimmer in heat_swimmers:
            swimmer.heat = heat_number
        reseeded.extend(heat_swimmers)

    return sorted(reseeded, key=lambda x: (x.heat or 999, x.lane or 999, x.full_name)) + dns
