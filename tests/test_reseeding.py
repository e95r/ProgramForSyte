from core.models import Swimmer
from core.reseeding import compress_lanes_within_heats, full_reseed


def make_swimmer(id_: int, heat: int | None, lane: int | None, status: str = "OK", seed: int | None = None):
    return Swimmer(
        id=id_,
        event_id=1,
        full_name=f"S{id_}",
        birth_year=2010,
        team="T",
        seed_time_raw=None,
        seed_time_cs=seed,
        heat=heat,
        lane=lane,
        status=status,
    )


def test_soft_compress_lanes_keeps_heats():
    swimmers = [
        make_swimmer(1, 1, 1),
        make_swimmer(2, 1, 2, status="DNS"),
        make_swimmer(3, 1, 3),
        make_swimmer(4, 2, 2),
    ]

    out = compress_lanes_within_heats(swimmers, lanes_count=6)
    active = {s.id: s for s in out if s.status != "DNS"}
    assert active[1].heat == 1 and active[1].lane == 3
    assert active[3].heat == 1 and active[3].lane == 4
    assert active[4].heat == 2 and active[4].lane == 3


def test_full_reseed_by_seed_time():
    swimmers = [
        make_swimmer(1, 2, 3, seed=6100),
        make_swimmer(2, 1, 1, seed=5900),
        make_swimmer(3, 3, 4, seed=6500),
        make_swimmer(4, 1, 2, status="DNS", seed=5000),
    ]

    out = full_reseed(swimmers, lanes_count=2)
    active = [s for s in out if s.status != "DNS"]
    assert [(s.id, s.heat, s.lane) for s in active] == [(1, 1, 1), (3, 1, 2), (2, 2, 1)]


def test_full_reseed_places_fastest_swimmer_in_last_heat():
    swimmers = [
        make_swimmer(1, 1, 1, seed=7000),
        make_swimmer(2, 1, 2, seed=6800),
        make_swimmer(3, 1, 3, seed=6400),
        make_swimmer(4, 1, 4, seed=6200),
        make_swimmer(5, 1, 5, seed=5900),
    ]

    out = full_reseed(swimmers, lanes_count=2)
    active = [s for s in out if s.status != "DNS"]
    heats: dict[int, set[int]] = {}
    for swimmer in active:
        heats.setdefault(swimmer.heat or 0, set()).add(swimmer.id)

    assert heats == {1: {1, 2}, 2: {3, 4}, 3: {5}}


def test_lane_order_for_six_lanes_places_fastest_in_center_lanes():
    swimmers = [
        make_swimmer(1, 1, None, seed=7000),
        make_swimmer(2, 1, None, seed=6900),
        make_swimmer(3, 1, None, seed=6800),
        make_swimmer(4, 1, None, seed=6700),
        make_swimmer(5, 1, None, seed=6600),
        make_swimmer(6, 1, None, seed=6500),
    ]

    out = compress_lanes_within_heats(swimmers, lanes_count=6)
    assert [(s.id, s.lane) for s in out if s.status != "DNS"] == [(2, 1), (4, 2), (6, 3), (5, 4), (3, 5), (1, 6)]
