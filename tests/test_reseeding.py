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

    out = compress_lanes_within_heats(swimmers)
    active = {s.id: s for s in out if s.status != "DNS"}
    assert active[1].heat == 1 and active[1].lane == 1
    assert active[3].heat == 1 and active[3].lane == 2
    assert active[4].heat == 2 and active[4].lane == 1


def test_full_reseed_by_seed_time():
    swimmers = [
        make_swimmer(1, 2, 3, seed=6100),
        make_swimmer(2, 1, 1, seed=5900),
        make_swimmer(3, 3, 4, seed=6500),
        make_swimmer(4, 1, 2, status="DNS", seed=5000),
    ]

    out = full_reseed(swimmers, lanes_count=2)
    active = [s for s in out if s.status != "DNS"]
    assert [(s.id, s.heat, s.lane) for s in active] == [(2, 1, 1), (1, 1, 2), (3, 2, 1)]
