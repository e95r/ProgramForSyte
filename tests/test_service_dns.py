from pathlib import Path

from core.service import MeetService


def test_mark_dns_only_marks_absent(tmp_path: Path):
    service = MeetService(tmp_path)
    event_id = service.repo.upsert_event("50 free", lanes_count=6)
    service.repo.add_swimmers(
        event_id,
        [
            {"full_name": "A", "heat": 1, "lane": 1, "seed_time_raw": "1.00", "seed_time_cs": 6000},
            {"full_name": "B", "heat": 1, "lane": 2, "seed_time_raw": "1.01", "seed_time_cs": 6100},
            {"full_name": "C", "heat": 1, "lane": 3, "seed_time_raw": "1.02", "seed_time_cs": 6200},
        ],
    )

    swimmers = service.repo.list_swimmers(event_id)
    service.mark_dns(event_id, [swimmers[1].id])
    out = service.repo.list_swimmers(event_id)

    active = [s for s in out if s.status == "OK"]
    assert [(s.full_name, s.lane) for s in active] == [("A", 1), ("C", 3)]

    dns = [s for s in out if s.status == "DNS"]
    assert dns and dns[0].full_name == "B"
    assert dns[0].heat is None and dns[0].lane is None
    service.close()


def test_restore_swimmer_reseeds_event(tmp_path: Path):
    service = MeetService(tmp_path)
    event_id = service.repo.upsert_event("50 free", lanes_count=6)
    service.repo.add_swimmers(
        event_id,
        [
            {"full_name": "A", "heat": 1, "lane": 1, "seed_time_raw": "1.00", "seed_time_cs": 6000},
            {"full_name": "B", "heat": 1, "lane": 2, "seed_time_raw": "1.01", "seed_time_cs": 6100},
            {"full_name": "C", "heat": 1, "lane": 3, "seed_time_raw": "1.02", "seed_time_cs": 6200},
        ],
    )

    swimmers = service.repo.list_swimmers(event_id)
    swimmer_b = next(s for s in swimmers if s.full_name == "B")
    service.mark_dns(event_id, [swimmer_b.id])
    service.restore_swimmers(event_id, [swimmer_b.id], mode="full")

    out = service.repo.list_swimmers(event_id)
    active = [s for s in out if s.status == "OK"]
    assert [(s.full_name, s.lane) for s in active] == [("A", 1), ("B", 2), ("C", 3)]
    service.close()
