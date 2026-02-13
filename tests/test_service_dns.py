from pathlib import Path

from core.service import MeetService


def test_mark_dns_soft(tmp_path: Path):
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
    service.mark_dns(event_id, [swimmers[1].id], mode="soft")
    out = service.repo.list_swimmers(event_id)
    active = [s for s in out if s.status == "OK"]
    assert [(s.full_name, s.lane) for s in active] == [("A", 1), ("C", 2)]
    dns = [s for s in out if s.status == "DNS"]
    assert dns and dns[0].full_name == "B"
    service.close()
