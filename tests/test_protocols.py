from pathlib import Path

from core.service import MeetService


def test_event_protocol_grouped_by_heats(tmp_path: Path):
    service = MeetService(tmp_path)
    try:
        event_id = service.repo.upsert_event("50m freestyle")
        service.repo.add_swimmers(
            event_id,
            [
                {"full_name": "A", "heat": 2, "lane": 1, "seed_time_raw": "00:35:00", "seed_time_cs": 3500},
                {"full_name": "B", "heat": 1, "lane": 2, "seed_time_raw": "00:33:00", "seed_time_cs": 3300},
            ],
        )
        swimmers = service.repo.list_swimmers(event_id)
        service.save_event_results(
            event_id,
            [
                {"swimmer_id": str(swimmers[0].id), "result_time_raw": "00:34:00", "result_mark": ""},
                {"swimmer_id": str(swimmers[1].id), "result_time_raw": "00:32:50", "result_mark": ""},
            ],
        )

        html = service.build_event_protocol(event_id, grouped=True)
        assert "Заплыв 1" in html
        assert "Заплыв 2" in html
        assert "1</td>" in html
    finally:
        service.close()


def test_final_protocol_contains_all_events(tmp_path: Path):
    service = MeetService(tmp_path)
    try:
        e1 = service.repo.upsert_event("50m")
        e2 = service.repo.upsert_event("100m")
        service.repo.add_swimmers(e1, [{"full_name": "A", "heat": 1, "lane": 1}])
        service.repo.add_swimmers(e2, [{"full_name": "B", "heat": 1, "lane": 2}])

        html = service.build_final_protocol(grouped=True)
        assert "50m" in html
        assert "100m" in html
        assert "Итоговый протокол соревнований" in html
        assert "size: A4" in html
    finally:
        service.close()


def test_final_protocol_sort_by_mark(tmp_path: Path):
    service = MeetService(tmp_path)
    try:
        event_id = service.repo.upsert_event("200m")
        service.repo.add_swimmers(
            event_id,
            [
                {"full_name": "A", "result_mark": "DNS"},
                {"full_name": "B", "result_mark": "DQ"},
                {"full_name": "C", "result_mark": ""},
            ],
        )

        html = service.build_final_protocol(grouped=False, sort_by="mark")
        assert html.index("DNS") < html.index("DQ") < html.index("<td>C</td>")
    finally:
        service.close()
