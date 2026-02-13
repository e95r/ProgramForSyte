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


def test_event_protocol_sort_by_team(tmp_path: Path):
    service = MeetService(tmp_path)
    try:
        event_id = service.repo.upsert_event("50m freestyle")
        service.repo.add_swimmers(
            event_id,
            [
                {"full_name": "B", "team": "Sharks"},
                {"full_name": "A", "team": "Dolphins"},
            ],
        )

        html = service.build_event_protocol(event_id, grouped=False, sort_by="team")
        assert html.index("<td>A</td>") < html.index("<td>B</td>")
    finally:
        service.close()


def test_event_protocol_grouped_by_status(tmp_path: Path):
    service = MeetService(tmp_path)
    try:
        event_id = service.repo.upsert_event("100m freestyle")
        service.repo.add_swimmers(
            event_id,
            [
                {"full_name": "A", "status": "DNS"},
                {"full_name": "B", "status": "OK"},
            ],
        )

        html = service.build_event_protocol(event_id, grouped=True, group_by="status")
        assert "<b>DNS</b>" in html
        assert "<b>OK</b>" in html
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


def test_final_protocol_grouped_and_sorted_by_mark(tmp_path: Path):
    service = MeetService(tmp_path)
    try:
        event_id = service.repo.upsert_event("200m")
        service.repo.add_swimmers(
            event_id,
            [
                {"full_name": "A", "heat": 1, "result_mark": "DNS"},
                {"full_name": "B", "heat": 1, "result_mark": "DQ"},
                {"full_name": "C", "heat": 1, "result_mark": ""},
            ],
        )

        html = service.build_final_protocol(grouped=True, sort_by="mark")
        heat_block_start = html.index("<b>Заплыв 1</b>")
        dns_pos = html.index("DNS", heat_block_start)
        dq_pos = html.index("DQ", heat_block_start)
        c_pos = html.index("<td>C</td>", heat_block_start)
        assert dns_pos < dq_pos < c_pos
    finally:
        service.close()


def test_final_protocol_place_sort_desc(tmp_path: Path):
    service = MeetService(tmp_path)
    try:
        event_id = service.repo.upsert_event("50m")
        service.repo.add_swimmers(
            event_id,
            [
                {"full_name": "Fast", "result_time_raw": "00:30:00", "result_time_cs": 3000},
                {"full_name": "Slow", "result_time_raw": "00:35:00", "result_time_cs": 3500},
            ],
        )

        html = service.build_final_protocol(grouped=False, sort_by="place", sort_desc=True)
        assert html.index("<td>Slow</td>") < html.index("<td>Fast</td>")
    finally:
        service.close()


def test_final_protocol_grouped_by_team(tmp_path: Path):
    service = MeetService(tmp_path)
    try:
        event_id = service.repo.upsert_event("100m")
        service.repo.add_swimmers(
            event_id,
            [
                {"full_name": "A", "team": "Sharks"},
                {"full_name": "B", "team": "Dolphins"},
                {"full_name": "C", "team": None},
            ],
        )

        html = service.build_final_protocol(grouped=True, group_by="team")
        assert "<b>Dolphins</b>" in html
        assert "<b>Sharks</b>" in html
        assert "<b>Без команды</b>" in html
    finally:
        service.close()
