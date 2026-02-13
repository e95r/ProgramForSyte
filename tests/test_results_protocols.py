from pathlib import Path

from core.service import MeetService


def test_save_results_and_protocols(tmp_path: Path):
    service = MeetService(tmp_path)
    event_id = service.repo.upsert_event("100 free", lanes_count=8)
    service.repo.add_swimmers(
        event_id,
        [
            {"full_name": "A", "heat": 1, "lane": 1},
            {"full_name": "B", "heat": 1, "lane": 2},
            {"full_name": "C", "heat": 1, "lane": 3},
        ],
    )
    swimmers = service.repo.list_swimmers(event_id)

    service.save_event_results(
        event_id,
        [
            {"swimmer_id": swimmers[0].id, "result_time_raw": "00:59:10", "result_status": "OK"},
            {"swimmer_id": swimmers[1].id, "result_time_raw": "01:00:30", "result_status": "OK"},
            {"swimmer_id": swimmers[2].id, "result_time_raw": "", "result_status": "DQ"},
        ],
    )

    event_protocol = service.save_event_protocol(event_id)
    event_protocol_by_heat = service.save_event_protocol(event_id, sort_mode="heat")
    start_protocol = service.save_start_protocol(event_id)
    final_protocol = service.save_final_protocol(sort_mode="heat")

    assert event_protocol.exists()
    assert event_protocol_by_heat.exists()
    assert start_protocol.exists()
    assert final_protocol.exists()
    protocol_text = event_protocol.read_text(encoding="utf-8")
    assert "Протокол дистанции" in protocol_text
    assert "DQ" in protocol_text
    assert "1" in protocol_text

    by_heat_text = event_protocol_by_heat.read_text(encoding="utf-8")
    assert "Заплыв 1" in by_heat_text

    start_text = start_protocol.read_text(encoding="utf-8")
    assert "Стартовый протокол" in start_text
    assert "Заявочное время" in start_text

    final_text = final_protocol.read_text(encoding="utf-8")
    assert "Итоговый протокол соревнования" in final_text
    assert "100 free" in final_text
    assert "Заплыв 1" in final_text

    service.close()
