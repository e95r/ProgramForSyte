from pathlib import Path

from core.service import MeetService


def test_search_by_name_across_all_events(tmp_path: Path):
    service = MeetService(tmp_path)
    first_event_id = service.repo.upsert_event("50 вольный стиль", lanes_count=8)
    second_event_id = service.repo.upsert_event("100 баттерфляй", lanes_count=8)
    service.repo.add_swimmers(
        first_event_id,
        [
            {"full_name": "Иван Петров", "heat": 1, "lane": 1},
            {"full_name": "Алексей Смирнов", "heat": 1, "lane": 2},
        ],
    )
    service.repo.add_swimmers(
        second_event_id,
        [
            {"full_name": "Мария Петрова", "heat": 1, "lane": 3},
        ],
    )

    result = service.repo.list_swimmers(None, "Петров")
    assert [(swimmer.full_name, swimmer.event_name) for swimmer in result] == [
        ("Иван Петров", "50 вольный стиль"),
        ("Мария Петрова", "100 баттерфляй"),
    ]
    service.close()


def test_search_handles_multiple_tokens_and_extra_spaces(tmp_path: Path):
    service = MeetService(tmp_path)
    event_id = service.repo.upsert_event("200 комплекс", lanes_count=8)
    service.repo.add_swimmers(
        event_id,
        [
            {"full_name": "Иван Сергеев Петров", "heat": 1, "lane": 1},
            {"full_name": "Иван Петрович", "heat": 1, "lane": 2},
        ],
    )

    result = service.repo.list_swimmers(event_id, "  иван    сергеев ")
    assert [swimmer.full_name for swimmer in result] == ["Иван Сергеев Петров"]
    service.close()
