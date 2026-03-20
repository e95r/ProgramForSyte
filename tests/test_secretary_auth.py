from pathlib import Path

from core.service import MeetService


def test_register_and_authenticate_multiple_secretaries(tmp_path: Path):
    service = MeetService(tmp_path)

    first = service.register_secretary(
        username="secretary-1",
        password="pass1234",
        password_hint="любимая команда",
        display_name="Главный секретарь",
    )
    second = service.register_secretary(
        username="secretary-2",
        password="pass5678",
        password_hint="цвет бейджа",
    )

    assert service.secretary_count() == 2
    assert first.display_name == "Главный секретарь"
    assert second.display_name == "secretary-2"

    auth_first = service.authenticate_secretary("secretary-1", "pass1234")
    auth_second = service.authenticate_secretary("secretary-2", "pass5678")

    assert auth_first is not None and auth_first.username == "secretary-1"
    assert auth_second is not None and auth_second.username == "secretary-2"
    service.close()


def test_secretary_password_hint_and_invalid_login(tmp_path: Path):
    service = MeetService(tmp_path)
    service.register_secretary(
        username="marshal",
        password="abcd",
        password_hint="номер кабинета",
        display_name="Секретарь заплывов",
    )

    assert service.authenticate_secretary("marshal", "wrong") is None
    assert service.get_secretary_password_hint("marshal") == "номер кабинета"
    assert service.get_secretary_password_hint("unknown") is None
    service.close()
