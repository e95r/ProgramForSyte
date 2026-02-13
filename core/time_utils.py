from __future__ import annotations

import re

TIME_RE = re.compile(r"^(\d{1,2})[:.](\d{2})[:.](\d{2})$")


def parse_seed_time_to_cs(value: str | float | int | None) -> int | None:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        text = f"{value:.2f}"
    else:
        text = str(value).strip().replace(",", ".")
    if not text:
        return None

    match = TIME_RE.match(text)
    if match:
        minutes = int(match.group(1))
        seconds = int(match.group(2))
        centis = int(match.group(3))
        return (minutes * 60 + seconds) * 100 + centis

    if "." in text:
        secs_part, centis_part = text.split(".", 1)
        if secs_part.isdigit() and centis_part.isdigit():
            sec = int(secs_part)
            cen = int((centis_part + "00")[:2])
            return sec * 100 + cen
    if text.isdigit():
        return int(text) * 100
    return None


def format_cs(value: int | None) -> str:
    if value is None:
        return ""
    total_seconds, centis = divmod(value, 100)
    minutes, seconds = divmod(total_seconds, 60)
    return f"{minutes:02d}:{seconds:02d}:{centis:02d}"
