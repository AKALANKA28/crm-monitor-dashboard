from __future__ import annotations

from datetime import date, datetime
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from app.config import Settings


def app_now(settings: Settings) -> datetime:
    return datetime.now(_app_zone(settings))


def app_today(settings: Settings) -> date:
    return app_now(settings).date()


def _app_zone(settings: Settings) -> ZoneInfo:
    try:
        return ZoneInfo(settings.app_timezone)
    except ZoneInfoNotFoundError as exc:
        raise ValueError(f"Invalid APP_TIMEZONE value: {settings.app_timezone}") from exc
