from __future__ import annotations

from typing import Any


DEFAULT_EMAIL_STATUSES = ["Success", "Failed", "Pending"]


def pending_as_failed_rows(rows: list[dict[str, Any]], status_column: str) -> tuple[list[dict[str, Any]], int]:
    output: list[dict[str, Any]] = []
    converted = 0

    for row in rows:
        item = dict(row)
        column = _status_column(item, status_column)
        if column and _status_key(item.get(column)) == "pending":
            item[column] = "Failed"
            converted += 1
        output.append(item)

    return output, converted


def _status_column(row: dict[str, Any], configured_column: str) -> str | None:
    candidates = [
        configured_column,
        "ValidationStatus",
        "CRMStatus",
    ]
    for candidate in candidates:
        if candidate in row:
            return candidate
    return None


def _status_key(value: Any) -> str:
    return str(value or "").strip().casefold().replace(" ", "").replace("_", "").replace("-", "")
