from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, date, timedelta
from pathlib import Path
from typing import Any
import json
import re

from app.config import Settings, DUMMY_DATA_FILE
from app.dummy_data import ensure_dummy_data

STATUS_ORDER = ["Success", "Failed", "Pending", "In Progress"]


@dataclass
class RequestFilters:
    request_id: str | None = None
    date_from: date | None = None
    date_to: date | None = None
    status: str | None = None
    q: str | None = None
    page: int = 1
    page_size: int = 25

    @property
    def safe_page(self) -> int:
        return max(1, self.page)

    @property
    def safe_page_size(self) -> int:
        return min(max(1, self.page_size), 200)


def _parse_datetime(value: Any) -> datetime | None:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())
    text = str(value).replace("T", " ")
    try:
        return datetime.fromisoformat(text)
    except ValueError:
        for fmt in ["%Y-%m-%d %H:%M:%S", "%Y-%m-%d"]:
            try:
                return datetime.strptime(text[:19], fmt)
            except ValueError:
                continue
    return None


def _normalize_status_counts(counts: dict[str, int]) -> dict[str, int]:
    normalized = {status: int(counts.get(status, 0)) for status in STATUS_ORDER}
    for status, count in counts.items():
        if status not in normalized:
            normalized[status] = int(count)
    return normalized


class DummyRepository:
    def __init__(self, path: Path = DUMMY_DATA_FILE):
        self.path = path
        ensure_dummy_data(path)

    def _load_rows(self) -> list[dict[str, Any]]:
        return json.loads(self.path.read_text(encoding="utf-8"))

    def _save_rows(self, rows: list[dict[str, Any]]) -> None:
        self.path.write_text(json.dumps(rows, indent=2), encoding="utf-8")

    def _apply_filters(self, rows: list[dict[str, Any]], filters: RequestFilters) -> list[dict[str, Any]]:
        filtered: list[dict[str, Any]] = []
        request_text = (filters.request_id or "").strip().lower()
        status_text = (filters.status or "").strip().lower()
        q_text = (filters.q or "").strip().lower()

        for row in rows:
            created = _parse_datetime(row.get("CreatedAt"))
            created_date = created.date() if created else None

            if request_text and request_text not in str(row.get("RequestId", "")).lower():
                continue
            if status_text and status_text != str(row.get("CRMStatus", "")).lower():
                continue
            if filters.date_from and (created_date is None or created_date < filters.date_from):
                continue
            if filters.date_to and (created_date is None or created_date > filters.date_to):
                continue
            if q_text:
                haystack = " ".join(str(value) for value in row.values()).lower()
                if q_text not in haystack:
                    continue
            filtered.append(row)

        return filtered

    def list_requests(self, filters: RequestFilters) -> dict[str, Any]:
        rows = self._apply_filters(self._load_rows(), filters)
        rows.sort(key=lambda row: str(row.get("CreatedAt", "")), reverse=True)
        total = len(rows)
        start = (filters.safe_page - 1) * filters.safe_page_size
        end = start + filters.safe_page_size
        page_rows = rows[start:end]
        columns = list(page_rows[0].keys()) if page_rows else self.get_columns()
        counts = self.get_status_counts_from_rows(rows)
        return {
            "rows": page_rows,
            "columns": columns,
            "total": total,
            "page": filters.safe_page,
            "page_size": filters.safe_page_size,
            "status_counts": counts,
        }

    def export_rows(self, filters: RequestFilters) -> list[dict[str, Any]]:
        rows = self._apply_filters(self._load_rows(), filters)
        rows.sort(key=lambda row: str(row.get("CreatedAt", "")), reverse=True)
        return rows

    def get_status_counts_from_rows(self, rows: list[dict[str, Any]]) -> dict[str, int]:
        counts: dict[str, int] = {}
        for row in rows:
            status = str(row.get("CRMStatus") or "Unknown")
            counts[status] = counts.get(status, 0) + 1
        return _normalize_status_counts(counts)

    def get_columns(self) -> list[str]:
        rows = self._load_rows()
        if rows:
            return list(rows[0].keys())
        return []

    def retry_request(self, request_id: str) -> dict[str, Any] | None:
        rows = self._load_rows()
        for row in rows:
            if str(row.get("RequestId")) == request_id:
                row["CRMStatus"] = "Pending"
                row["UpdatedAt"] = datetime.now().replace(microsecond=0).isoformat(sep=" ")
                row["Attempts"] = int(row.get("Attempts") or 0) + 1
                row["LastError"] = "Moved back to Pending for retry."
                self._save_rows(rows)
                return row
        return None


class SqlServerRepository:
    def __init__(self, settings: Settings):
        self.settings = settings
        try:
            import pyodbc  # type: ignore
        except ImportError as exc:
            raise RuntimeError("pyodbc is required for SQL Server mode. Install requirements.txt first.") from exc
        self.pyodbc = pyodbc

    def _connect(self):
        return self.pyodbc.connect(self.settings.sql_connection_string)

    def _table_name(self) -> str:
        parts = self.settings.db_table.split(".")
        return ".".join(self._quote_identifier(part.strip("[] ")) for part in parts)

    def _column(self, name: str) -> str:
        return self._quote_identifier(name.strip("[] "))

    @staticmethod
    def _quote_identifier(identifier: str) -> str:
        if not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", identifier):
            raise ValueError(f"Unsafe SQL identifier configured: {identifier}")
        return f"[{identifier}]"

    def _where_clause(self, filters: RequestFilters) -> tuple[str, list[Any]]:
        clauses: list[str] = []
        params: list[Any] = []
        created_col = self._column(self.settings.created_at_column)
        request_col = self._column(self.settings.request_id_column)
        status_col = self._column(self.settings.status_column)

        if filters.date_from:
            clauses.append(f"{created_col} >= ?")
            params.append(datetime.combine(filters.date_from, datetime.min.time()))
        if filters.date_to:
            clauses.append(f"{created_col} < ?")
            params.append(datetime.combine(filters.date_to + timedelta(days=1), datetime.min.time()))
        if filters.request_id:
            clauses.append(f"CAST({request_col} AS NVARCHAR(255)) LIKE ?")
            params.append(f"%{filters.request_id.strip()}%")
        if filters.status:
            clauses.append(f"{status_col} = ?")
            params.append(filters.status.strip())
        if filters.q:
            search_cols = [request_col, status_col]
            clauses.append("(" + " OR ".join(f"CAST({col} AS NVARCHAR(MAX)) LIKE ?" for col in search_cols) + ")")
            params.extend([f"%{filters.q.strip()}%"] * len(search_cols))

        where = " WHERE " + " AND ".join(clauses) if clauses else ""
        return where, params

    @staticmethod
    def _cursor_rows(cursor) -> list[dict[str, Any]]:
        columns = [col[0] for col in cursor.description]
        output: list[dict[str, Any]] = []
        for db_row in cursor.fetchall():
            item: dict[str, Any] = {}
            for column, value in zip(columns, db_row):
                if isinstance(value, (datetime, date)):
                    item[column] = value.isoformat(sep=" ") if isinstance(value, datetime) else value.isoformat()
                else:
                    item[column] = value
            output.append(item)
        return output

    def list_requests(self, filters: RequestFilters) -> dict[str, Any]:
        where, params = self._where_clause(filters)
        table = self._table_name()
        created_col = self._column(self.settings.created_at_column)
        status_col = self._column(self.settings.status_column)
        offset = (filters.safe_page - 1) * filters.safe_page_size

        with self._connect() as conn:
            count_cursor = conn.cursor().execute(f"SELECT COUNT(*) FROM {table}{where}", params)
            total = int(count_cursor.fetchone()[0])

            status_cursor = conn.cursor().execute(
                f"SELECT {status_col} AS CRMStatus, COUNT(*) AS StatusCount FROM {table}{where} GROUP BY {status_col}",
                params,
            )
            counts = {str(row[0] or "Unknown"): int(row[1]) for row in status_cursor.fetchall()}

            data_cursor = conn.cursor().execute(
                f"SELECT * FROM {table}{where} ORDER BY {created_col} DESC OFFSET ? ROWS FETCH NEXT ? ROWS ONLY",
                params + [offset, filters.safe_page_size],
            )
            rows = self._cursor_rows(data_cursor)
            columns = [col[0] for col in data_cursor.description] if data_cursor.description else []

        return {
            "rows": rows,
            "columns": columns,
            "total": total,
            "page": filters.safe_page,
            "page_size": filters.safe_page_size,
            "status_counts": _normalize_status_counts(counts),
        }

    def export_rows(self, filters: RequestFilters) -> list[dict[str, Any]]:
        where, params = self._where_clause(filters)
        table = self._table_name()
        created_col = self._column(self.settings.created_at_column)
        with self._connect() as conn:
            cursor = conn.cursor().execute(f"SELECT * FROM {table}{where} ORDER BY {created_col} DESC", params)
            return self._cursor_rows(cursor)

    def retry_request(self, request_id: str) -> dict[str, Any] | None:
        table = self._table_name()
        request_col = self._column(self.settings.request_id_column)
        status_col = self._column(self.settings.status_column)
        created_col = self._column(self.settings.created_at_column)
        with self._connect() as conn:
            cursor = conn.cursor()
            cursor.execute(
                f"UPDATE {table} SET {status_col} = ?, UpdatedAt = GETDATE() WHERE CAST({request_col} AS NVARCHAR(255)) = ?",
                ["Pending", request_id],
            )
            conn.commit()
            select_cursor = conn.cursor().execute(
                f"SELECT TOP 1 * FROM {table} WHERE CAST({request_col} AS NVARCHAR(255)) = ? ORDER BY {created_col} DESC",
                [request_id],
            )
            rows = self._cursor_rows(select_cursor)
            return rows[0] if rows else None


def get_repository(settings: Settings):
    if settings.data_source == "mssql":
        return SqlServerRepository(settings)
    return DummyRepository()
