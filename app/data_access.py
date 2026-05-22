from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, date, timedelta
from typing import Any
import copy
import re
import threading
import time

# Removed the dummy data imports
from app.config import Settings

STATUS_ORDER = ["Success", "Failed", "Pending", "In Progress"]

def _canonical_status(value: Any) -> str:
    text = str(value or "Unknown").strip()
    if not text:
        return "Unknown"
    normalized = text.casefold()
    mapping = {
        "success": "Success",
        "failed": "Failed",
        "pending": "Pending",
        "in progress": "In Progress",
        "inprogress": "In Progress",
        "in_progress": "In Progress",
        "in-progress": "In Progress",
    }
    return mapping.get(normalized, text)

@dataclass
class RequestFilters:
    request_id: str | None = None
    date_from: date | None = None
    date_to: date | None = None
    status: str | None = None
    status_list: list[str] | None = None
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
    rolled_up: dict[str, int] = {}
    for status, count in counts.items():
        canonical = _canonical_status(status)
        rolled_up[canonical] = rolled_up.get(canonical, 0) + int(count)

    normalized = {status: int(rolled_up.get(status, 0)) for status in STATUS_ORDER}
    for status, count in rolled_up.items():
        if status not in normalized:
            normalized[status] = int(count)
    return normalized

class SqlServerRepository:
    def __init__(self, settings: Settings):
        self.settings = settings
        try:
            import pyodbc  # type: ignore
        except ImportError as exc:
            raise RuntimeError(
                "pyodbc is required for Azure SQL/SQL Server mode. Install requirements.txt first."
            ) from exc
        self.pyodbc = pyodbc

    def _connect(self):
        return self.pyodbc.connect(self.settings.sql_connection_string)

    def _table_name(self) -> str:
        parts = self.settings.db_table.split(".")
        return ".".join(self._quote_identifier(part.strip("[] ")) for part in parts)

    def _submissions_table_name(self) -> str:
        parts = "dbo.Submissions".split(".")
        return ".".join(self._quote_identifier(part.strip("[] ")) for part in parts)

    def _outlets_table_name(self) -> str:
        parts = "dbo.Outlets".split(".")
        return ".".join(self._quote_identifier(part.strip("[] ")) for part in parts)

    def _validation_failures_table_name(self) -> str:
        parts = "dbo.RequestsValidationFailures".split(".")
        return ".".join(self._quote_identifier(part.strip("[] ")) for part in parts)

    def _column(self, name: str) -> str:
        return self._quote_identifier(name.strip("[] "))

    @staticmethod
    def _quote_identifier(identifier: str) -> str:
        if not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", identifier):
            raise ValueError(f"Unsafe SQL identifier configured: {identifier}")
        return f"[{identifier}]"

    def _where_clause(self, filters: RequestFilters, table_alias: str | None = None) -> tuple[str, list[Any]]:
        clauses: list[str] = []
        params: list[Any] = []
        prefix = f"{table_alias}." if table_alias else ""
        created_col = f"{prefix}{self._column(self.settings.created_at_column)}"
        request_col = f"{prefix}{self._column(self.settings.request_id_column)}"
        status_col = f"{prefix}{self._column(self.settings.status_column)}"

        if filters.date_from:
            clauses.append(f"{created_col} >= ?")
            params.append(datetime.combine(filters.date_from, datetime.min.time()))
        if filters.date_to:
            clauses.append(f"{created_col} < ?")
            params.append(datetime.combine(filters.date_to + timedelta(days=1), datetime.min.time()))
        if filters.request_id:
            clauses.append(f"CAST({request_col} AS NVARCHAR(255)) LIKE ?")
            params.append(f"%{filters.request_id.strip()}%")
        if filters.status_list:
            statuses = [value.strip() for value in filters.status_list if value and value.strip()]
            if statuses:
                placeholders = ", ".join(["?"] * len(statuses))
                clauses.append(f"{status_col} IN ({placeholders})")
                params.extend(statuses)
        elif filters.status:
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

    def export_rows_for_email(self, filters: RequestFilters) -> list[dict[str, Any]]:
        where, params = self._where_clause(filters, table_alias="c")
        table = self._table_name()
        submissions_table = self._submissions_table_name()
        outlets_table = self._outlets_table_name()
        validation_failures_table = self._validation_failures_table_name()
        created_col = self._column(self.settings.created_at_column)
        request_col = self._column(self.settings.request_id_column)
        sub_request_col = self._column("RequestId")
        sub_document_col = self._column("DocumentCount")
        sub_is_deleted_col = self._column("IsDeleted")
        sub_outlet_col = self._column("OutletId")
        sub_mode_col = self._column("Mode")
        sub_received_col = self._column("ReceivedAt")
        sub_processed_col = self._column("ProcessedAt")
        outlet_id_col = self._column("OutletId")
        outlet_name_col = self._column("Name")
        submission_id_col = self._column("Id")
        validation_request_col = self._column("RequestId")
        validation_error_col = self._column("ValidationError")
        validation_id_col = self._column("Id")

        submissions_subquery = (
            "SELECT SubRequestId, OutletId, Mode, DocumentCount FROM ("
            f"SELECT {sub_request_col} AS SubRequestId, {sub_outlet_col} AS OutletId, {sub_mode_col} AS Mode, "
            f"{sub_document_col} AS DocumentCount, "
            f"ROW_NUMBER() OVER (PARTITION BY {sub_request_col} ORDER BY COALESCE({sub_processed_col}, {sub_received_col}) DESC, {submission_id_col} DESC) AS RowNum "
            f"FROM {submissions_table} "
            f"WHERE {sub_is_deleted_col} = 0"
            ") AS ranked WHERE RowNum = 1"
        )

        validation_failures_subquery = (
            f"SELECT {validation_request_col} AS ValidationRequestId, "
            f"STRING_AGG(CAST({validation_error_col} AS NVARCHAR(MAX)), NCHAR(10)) "
            f"WITHIN GROUP (ORDER BY {validation_id_col}) AS ValidationError "
            f"FROM {validation_failures_table} "
            f"WHERE {validation_error_col} IS NOT NULL "
            f"GROUP BY {validation_request_col}"
        )

        query = (
            "SELECT c.*, s.DocumentCount, vf.ValidationError, "
            f"CASE WHEN UPPER(s.Mode) = 'EMAIL' THEN 'ATT' ELSE o.{outlet_name_col} END AS OutletName "
            f"FROM {table} AS c "
            f"LEFT JOIN ({submissions_subquery}) AS s ON s.SubRequestId = c.{request_col} "
            f"LEFT JOIN {outlets_table} AS o ON o.{outlet_id_col} = s.OutletId "
            f"LEFT JOIN ({validation_failures_subquery}) AS vf ON vf.ValidationRequestId = c.{request_col} "
            f"{where} "
            f"ORDER BY c.{created_col} DESC"
        )

        with self._connect() as conn:
            cursor = conn.cursor().execute(query, params)
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


class CachedRepository:
    def __init__(self, repository, ttl_seconds: int, max_entries: int):
        self.repository = repository
        self.ttl_seconds = max(0, int(ttl_seconds))
        self.max_entries = max(0, int(max_entries))
        self._lock = threading.Lock()
        self._cache: dict[tuple[Any, ...], tuple[float, dict[str, Any]]] = {}

    @staticmethod
    def _filter_key(filters: RequestFilters) -> tuple[Any, ...]:
        return (
            filters.request_id or "",
            filters.date_from.isoformat() if filters.date_from else "",
            filters.date_to.isoformat() if filters.date_to else "",
            filters.status or "",
            filters.q or "",
            filters.safe_page,
            filters.safe_page_size,
        )

    def _enabled(self) -> bool:
        return self.ttl_seconds > 0 and self.max_entries > 0

    def clear_cache(self) -> None:
        with self._lock:
            self._cache.clear()

    def list_requests(self, filters: RequestFilters) -> dict[str, Any]:
        if not self._enabled():
            return self.repository.list_requests(filters)

        key = self._filter_key(filters)
        now = time.monotonic()
        with self._lock:
            cached = self._cache.get(key)
            if cached:
                expires_at, value = cached
                if expires_at > now:
                    return copy.deepcopy(value)
                self._cache.pop(key, None)

        value = self.repository.list_requests(filters)
        with self._lock:
            self._cache[key] = (now + self.ttl_seconds, copy.deepcopy(value))
            while len(self._cache) > self.max_entries:
                self._cache.pop(next(iter(self._cache)))
        return value

    def export_rows(self, filters: RequestFilters) -> list[dict[str, Any]]:
        return self.repository.export_rows(filters)

    def export_rows_for_email(self, filters: RequestFilters) -> list[dict[str, Any]]:
        return self.repository.export_rows_for_email(filters)

    def retry_request(self, request_id: str) -> dict[str, Any] | None:
        row = self.repository.retry_request(request_id)
        if row:
            self.clear_cache()
        return row

# Updated to directly return the SqlServerRepository wrapped in cache
def get_repository(settings: Settings):
    repository = SqlServerRepository(settings)
    return CachedRepository(
        repository,
        ttl_seconds=settings.api_cache_ttl_seconds,
        max_entries=settings.api_cache_max_entries,
    )
