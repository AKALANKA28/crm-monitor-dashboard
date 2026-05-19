from __future__ import annotations

from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

try:
    from openpyxl.drawing.image import Image as ExcelImage
except ImportError:
    ExcelImage = None

from app.config import Settings
from app.data_access import RequestFilters, STATUS_ORDER, _normalize_status_counts

# --- NEW MODERN COLOR PALETTE ---
BRAND_COLORS = {
    "slate": "1E293B",       # Main dark headers
    "lime": "CCFF00",        # Accent
    "white": "FFFFFF",       # Pure white for all backgrounds
    "panel": "FFFFFF",       # Pure white for cards/tables
    "border": "E2E8F0",      # Light modern border
    "text": "0F172A",        # Main text
    "muted": "64748B",       # Subtitle text
    
    # Status Colors (Text)
    "success": "15803D",
    "failed": "B91C1C",
    "pending": "B45309",
    "progress": "4338CA",
    
    # Status Colors (Backgrounds)
    "success_bg": "DCFCE7",
    "failed_bg": "FEE2E2",
    "pending_bg": "FEF3C7",
    "progress_bg": "E0E7FF",
}

STATUS_STYLES = {
    "Success": (BRAND_COLORS["success_bg"], BRAND_COLORS["success"]),
    "Failed": (BRAND_COLORS["failed_bg"], BRAND_COLORS["failed"]),
    "Pending": (BRAND_COLORS["pending_bg"], BRAND_COLORS["pending"]),
    "In Progress": (BRAND_COLORS["progress_bg"], BRAND_COLORS["progress"]),
}


def build_excel(rows: list[dict[str, Any]], filters: RequestFilters, settings: Settings) -> bytes:
    wb = Workbook()
    summary = wb.active
    summary.title = "Summary"

    counts = _status_counts(rows, settings)
    _build_summary_sheet(summary, rows, filters, counts, settings)

    requests = wb.create_sheet(_safe_sheet_name("Requests"))
    _build_requests_sheet(requests, rows, settings)

    stream = BytesIO()
    wb.save(stream)
    return stream.getvalue()


def _build_summary_sheet(ws, rows: list[dict[str, Any]], filters: RequestFilters, counts: dict[str, int], settings: Settings) -> None:
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90
    
    # Set a pure white background for the whole summary sheet
    _style_cells(ws, "A1:M40", fill=BRAND_COLORS["white"])

    for column, width in {
        "A": 15, "B": 15, "C": 15, "D": 15, "E": 15, 
        "F": 15, "G": 15, "H": 15, "I": 15, "J": 15,
    }.items():
        ws.column_dimensions[column].width = width

    for row_num in range(1, 32):
        ws.row_dimensions[row_num].height = 24

    _build_header(ws, settings)
    _build_kpi_cards(ws, len(rows), counts)
    _build_filters_table(ws, filters)
    _build_status_table(ws, counts)
    _build_footer(ws)


def _build_header(ws, settings: Settings) -> None:
    # Header Area (White background)
    _style_cells(ws, "A1:J5", fill=BRAND_COLORS["white"], border=BRAND_COLORS["border"])
    ws.merge_cells("A1:B5")
    ws.merge_cells("C1:J2")
    ws.merge_cells("C3:J3")
    ws.merge_cells("C4:J4")

    _add_logo(ws, settings, "A1")

    ws["C1"] = "CRM Request Monitor Export"
    ws["C1"].font = Font(name="Inter", size=22, bold=True, color=BRAND_COLORS["slate"])
    ws["C1"].alignment = Alignment(vertical="center")

    ws["C3"] = f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    ws["C3"].font = Font(name="Inter", size=11, color=BRAND_COLORS["muted"])
    ws["C3"].alignment = Alignment(vertical="center")


def _build_kpi_cards(ws, total: int, counts: dict[str, int]) -> None:
    cards = [
        ("A7:B10", "Total Requests", total, BRAND_COLORS["slate"]),
        ("C7:D10", "Success", counts.get("Success", 0), BRAND_COLORS["success"]),
        ("E7:F10", "Failed", counts.get("Failed", 0), BRAND_COLORS["failed"]),
        ("G7:H10", "Pending", counts.get("Pending", 0), BRAND_COLORS["pending"]),
        ("I7:J10", "In Progress", counts.get("In Progress", 0), BRAND_COLORS["progress"]),
    ]

    for cell_range, label, value, text_color in cards:
        start_cell = cell_range.split(":")[0]
        col = "".join(filter(str.isalpha, start_cell))
        row = int("".join(filter(str.isdigit, start_cell)))

        # Clean white card with modern border
        _style_cells(ws, cell_range, fill=BRAND_COLORS["panel"], border=BRAND_COLORS["border"])
        ws.merge_cells(start_row=row, start_column=ws[col + str(row)].column, end_row=row + 1, end_column=ws[col + str(row)].column + 1)
        ws.merge_cells(start_row=row + 2, start_column=ws[col + str(row)].column, end_row=row + 3, end_column=ws[col + str(row)].column + 1)

        ws[f"{col}{row}"] = label.upper()
        ws[f"{col}{row}"].font = Font(name="Inter", size=10, bold=True, color=BRAND_COLORS["muted"])
        ws[f"{col}{row}"].alignment = Alignment(horizontal="center", vertical="center")

        ws[f"{col}{row + 2}"] = value
        ws[f"{col}{row + 2}"].font = Font(name="Inter", size=26, bold=True, color=text_color)
        ws[f"{col}{row + 2}"].alignment = Alignment(horizontal="center", vertical="center")


def _build_filters_table(ws, filters: RequestFilters) -> None:
    _section_title(ws, "A12:D12", "Applied Filters")
    rows = [
        ("Request ID", filters.request_id or "All"),
        ("Date From", filters.date_from.isoformat() if filters.date_from else "All"),
        ("Date To", filters.date_to.isoformat() if filters.date_to else "All"),
        ("Status", filters.status or "All"),
        ("Search", filters.q or "All"),
    ]
    for row_num, (label, value) in enumerate(rows, start=13):
        ws[f"A{row_num}"] = label
        ws[f"B{row_num}"] = value
        ws.merge_cells(f"B{row_num}:D{row_num}")
        
        _style_cells(ws, f"A{row_num}:D{row_num}", fill=BRAND_COLORS["panel"], border=BRAND_COLORS["border"])
        ws[f"A{row_num}"].font = Font(name="Inter", bold=True, color=BRAND_COLORS["text"])
        ws[f"B{row_num}"].font = Font(name="Inter", color=BRAND_COLORS["text"])
        ws[f"B{row_num}"].alignment = Alignment(horizontal="right")


def _build_status_table(ws, counts: dict[str, int]) -> None:
    _section_title(ws, "F12:J12", "Status Breakdown")
    ws["F13"] = "STATUS"
    ws["I13"] = "COUNT"
    ws.merge_cells("F13:H13")
    ws.merge_cells("I13:J13")
    
    _style_cells(ws, "F13:J13", fill=BRAND_COLORS["white"], border=BRAND_COLORS["border"])
    for cell_ref in ["F13", "I13"]:
        ws[cell_ref].font = Font(name="Inter", size=10, bold=True, color=BRAND_COLORS["muted"])
        ws[cell_ref].alignment = Alignment(horizontal="left" if cell_ref == "F13" else "right", vertical="center")

    for row_num, status in enumerate(STATUS_ORDER, start=14):
        ws[f"F{row_num}"] = status
        ws[f"I{row_num}"] = counts.get(status, 0)
        ws.merge_cells(f"F{row_num}:H{row_num}")
        ws.merge_cells(f"I{row_num}:J{row_num}")
        
        _style_cells(ws, f"F{row_num}:J{row_num}", fill=BRAND_COLORS["panel"], border=BRAND_COLORS["border"])
        
        # Color code the status text
        _, font_color = STATUS_STYLES.get(status, (None, BRAND_COLORS["text"]))
        ws[f"F{row_num}"].font = Font(name="Inter", bold=True, color=font_color)
        ws[f"I{row_num}"].font = Font(name="Inter", bold=True, color=BRAND_COLORS["text"])
        ws[f"I{row_num}"].alignment = Alignment(horizontal="right")

    next_row = 14 + len(STATUS_ORDER)
    for status, count in counts.items():
        if status in STATUS_ORDER:
            continue
        ws[f"F{next_row}"] = status
        ws[f"I{next_row}"] = count
        ws.merge_cells(f"F{next_row}:H{next_row}")
        ws.merge_cells(f"I{next_row}:J{next_row}")
        _style_cells(ws, f"F{next_row}:J{next_row}", fill=BRAND_COLORS["panel"], border=BRAND_COLORS["border"])
        ws[f"I{next_row}"].alignment = Alignment(horizontal="right")
        next_row += 1


def _build_footer(ws) -> None:
    ws.merge_cells("A22:J22")
    ws["A22"] = "Prepared by Earnest Insurance | CRM Monitor"
    ws["A22"].font = Font(name="Inter", size=10, color=BRAND_COLORS["muted"], italic=True)
    ws["A22"].alignment = Alignment(horizontal="center")


def _build_requests_sheet(ws, rows: list[dict[str, Any]], settings: Settings) -> None:
    columns = list(rows[0].keys()) if rows else ["No data matched the selected filters"]
    
    # Write headers
    ws.append([str(c).upper() for c in columns])
    
    # Write data
    for row in rows:
        ws.append([row.get(column) for column in columns])

    status_col_idx = _status_column_index(columns, settings)
    _apply_requests_table_style(ws, max(1, ws.max_row), max(1, ws.max_column), status_col_idx)


def _apply_requests_table_style(ws, max_row: int, max_col: int, status_col_idx: int | None) -> None:
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(max_col)}{max_row}"
    ws.row_dimensions[1].height = 24

    thin = Side(style="thin", color=BRAND_COLORS["border"])
    
    # Header Style (Dark Slate)
    for cell in ws[1]:
        cell.fill = _fill(BRAND_COLORS["slate"])
        cell.font = Font(name="Inter", size=10, color=BRAND_COLORS["white"], bold=True)
        cell.alignment = Alignment(horizontal="left", vertical="center")
        cell.border = Border(bottom=thin)

    # Data Rows (Pure White Background)
    row_fill_white = _fill(BRAND_COLORS["white"])
    
    for row in ws.iter_rows(min_row=2, max_row=max_row, max_col=max_col):
        for cell in row:
            cell.fill = row_fill_white
            cell.border = Border(bottom=thin, top=thin, left=thin, right=thin)
            cell.alignment = Alignment(vertical="center")
            cell.font = Font(name="Inter", color=BRAND_COLORS["text"])
            
        # Apply Status Colors if applicable
        if status_col_idx:
            status_cell = row[status_col_idx - 1]
            fill, font_color = STATUS_STYLES.get(_canonical_status(status_cell.value), (None, None))
            if fill and font_color:
                status_cell.fill = _fill(fill)
                status_cell.font = Font(name="Inter", bold=True, color=font_color)
                status_cell.alignment = Alignment(horizontal="center", vertical="center")

    # Auto-adjust column widths
    for col_idx in range(1, max_col + 1):
        letter = get_column_letter(col_idx)
        max_length = 12
        for cell in ws[letter]:
            max_length = max(max_length, min(len(str(cell.value or "")), 45))
        ws.column_dimensions[letter].width = max_length + 3


def _status_counts(rows: list[dict[str, Any]], settings: Settings) -> dict[str, int]:
    counts: dict[str, int] = {}
    for row in rows:
        status = _status_value(row, settings)
        counts[status] = counts.get(status, 0) + 1
    return _normalize_status_counts(counts)


def _status_value(row: dict[str, Any], settings: Settings) -> str:
    for key in [settings.status_column, "CRMStatus", "crmStatus", "crm_status", "Status", "status"]:
        if key in row:
            return _canonical_status(row.get(key))
    return "Unknown"


def _canonical_status(value: Any) -> str:
    text = str(value or "").strip()
    if not text:
        return "Unknown"
    normalized = text.casefold().replace(" ", "").replace("_", "").replace("-", "")
    mapping = {
        "success": "Success",
        "failed": "Failed",
        "fail": "Failed",
        "failure": "Failed",
        "error": "Failed",
        "pending": "Pending",
        "inprogress": "In Progress",
        "processing": "In Progress",
    }
    return mapping.get(normalized, text)


def _status_column_index(columns: list[str], settings: Settings) -> int | None:
    candidates = {settings.status_column, "CRMStatus", "crmStatus", "crm_status", "Status", "status"}
    for index, column in enumerate(columns, start=1):
        if column in candidates:
            return index
    return None


def _add_logo(ws, settings: Settings, cell: str) -> None:
    if ExcelImage is None:
        return
    logo_path = _resolve_logo_path(settings)
    if not logo_path.exists():
        return
    try:
        logo = ExcelImage(str(logo_path))
    except Exception:
        return
    logo.width = 140  
    logo.height = 70
    ws.add_image(logo, cell)


def _resolve_logo_path(settings: Settings) -> Path:
    path = Path(settings.excel_logo_path)
    if path.is_absolute():
        return path
    return Path(__file__).resolve().parent.parent / path


def _section_title(ws, cell_range: str, title: str) -> None:
    start_cell = cell_range.split(":")[0]
    ws.merge_cells(cell_range)
    ws[start_cell] = title.upper()
    ws[start_cell].font = Font(name="Inter", size=11, bold=True, color=BRAND_COLORS["slate"])
    ws[start_cell].alignment = Alignment(horizontal="left", vertical="bottom")


def _style_cells(ws, cell_range: str, fill: str | None = None, border: str | None = None) -> None:
    for row in ws[cell_range]:
        for cell in row:
            if fill:
                cell.fill = _fill(fill)
            if border:
                cell.border = _border(border)
            cell.alignment = Alignment(vertical="center")


def _fill(color: str) -> PatternFill:
    return PatternFill(fill_type="solid", fgColor=color)


def _border(color: str) -> Border:
    side = Side(style="thin", color=color)
    return Border(left=side, right=side, top=side, bottom=side)


def _safe_sheet_name(value: str) -> str:
    return value[:31].replace("/", "-").replace("\\", "-").replace("?", "")