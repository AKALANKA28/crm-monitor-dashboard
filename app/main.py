from __future__ import annotations

from datetime import date, datetime
from io import BytesIO
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, HTTPException, Query, Request
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from app.config import get_settings
from app.data_access import RequestFilters, get_repository
from app.excel_export import build_excel

settings = get_settings()
repository = get_repository(settings)

BASE_DIR = Path(__file__).resolve().parent

app = FastAPI(title=settings.app_title)
app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))


def build_filters(
    request_id: Optional[str] = None,
    date_from: Optional[date] = None,
    date_to: Optional[date] = None,
    status: Optional[str] = None,
    q: Optional[str] = None,
    page: int = 1,
    page_size: int = 25,
) -> RequestFilters:
    return RequestFilters(
        request_id=request_id or None,
        date_from=date_from,
        date_to=date_to,
        status=status or None,
        q=q or None,
        page=page,
        page_size=page_size,
    )


@app.get("/", response_class=HTMLResponse)
def dashboard(request: Request):
    return templates.TemplateResponse(
        "index.html",
        {
            "request": request,
            "app_title": settings.app_title,
            "data_source": settings.data_source,
        },
    )


@app.get("/api/requests")
def api_requests(
    request_id: Optional[str] = Query(None),
    date_from: Optional[date] = Query(None),
    date_to: Optional[date] = Query(None),
    status: Optional[str] = Query(None),
    q: Optional[str] = Query(None),
    page: int = Query(1, ge=1),
    page_size: int = Query(25, ge=1, le=200),
):
    filters = build_filters(request_id, date_from, date_to, status, q, page, page_size)
    return repository.list_requests(filters)


@app.post("/api/requests/{request_id}/retry")
def retry_request(request_id: str):
    row = repository.retry_request(request_id)
    if not row:
        raise HTTPException(status_code=404, detail="Request ID not found")
    return {"message": "Request moved back to Pending", "row": row}


@app.get("/api/export")
def export_excel(
    request_id: Optional[str] = Query(None),
    date_from: Optional[date] = Query(None),
    date_to: Optional[date] = Query(None),
    status: Optional[str] = Query(None),
    q: Optional[str] = Query(None),
):
    filters = build_filters(request_id, date_from, date_to, status, q, page=1, page_size=200)
    rows = repository.export_rows(filters)
    content = build_excel(rows, filters, settings)
    filename = f"crm_requests_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    headers = {"Content-Disposition": f"attachment; filename={filename}"}
    return StreamingResponse(
        BytesIO(content),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )
