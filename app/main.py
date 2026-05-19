from __future__ import annotations
from datetime import date, datetime, timedelta
from io import BytesIO
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, HTTPException, Query, Request
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from app.config import get_settings
from app.data_access import RequestFilters, get_repository
from app.excel_export import MissingEmailFieldsError, build_excel
from app.email_service import MicrosoftGraphEmailService

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

# --- NEW ENDPOINT: SEND EMAIL ---
@app.post("/api/send-email")
def send_email_report(
    email: str = Query(..., description="Email address to send the report to"),
    request_id: Optional[str] = Query(None),
    date_from: Optional[date] = Query(None),
    date_to: Optional[date] = Query(None),
    status: Optional[str] = Query(None),
    q: Optional[str] = Query(None),
):
    try:
        if not settings.graph_client_id:
            raise HTTPException(
                status_code=400,
                detail="GRAPH_CLIENT_ID is required before testing email sending.",
            )

        if date_from is None and date_to is None:
            yesterday = date.today() - timedelta(days=1)
            date_from = yesterday
            date_to = yesterday

        status_list: list[str] | None = None
        if status:
            status_list = [value.strip() for value in status.split(",") if value.strip()]
        else:
            status_list = ["Success", "Failed"]

        # 1. Generate the Excel File
        filters = build_filters(request_id, date_from, date_to, None, q, page=1, page_size=200)
        filters.status_list = status_list
        rows = repository.export_rows(filters)
        content = build_excel(
            rows,
            filters,
            settings,
            export_profile="email",
            allow_missing_email_fields=True,
        )
        
        # 2. Setup Email Meta
        filename = f"Earnest_CRM_Report_{datetime.now().strftime('%Y%m%d')}.xlsx"
        subject = f"Earnest CRM Report - {datetime.now().strftime('%b %d, %Y')}"
        body = "Hello,\n\nPlease find the requested CRM Request Status Report attached.\n\nThank you,\nEarnest Automated Systems"
        
        # 3. Send Email via Microsoft Graph
        email_service = MicrosoftGraphEmailService(settings)
        email_service.send_report_email(
            to_email=email,
            subject=subject,
            body=body,
            file_bytes=content,
            filename=filename
        )
        
        return {"message": f"Email successfully sent to {email}"}
        
    except HTTPException:
        raise
    except MissingEmailFieldsError as e:
        raise HTTPException(
            status_code=422,
            detail={
                "message": "Some required email template fields could not be matched. Send me the matching database column names for these fields.",
                "missing_fields": e.missing_fields,
                "available_columns": e.available_columns,
            },
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
