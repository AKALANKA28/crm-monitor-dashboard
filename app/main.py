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

RECIPIENT_KEY_MAP = {
    "recipient_emails": "to",
    "cc_emails": "cc",
    "bcc_emails": "bcc",
}

def _split_emails(value: str) -> list[str]:
    return [item.strip() for item in value.replace(";", ",").split(",") if item.strip()]

def _load_email_recipients(path: Path) -> dict[str, list[str]]:
    recipients = {"to": [], "cc": [], "bcc": []}
    if not path.exists():
        return recipients
    for line in path.read_text(encoding="utf-8").splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("#") or stripped.startswith(";"):
            continue
        if "=" not in stripped:
            continue
        key, value = stripped.split("=", 1)
        normalized_key = key.strip().lower()
        target = RECIPIENT_KEY_MAP.get(normalized_key)
        if not target:
            continue
        recipients[target].extend(_split_emails(value))
    return recipients

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
    email: Optional[str] = Query(None, description="Email address(es) to send the report to"),
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
        rows = repository.export_rows_for_email(filters)
        content = build_excel(
            rows,
            filters,
            settings,
            export_profile="email",
            allow_missing_email_fields=True,
        )
        
        # 2. Setup Email Meta
        report_date = date_from.strftime("%b %d, %Y") if date_from else datetime.now().strftime("%b %d, %Y")
        filename = f"Algospring_CRM_Report_{datetime.now().strftime('%Y%m%d')}.xlsx"
        subject = f"CRM Request Status Report - {report_date}"
        body = (
            "Dear Amit,\n\n"
            f"Please find attached the CRM Request Status Report for {report_date}. "
            "This report includes Success and Failed requests for the specified date.\n\n"
            "If you require additional fields or a different date range, please let us know.\n\n"
            "Sincerely,\n"
            "Algospring Automated System"
        )
        
        # 3. Resolve recipients
        to_emails: list[str]
        cc_emails: list[str]
        bcc_emails: list[str]
        if email:
            to_emails = _split_emails(email)
            cc_emails = []
            bcc_emails = []
        else:
            recipients = _load_email_recipients(Path(settings.email_recipients_file))
            to_emails = recipients["to"]
            cc_emails = recipients["cc"]
            bcc_emails = recipients["bcc"]

        if not to_emails:
            raise HTTPException(
                status_code=400,
                detail=(
                    "Recipient list is empty. Provide 'email' or update "
                    f"{settings.email_recipients_file}."
                ),
            )

        # 4. Send Email via Microsoft Graph
        email_service = MicrosoftGraphEmailService(settings)
        email_service.send_report_email(
            to_email=to_emails,
            cc_emails=cc_emails,
            bcc_emails=bcc_emails,
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
