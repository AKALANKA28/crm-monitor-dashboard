from __future__ import annotations

import argparse
from datetime import date, datetime, timedelta
from pathlib import Path

from app.config import get_settings
from app.data_access import RequestFilters, get_repository
from app.email_report import DEFAULT_EMAIL_STATUSES, pending_as_failed_rows
from app.email_service import MicrosoftGraphEmailService
from app.excel_export import EMAIL_REQUEST_FIELDS, build_excel


def _sample_email_row() -> dict[str, str]:
    row = {field: f"Sample {field}" for field in EMAIL_REQUEST_FIELDS}
    row.update(
        {
            "Request ID": "TEST-EMAIL-001",
            "Email": "customer@example.com",
            "CRM_Email Id": "customer@example.com",
            "Mobile Number": "+971501234567",
            "CRM_Mobile No. *": "+971501234567",
            "Submitted ON": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Bot Status(Passed/Error)": "Passed",
            "Bot Error Comment": "",
            "Pass Status(Lead / Prospect)": "Lead",
            "CRM_Date of Birth": "1990-01-01",
            "CRM_Vehicle Value": "50000",
        }
    )
    return row


def _parse_date(value: str | None) -> date | None:
    if not value:
        return None
    return date.fromisoformat(value)


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


def build_sample_attachment() -> tuple[bytes, str]:
    settings = get_settings()
    filters = RequestFilters(date_from=date.today(), date_to=date.today())
    content = build_excel(
        [_sample_email_row()],
        filters,
        settings,
        export_profile="email",
        allow_missing_email_fields=True,
    )
    filename = f"Algospring_CRM_Email_Test_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return content, filename


def build_db_attachment(
    date_from: date | None,
    date_to: date | None,
    status_list: list[str] | None,
) -> tuple[bytes, str]:
    settings = get_settings()
    repository = get_repository(settings)

    if date_from is None and date_to is None:
        yesterday = date.today() - timedelta(days=1)
        date_from = yesterday
        date_to = yesterday
    elif date_from is None:
        date_from = date_to
    elif date_to is None:
        date_to = date_from

    filters = RequestFilters(
        date_from=date_from,
        date_to=date_to,
        page=1,
        page_size=200,
    )
    if status_list:
        filters.status_list = status_list

    rows = repository.export_rows_for_email(filters)
    report_rows, pending_as_failed_count = pending_as_failed_rows(rows, settings.status_column)
    print(f"Rows fetched: {len(rows)}")
    print(f"Date filter column: {settings.created_at_column}")
    print(f"Pending rows included as Failed: {pending_as_failed_count}")
    content = build_excel(
        report_rows,
        filters,
        settings,
        export_profile="email",
        allow_missing_email_fields=True,
    )
    filename = f"Algospring_CRM_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return content, filename


def main() -> None:
    parser = argparse.ArgumentParser(description="Build and optionally send a test CRM email report.")
    parser.add_argument("--to", help="Recipient email address.")
    parser.add_argument("--send", action="store_true", help="Actually send the email via Outlook.")
    parser.add_argument(
        "--out",
        default="tmp_email_test_report.xlsx",
        help="Local dry-run output path for the generated attachment.",
    )
    parser.add_argument("--date-from", help="Filter start date (YYYY-MM-DD). Defaults to yesterday.")
    parser.add_argument("--date-to", help="Filter end date (YYYY-MM-DD). Defaults to yesterday.")
    parser.add_argument(
        "--status",
        help="Comma-separated statuses (default: Success,Failed,Pending).",
    )
    parser.add_argument(
        "--recipients-file",
        help="Path to recipients file (defaults to EMAIL_RECIPIENTS_FILE).",
    )
    parser.add_argument("--sample", action="store_true", help="Use a sample row instead of DB data.")
    args = parser.parse_args()

    report_date: str
    if args.sample:
        content, filename = build_sample_attachment()
        report_date = datetime.now().strftime("%b %d, %Y")
    else:
        date_from = _parse_date(args.date_from)
        date_to = _parse_date(args.date_to)
        if date_from is None and date_to is None:
            yesterday = date.today() - timedelta(days=1)
            date_from = yesterday
            date_to = yesterday
        elif date_from is None:
            date_from = date_to
        elif date_to is None:
            date_to = date_from
        report_date = (date_from or date.today()).strftime("%b %d, %Y")

        if args.status:
            status_list = [value.strip() for value in args.status.split(",") if value.strip()]
        else:
            status_list = DEFAULT_EMAIL_STATUSES

        print(f"Server date: {date.today().isoformat()}")
        print(f"Report date range: {date_from.isoformat()} to {date_to.isoformat()}")

        content, filename = build_db_attachment(date_from, date_to, status_list)

    output_path = Path(args.out)
    output_path.write_bytes(content)
    print(f"Built test attachment: {output_path.resolve()}")

    if not args.send:
        print("Dry run only. Add --send --to recipient@example.com to send the email.")
        return

    settings = get_settings()
    if not settings.graph_client_id:
        raise SystemExit("GRAPH_CLIENT_ID is missing in .env")

    recipients_file = Path(args.recipients_file or settings.email_recipients_file)
    if args.to:
        to_emails = _split_emails(args.to)
        cc_emails: list[str] = []
        bcc_emails: list[str] = []
    else:
        recipients = _load_email_recipients(recipients_file)
        to_emails = recipients["to"]
        cc_emails = recipients["cc"]
        bcc_emails = recipients["bcc"]
        if not to_emails:
            raise SystemExit(
                f"Recipient list is empty. Provide --to or update {recipients_file}."
            )

    service = MicrosoftGraphEmailService(settings)
    service.send_report_email(
        to_email=to_emails,
        cc_emails=cc_emails,
        bcc_emails=bcc_emails,
        subject=f"CRM Request Status Report - {report_date}",
        body=(
            "Dear Team,\n\n"
            f"Please find attached the CRM Request Status Report for {report_date}. "
            "This report includes Success and Failed requests for the specified date. "
            "Pending requests are included under Failed.\n\n"
            "If you require additional fields or a different date range, please let us know.\n\n"
            "Sincerely,\n"
            "Algospring Automated System"
        ),
        file_bytes=content,
        filename=filename,
        allow_device_flow=True,
    )
    print(f"Sent test email to {args.to}")


if __name__ == "__main__":
    main()
