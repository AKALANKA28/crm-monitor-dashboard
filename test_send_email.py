from __future__ import annotations

import argparse
from datetime import date, datetime, timedelta
from pathlib import Path

from app.config import get_settings
from app.data_access import RequestFilters, get_repository
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
    filename = f"Earnest_CRM_Email_Test_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
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

    filters = RequestFilters(
        date_from=date_from,
        date_to=date_to,
        page=1,
        page_size=200,
    )
    if status_list:
        filters.status_list = status_list

    rows = repository.export_rows(filters)
    content = build_excel(
        rows,
        filters,
        settings,
        export_profile="email",
        allow_missing_email_fields=True,
    )
    filename = f"Earnest_CRM_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
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
        help="Comma-separated statuses (default: Success,Failed).",
    )
    parser.add_argument("--sample", action="store_true", help="Use a sample row instead of DB data.")
    args = parser.parse_args()

    if args.sample:
        content, filename = build_sample_attachment()
    else:
        date_from = _parse_date(args.date_from)
        date_to = _parse_date(args.date_to)
        status_list = None
        if args.status:
            status_list = [value.strip() for value in args.status.split(",") if value.strip()]
        else:
            status_list = ["Success", "Failed"]
        content, filename = build_db_attachment(date_from, date_to, status_list)
    output_path = Path(args.out)
    output_path.write_bytes(content)
    print(f"Built test attachment: {output_path.resolve()}")

    if not args.send:
        print("Dry run only. Add --send --to recipient@example.com to send the email.")
        return

    if not args.to:
        raise SystemExit("--to is required when using --send")

    settings = get_settings()
    if not settings.graph_client_id:
        raise SystemExit("GRAPH_CLIENT_ID is missing in .env")

    service = MicrosoftGraphEmailService(settings)
    service.send_report_email(
        to_email=args.to,
        subject=f"Earnest CRM Email Test - {datetime.now().strftime('%b %d, %Y')}",
        body=(
            "Hello,\n\n"
            "This is a test email for the Earnest CRM report attachment.\n\n"
            "Thank you,\n"
            "Earnest Automated Systems"
        ),
        file_bytes=content,
        filename=filename,
        allow_device_flow=True,
    )
    print(f"Sent test email to {args.to}")


if __name__ == "__main__":
    main()
