# CRM Request Monitor Dashboard

A simple, professional monitoring web dashboard for CRM request statuses. It starts with dummy JSON data, then can be switched to Azure SQL once database credentials are ready.

## Features

- Daily and date-range request monitoring
- Summary cards for Total, Success, Failed, Pending, and In Progress
- Filters by Request ID, Date From, Date To, CRM Status, and Search text
- Dynamic table rendering for wide tables with 30 to 40+ columns
- Retry action for Failed and In Progress rows, which sets `CRMStatus` back to `Pending`
- Excel export based on the active filters
- Dummy data included/generated for local testing
- Azure SQL-ready repository layer

## Project structure

```text
crm-monitor-dashboard/
├── app/
│   ├── data/                  # dummy JSON data is generated here on first run
│   ├── static/
│   │   ├── app.js
│   │   └── styles.css
│   ├── templates/
│   │   └── index.html
│   ├── config.py
│   ├── data_access.py
│   ├── dummy_data.py
│   └── main.py
├── .env.example
├── .gitignore
├── requirements.txt
├── requirements-mssql.txt
└── README.md
```

## Run locally with dummy data

```bash
cd crm-monitor-dashboard
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
cp .env.example .env
uvicorn app.main:app --reload
```

Open the dashboard:

```text
http://127.0.0.1:8000
```

The app generates dummy rows in `app/data/crm_requests.json` on first run.

## Switch to Azure SQL later

Install the Azure SQL extras:

```bash
pip install -r requirements-mssql.txt
```

Then update `.env`:

```env
APP_DATA_SOURCE=azure
DB_CONNECTION_STRING=DRIVER={ODBC Driver 18 for SQL Server};SERVER=your-server.database.windows.net;DATABASE=your-db;UID=your-user;PWD=your-password;Encrypt=yes;TrustServerCertificate=no;Connection Timeout=30;
DB_TABLE=dbo.Customers
DB_CREATED_AT_COLUMN=CreatedAt
DB_STATUS_COLUMN=CRMStatus
DB_REQUEST_ID_COLUMN=RequestId
```

You can also use the individual `DB_*` values in `.env.example` instead of `DB_CONNECTION_STRING`.

## Dashboard read cache

The `/api/requests` endpoint uses a small in-memory TTL cache so repeated auto-refreshes with the same filters do not hit the database every time. By default, identical dashboard reads are cached for 20 seconds and the cache is cleared after a retry action.

```env
API_CACHE_TTL_SECONDS=20
API_CACHE_MAX_ENTRIES=256
```

Set `API_CACHE_TTL_SECONDS=0` to disable caching. Exports are not cached, so downloaded Excel files always query the current filtered data.

## SQL note for daily counts

Your current query works:

```sql
SELECT COUNT(*) AS SuccessCount
FROM [dbo].[Customers]
WHERE CAST(CreatedAt AS DATE) = '2026-05-18';
```

For production performance, especially when `CreatedAt` is indexed, prefer a date range because it avoids casting the column:

```sql
SELECT CRMStatus, COUNT(*) AS StatusCount
FROM [dbo].[Customers]
WHERE CreatedAt >= '2026-05-18'
  AND CreatedAt < DATEADD(DAY, 1, '2026-05-18')
GROUP BY CRMStatus;
```

The dashboard repository uses this range-based approach for Azure SQL mode.

## Retry action

In dummy mode, clicking **Retry** updates the JSON row:

```text
CRMStatus = Pending
UpdatedAt = current timestamp
Attempts = Attempts + 1
LastError = Moved back to Pending for retry.
```

In Azure SQL mode, the app runs an update similar to:

```sql
UPDATE dbo.Customers
SET CRMStatus = 'Pending', UpdatedAt = GETDATE()
WHERE RequestId = @requestId;
```

If your actual table does not have `UpdatedAt`, either add it or adjust `SqlServerRepository.retry_request()` in `app/data_access.py`.

## Customer-facing Excel branding

The Excel export uses the Earnest Insurance color palette and can include the company logo in the Summary sheet header.

Excel layout, color palette, logo placement, summary cards, charts, and request table styling are managed in:

```text
app/excel_export.py
```

Place the logo image here:

```text
app/static/earnest-logo.png
```

Or override the path in `.env`:

```env
EXCEL_LOGO_PATH=app/static/earnest-logo.png
```

Logo embedding requires Pillow, which is included in `requirements.txt`.

## API endpoints

```text
GET  /                         Dashboard UI
GET  /api/requests             Filtered paginated JSON data
POST /api/requests/{id}/retry  Move selected request back to Pending
GET  /api/export               Download filtered rows as Excel
```

## Customizing for your real table

Most changes are done through `.env`:

```env
DB_TABLE=dbo.Customers
DB_CREATED_AT_COLUMN=CreatedAt
DB_STATUS_COLUMN=CRMStatus
DB_REQUEST_ID_COLUMN=RequestId
```

Because your table may have 30 to 40+ columns, the dashboard reads `SELECT *` and renders columns dynamically. For a very large production table, consider limiting the visible columns in SQL mode and keeping full details only in the Excel export.
