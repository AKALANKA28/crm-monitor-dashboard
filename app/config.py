from functools import lru_cache
from pathlib import Path
from pydantic import BaseModel
from dotenv import load_dotenv
import os

ROOT_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT_DIR / "app" / "data"
DUMMY_DATA_FILE = DATA_DIR / "crm_requests.json"

load_dotenv(ROOT_DIR / ".env")


class Settings(BaseModel):
    app_title: str = os.getenv("APP_TITLE", "CRM Request Monitor")
    data_source: str = os.getenv("APP_DATA_SOURCE", "dummy").lower()

    db_table: str = os.getenv("DB_TABLE", "dbo.Customers")
    created_at_column: str = os.getenv("DB_CREATED_AT_COLUMN", "CreatedAt")
    status_column: str = os.getenv("DB_STATUS_COLUMN", "CRMStatus")
    request_id_column: str = os.getenv("DB_REQUEST_ID_COLUMN", "RequestId")

    sqlserver_connection_string: str | None = os.getenv("SQLSERVER_CONNECTION_STRING")
    sqlserver_driver: str = os.getenv("SQLSERVER_DRIVER", "ODBC Driver 18 for SQL Server")
    sqlserver_server: str = os.getenv("SQLSERVER_SERVER", "localhost")
    sqlserver_database: str = os.getenv("SQLSERVER_DATABASE", "YourDatabase")
    sqlserver_username: str = os.getenv("SQLSERVER_USERNAME", "")
    sqlserver_password: str = os.getenv("SQLSERVER_PASSWORD", "")
    sqlserver_trust_certificate: str = os.getenv("SQLSERVER_TRUST_CERTIFICATE", "yes")

    @property
    def sql_connection_string(self) -> str:
        if self.sqlserver_connection_string:
            return self.sqlserver_connection_string
        return (
            f"DRIVER={{{self.sqlserver_driver}}};"
            f"SERVER={self.sqlserver_server};"
            f"DATABASE={self.sqlserver_database};"
            f"UID={self.sqlserver_username};"
            f"PWD={self.sqlserver_password};"
            f"TrustServerCertificate={self.sqlserver_trust_certificate};"
        )


@lru_cache
def get_settings() -> Settings:
    return Settings()
