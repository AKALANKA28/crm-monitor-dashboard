from functools import lru_cache
from pathlib import Path
from pydantic import BaseModel
from dotenv import load_dotenv
import os

ROOT_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT_DIR / "app" / "data"

load_dotenv(ROOT_DIR / ".env", interpolate=False)

def _env_int(name: str, default: int) -> int:
    value = os.getenv(name)
    if value is None or value.strip() == "":
        return default
    return int(value)

def _env_list(name: str, default: str) -> list[str]:
    value = os.getenv(name, default)
    return [item for item in value.replace(",", " ").split() if item]

def _graph_scopes() -> list[str]:
    reserved = {"openid", "profile", "offline_access", "email"}
    scopes = _env_list(
        "GRAPH_SCOPES",
        os.getenv("OUTLOOK_SCOPES", "https://graph.microsoft.com/Mail.Send"),
    )
    return [scope for scope in scopes if scope.casefold() not in reserved]

class Settings(BaseModel):
    app_title: str = os.getenv("APP_TITLE", "CRM Request Monitor")
    data_source: str = os.getenv("APP_DATA_SOURCE", "dummy").lower()

    db_table: str = os.getenv("DB_TABLE", "dbo.Customers")
    created_at_column: str = os.getenv("DB_CREATED_AT_COLUMN", "CreatedAt")
    status_column: str = os.getenv("DB_STATUS_COLUMN", "CRMStatus")
    request_id_column: str = os.getenv("DB_REQUEST_ID_COLUMN", "RequestId")
    excel_logo_path: str = os.getenv("EXCEL_LOGO_PATH", "app/static/earnest-logo.png")

    # --- SQL DATABASE SETTINGS ---
    db_connection_string: str | None = os.getenv("DB_CONNECTION_STRING")
    db_driver: str = os.getenv("DB_DRIVER", "ODBC Driver 18 for SQL Server")
    db_server: str = os.getenv("DB_SERVER", "localhost")
    db_database: str = os.getenv("DB_DATABASE", "YourDatabase")
    db_username: str = os.getenv("DB_USERNAME", "")
    db_password: str = os.getenv("DB_PASSWORD", "")
    db_encrypt: str = os.getenv("DB_ENCRYPT", "yes")
    db_trust_certificate: str = os.getenv("DB_TRUST_CERTIFICATE", "no")
    db_connection_timeout: str = os.getenv("DB_CONNECTION_TIMEOUT", "30")

    # --- CACHING SETTINGS (Restored) ---
    api_cache_ttl_seconds: int = _env_int("API_CACHE_TTL_SECONDS", 20)
    api_cache_max_entries: int = _env_int("API_CACHE_MAX_ENTRIES", 256)

    # --- MICROSOFT GRAPH EMAIL SETTINGS ---
    graph_client_id: str = os.getenv("GRAPH_CLIENT_ID") or os.getenv("OUTLOOK_CLIENT_ID", "")
    graph_tenant_id: str = os.getenv("GRAPH_TENANT_ID") or os.getenv("OUTLOOK_TENANT_ID", "common")
    graph_scopes: list[str] = _graph_scopes()
    graph_token_cache: str = os.getenv("GRAPH_TOKEN_CACHE") or os.getenv("OUTLOOK_TOKEN_CACHE", "config/graph_token_cache.json")
    graph_client_secret: str = os.getenv("GRAPH_CLIENT_SECRET") or os.getenv("OUTLOOK_CLIENT_SECRET", "")
    graph_sender_email: str = os.getenv("GRAPH_SENDER_EMAIL", "")
    graph_refresh_token: str = os.getenv("GRAPH_REFRESH_TOKEN") or os.getenv("OUTLOOK_REFRESH_TOKEN", "")
    graph_http_timeout: int = _env_int("GRAPH_HTTP_TIMEOUT", 30)

    # Backward-compatible aliases for older code/config.
    outlook_client_id: str = graph_client_id
    outlook_tenant_id: str = graph_tenant_id
    outlook_scopes: list[str] = graph_scopes
    outlook_token_cache: str = graph_token_cache
    outlook_client_secret: str = graph_client_secret

    @property
    def sql_connection_string(self) -> str:
        if self.db_connection_string:
            return self.db_connection_string
        return (
            f"DRIVER={{{self.db_driver}}};"
            f"SERVER={self.db_server};"
            f"DATABASE={self.db_database};"
            f"UID={self.db_username};"
            f"PWD={self.db_password};"
            f"Encrypt={self.db_encrypt};"
            f"TrustServerCertificate={self.db_trust_certificate};"
            f"Connection Timeout={self.db_connection_timeout};"
        )

@lru_cache
def get_settings() -> Settings:
    return Settings()
