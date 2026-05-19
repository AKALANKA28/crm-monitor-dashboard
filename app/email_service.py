import base64
import json
import msal
import requests
from pathlib import Path
from app.config import Settings

class MicrosoftGraphEmailService:
    def __init__(self, settings: Settings):
        self.settings = settings
        self.authority = f"https://login.microsoftonline.com/{settings.graph_tenant_id}"
        self.cache_path = Path(settings.graph_token_cache)
        self.app = None
        self.cache = None

    def _build_msal_app(self):
        if self.app:
            return self.app

        if self.settings.graph_client_secret:
            self.cache = None
            self.app = msal.ConfidentialClientApplication(
                self.settings.graph_client_id,
                authority=self.authority,
                client_credential=self.settings.graph_client_secret,
            )
        else:
            self.cache = msal.SerializableTokenCache()
            if self.cache_path.exists():
                self.cache.deserialize(self.cache_path.read_text())
            
            self.app = msal.PublicClientApplication(
                self.settings.graph_client_id,
                authority=self.authority,
                token_cache=self.cache
            )
        return self.app

    def _save_cache(self) -> None:
        if not self.cache or not self.cache.has_state_changed:
            return
        self.cache_path.parent.mkdir(parents=True, exist_ok=True)
        self.cache_path.write_text(self.cache.serialize())

    def _get_access_token(self, allow_device_flow: bool = False) -> str:
        result = None

        if self.settings.graph_refresh_token:
            result = self._acquire_token_by_refresh_token()

        app = None
        if not result:
            app = self._build_msal_app()
            accounts = app.get_accounts()
            if accounts:
                result = app.acquire_token_silent(self.settings.graph_scopes, account=accounts[0])

        if not result and self.settings.graph_client_secret:
            app = app or self._build_msal_app()
            result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])

        if not result and allow_device_flow and not self.settings.graph_client_secret:
            app = app or self._build_msal_app()
            flow = app.initiate_device_flow(scopes=self.settings.graph_scopes)
            if "user_code" not in flow:
                raise Exception(f"Failed to create device login flow: {flow}")
            print(flow["message"])
            result = app.acquire_token_by_device_flow(flow)
            
        if not result or "access_token" not in result:
            error = result.get("error_description") if isinstance(result, dict) else None
            detail = f" Details: {error}" if error else ""
            raise Exception(
                "Failed to acquire token. Run the email test with device-code login or ensure the token cache is populated."
                + detail
            )

        self._save_cache()
            
        return result["access_token"]

    def _acquire_token_by_refresh_token(self) -> dict:
        token_url = f"{self.authority}/oauth2/v2.0/token"
        payload = {
            "client_id": self.settings.graph_client_id,
            "grant_type": "refresh_token",
            "refresh_token": self.settings.graph_refresh_token,
            "scope": " ".join(self.settings.graph_scopes),
        }
        if self.settings.graph_client_secret:
            payload["client_secret"] = self.settings.graph_client_secret

        response = requests.post(token_url, data=payload, timeout=self.settings.graph_http_timeout)
        try:
            result = response.json()
        except ValueError:
            result = {"error_description": response.text}

        if response.status_code >= 400:
            return result
        return result

    def _send_mail_endpoint(self) -> str:
        if self.settings.graph_client_secret:
            if self.settings.graph_sender_email:
                return f"https://graph.microsoft.com/v1.0/users/{self.settings.graph_sender_email}/sendMail"
            raise Exception("GRAPH_SENDER_EMAIL is required when using GRAPH_CLIENT_SECRET app-only auth.")
        return "https://graph.microsoft.com/v1.0/me/sendMail"

    @staticmethod
    def _split_emails(value: str) -> list[str]:
        return [item.strip() for item in value.replace(";", ",").split(",") if item.strip()]

    def _normalize_emails(self, emails: str | list[str] | None) -> list[str]:
        if not emails:
            return []
        if isinstance(emails, list):
            output: list[str] = []
            for item in emails:
                if item:
                    output.extend(self._split_emails(str(item)))
            return output
        return self._split_emails(str(emails))

    @staticmethod
    def _build_recipients(emails: list[str]) -> list[dict[str, dict[str, str]]]:
        return [{"emailAddress": {"address": email}} for email in emails]

    def send_report_email(
        self,
        to_email: str | list[str],
        subject: str,
        body: str,
        file_bytes: bytes,
        filename: str,
        cc_emails: str | list[str] | None = None,
        bcc_emails: str | list[str] | None = None,
        allow_device_flow: bool = False,
    ):
        token = self._get_access_token(allow_device_flow=allow_device_flow)
        endpoint = self._send_mail_endpoint()

        to_list = self._normalize_emails(to_email)
        if not to_list:
            raise Exception("At least one recipient email address is required.")
        cc_list = self._normalize_emails(cc_emails)
        bcc_list = self._normalize_emails(bcc_emails)
        
        email_data = {
            "message": {
                "subject": subject,
                "body": {
                    "contentType": "Text",
                    "content": body
                },
                "toRecipients": self._build_recipients(to_list),
                "attachments": [
                    {
                        "@odata.type": "#microsoft.graph.fileAttachment",
                        "name": filename,
                        "contentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        "contentBytes": base64.b64encode(file_bytes).decode("utf-8")
                    }
                ]
            },
            "saveToSentItems": "true"
        }

        if cc_list:
            email_data["message"]["ccRecipients"] = self._build_recipients(cc_list)
        if bcc_list:
            email_data["message"]["bccRecipients"] = self._build_recipients(bcc_list)

        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }

        try:
            response = requests.post(
                endpoint,
                headers=headers,
                json=email_data,
                timeout=self.settings.graph_http_timeout,
            )
        except requests.RequestException as exc:
            raise Exception(
                "Failed to reach Microsoft Graph. Check network access, proxy settings, or TLS inspection."
            ) from exc
        
        if response.status_code != 202:
            raise Exception(f"Failed to send email ({response.status_code}): {response.text}")


OutlookEmailService = MicrosoftGraphEmailService
