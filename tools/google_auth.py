# tools/google_auth.py
"""
OAuth 2.0 helpers for Gmail and Google Calendar APIs.

Credentials are read from .env (GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET).
No credentials.json needed on disk.

First run: opens a browser window for each account to grant access.
Subsequent runs: uses saved token files (auto-refreshed).

Token files:
  token_gmail.json    — ryanliu61799@gmail.com (Gmail read)
  token_calendar.json — rliu07979@gmail.com   (Calendar write)
"""

import json
import os

from dotenv import load_dotenv
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

load_dotenv()

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

GMAIL_SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]
GMAIL_SEND_SCOPES = ["https://www.googleapis.com/auth/gmail.send"]
CALENDAR_SCOPES = ["https://www.googleapis.com/auth/calendar.events"]

GMAIL_TOKEN = os.path.join(PROJECT_ROOT, "token_gmail.json")
GMAIL_SEND_TOKEN = os.path.join(PROJECT_ROOT, "token_gmail_send.json")
CALENDAR_TOKEN = os.path.join(PROJECT_ROOT, "token_calendar_work.json")


def _client_config() -> dict:
    """Build OAuth client config dict from environment variables."""
    client_id = os.environ.get("GOOGLE_CLIENT_ID")
    client_secret = os.environ.get("GOOGLE_CLIENT_SECRET")
    project_id = os.environ.get("GOOGLE_PROJECT_ID", "")
    if not client_id or not client_secret:
        raise EnvironmentError(
            "GOOGLE_CLIENT_ID and GOOGLE_CLIENT_SECRET must be set in .env"
        )
    return {
        "installed": {
            "client_id": client_id,
            "client_secret": client_secret,
            "project_id": project_id,
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
            "redirect_uris": ["http://localhost"],
        }
    }


def _get_credentials(token_path: str, scopes: list[str], account_hint: str) -> Credentials:
    """Load or create OAuth credentials for a given token file."""
    creds = None

    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, scopes)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            print(f"\n🔐 Opening browser to authorise {account_hint}...")
            flow = InstalledAppFlow.from_client_config(_client_config(), scopes)
            creds = flow.run_local_server(port=0)

        with open(token_path, "w") as f:
            f.write(creds.to_json())
        print(f"   ✅ Token saved → {os.path.basename(token_path)}")

    return creds


def get_gmail_service():
    """Return an authenticated Gmail API service for ryanliu61799@gmail.com."""
    creds = _get_credentials(GMAIL_TOKEN, GMAIL_SCOPES, "ryanliu61799@gmail.com (Gmail)")
    return build("gmail", "v1", credentials=creds)


def get_gmail_send_service():
    """Return an authenticated Gmail API service with send scope for ryanliu61799@gmail.com."""
    creds = _get_credentials(GMAIL_SEND_TOKEN, GMAIL_SEND_SCOPES, "ryanliu61799@gmail.com (Gmail Send)")
    return build("gmail", "v1", credentials=creds)


def get_calendar_service():
    """Return an authenticated Google Calendar API service for ryanliu61799@gmail.com."""
    creds = _get_credentials(CALENDAR_TOKEN, CALENDAR_SCOPES, "ryanliu61799@gmail.com (Calendar)")
    return build("calendar", "v3", credentials=creds)
