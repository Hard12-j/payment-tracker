"""
Google plumbing shared by the whole app (OAuth login, Drive uploads, Sheets access).

* One OAuth login covers both Drive and Sheets.
* On your PC the login can open a browser (credentials.json -> token.json).
* On a SERVER there is no browser: put the contents of token.json in the setting
  GOOGLE_TOKEN_JSON (create it once on your PC with  python authorize_google.py).
* googleapiclient objects are not thread-safe and Streamlit serves users in threads, so every
  Google call in the app runs under API_LOCK.
"""
import json
import os
import threading
from pathlib import Path
from typing import Optional, Union

import google_auth_httplib2  # noqa: F401 (required for OAuth HTTP transport)
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseUpload

BASE_DIR = Path(__file__).resolve().parent
CREDENTIALS_FILE = BASE_DIR / "credentials.json"
TOKEN_FILE = BASE_DIR / "token.json"

SCOPES = [
    "https://www.googleapis.com/auth/drive.file",      # files/folders this app creates
    "https://www.googleapis.com/auth/spreadsheets",    # the payments Google Sheet
]

API_LOCK = threading.RLock()
_FOLDER_CACHE: dict = {}
_SERVICES: dict = {}


def get_setting(key: str, default: str = "") -> str:
    """Environment variable, then streamlit secrets."""
    if key in os.environ:
        return os.environ[key]
    try:
        import streamlit as st
        return str(st.secrets[key])
    except Exception:
        return default


# --------------------------------------------------------------------------
# Login
# --------------------------------------------------------------------------
def _token_info():
    """Returns (token_dict_or_None, came_from_settings)."""
    raw = get_setting("GOOGLE_TOKEN_JSON").strip()
    if raw:
        return json.loads(raw), True
    if TOKEN_FILE.exists():
        try:
            return json.loads(TOKEN_FILE.read_text(encoding="utf-8")), False
        except Exception:
            return None, False
    return None, False


def get_credentials() -> Credentials:
    with API_LOCK:
        info, from_settings = _token_info()
        creds: Optional[Credentials] = None

        if info:
            granted = set(info.get("scopes") or [])
            if granted and not set(SCOPES) <= granted:
                info = None            # login was created before Sheets was added -> must redo
            else:
                try:
                    creds = Credentials.from_authorized_user_info(info, SCOPES)
                except Exception:
                    creds = None

        if creds and not creds.valid and creds.refresh_token:
            try:
                creds.refresh(Request())
                if not from_settings:
                    TOKEN_FILE.write_text(creds.to_json(), encoding="utf-8")
            except Exception:
                creds = None

        if creds and creds.valid:
            return creds

        # Need a fresh login.
        if from_settings:
            raise RuntimeError(
                "Google login in GOOGLE_TOKEN_JSON is invalid/expired. On your PC run "
                "'python authorize_google.py' and paste the new value into your secrets.")
        if not CREDENTIALS_FILE.exists():
            raise FileNotFoundError(
                f"Google OAuth credentials file not found: '{CREDENTIALS_FILE.name}'. "
                "See README.md (Google setup).")
        flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_FILE), SCOPES)
        creds = flow.run_local_server(port=0)
        TOKEN_FILE.write_text(creds.to_json(), encoding="utf-8")
        return creds


def _service(name: str, version: str):
    with API_LOCK:
        if name not in _SERVICES:
            _SERVICES[name] = build(name, version, credentials=get_credentials(),
                                    cache_discovery=False)
        return _SERVICES[name]


def get_drive_service():
    return _service("drive", "v3")


def get_sheets_service():
    return _service("sheets", "v4")


# --------------------------------------------------------------------------
# Drive: folders + uploads
# --------------------------------------------------------------------------
def _q(value: str) -> str:
    return value.replace("\\", "\\\\").replace("'", "\\'")


def get_or_create_folder(service, folder_name: str) -> str:
    if folder_name in _FOLDER_CACHE:
        return _FOLDER_CACHE[folder_name]
    query = ("mimeType = 'application/vnd.google-apps.folder' "
             f"and name = '{_q(folder_name)}' and trashed = false")
    files = service.files().list(q=query, spaces="drive", fields="files(id, name)",
                                 pageSize=10).execute().get("files", [])
    if files:
        folder_id = files[0]["id"]
    else:
        folder_id = service.files().create(
            body={"name": folder_name, "mimeType": "application/vnd.google-apps.folder"},
            fields="id").execute().get("id")
    _FOLDER_CACHE[folder_name] = folder_id
    return folder_id


def _make_public(service, file_id: str):
    try:
        service.permissions().create(fileId=file_id,
                                     body={"type": "anyone", "role": "reader"}).execute()
    except Exception:
        pass  # permission may already exist


def upload_file(source: Union[str, Path, bytes], filename: str, folder_name: str,
                mimetype: str, public: bool) -> str:
    """Upload a file (path or raw bytes) into a Drive folder and return its webViewLink.
    If a file with the same name is already in the folder, its link is returned (no duplicate)."""
    with API_LOCK:
        service = get_drive_service()
        folder_id = get_or_create_folder(service, folder_name)

        existing = service.files().list(
            q=f"'{folder_id}' in parents and name = '{_q(filename)}' and trashed = false",
            spaces="drive", fields="files(id, name, webViewLink)", pageSize=10,
        ).execute().get("files", [])
        if existing:
            file_obj = existing[0]
            if public:
                _make_public(service, file_obj["id"])
            link = file_obj.get("webViewLink") or service.files().get(
                fileId=file_obj["id"], fields="webViewLink").execute().get("webViewLink", "")
            return link

        if isinstance(source, (bytes, bytearray)):
            import io
            media = MediaIoBaseUpload(io.BytesIO(bytes(source)), mimetype=mimetype, resumable=False)
        else:
            media = MediaFileUpload(str(source), mimetype=mimetype, resumable=True)
        uploaded = service.files().create(
            body={"name": filename, "parents": [folder_id]},
            media_body=media, fields="id, webViewLink").execute()
        file_id = uploaded.get("id")
        if public:
            _make_public(service, file_id)
        return uploaded.get("webViewLink") or service.files().get(
            fileId=file_id, fields="webViewLink").execute().get("webViewLink", "")


def upload_receipt(pdf_path: Union[str, Path], folder_name: Optional[str] = None) -> str:
    """Upload a PDF receipt (anyone-with-link can view) and return its link."""
    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"Receipt PDF not found: {pdf_path}")
    return upload_file(pdf_path, pdf_path.name,
                       folder_name or get_setting("DRIVE_FOLDER_NAME", "Payment Receipts"),
                       "application/pdf", public=True)


def upload_screenshot(data: bytes, filename: str, mimetype: str = "image/png") -> str:
    """Upload a payer's payment screenshot and return its Drive link.
    Private by default (only the Google account that owns the app can open it);
    set SCREENSHOT_LINK_PUBLIC = true to make it 'anyone with the link'."""
    public = get_setting("SCREENSHOT_LINK_PUBLIC", "false").strip().lower() in ("true", "1", "yes")
    return upload_file(data, filename,
                       get_setting("DRIVE_SCREENSHOT_FOLDER", "Payment Screenshots"),
                       mimetype, public=public)


def delete_file_by_id(file_id: str):
    with API_LOCK:
        get_drive_service().files().delete(fileId=file_id).execute()


# --------------------------------------------------------------------------
# Sheets: find or create the payments spreadsheet
# --------------------------------------------------------------------------
def find_or_create_spreadsheet(title: str, first_sheet: str) -> str:
    """Find a spreadsheet created by this app with this name, or create it."""
    with API_LOCK:
        drive = get_drive_service()
        found = drive.files().list(
            q=("mimeType = 'application/vnd.google-apps.spreadsheet' "
               f"and name = '{_q(title)}' and trashed = false"),
            spaces="drive", fields="files(id, name)", pageSize=10,
        ).execute().get("files", [])
        if found:
            return found[0]["id"]
        created = get_sheets_service().spreadsheets().create(
            body={"properties": {"title": title},
                  "sheets": [{"properties": {"title": first_sheet}}]},
            fields="spreadsheetId").execute()
        return created["spreadsheetId"]
