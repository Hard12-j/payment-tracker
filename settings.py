"""Shared settings. Values come from environment variables, then .streamlit/secrets.toml,
then (optionally) a .env file."""
import os
from pathlib import Path

try:  # optional: lets you keep using a .env file
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

import streamlit as st

BASE_DIR = Path(__file__).parent
EXCEL_FILE = BASE_DIR / "payments.xlsx"
RECEIPT_DIR = BASE_DIR / "receipts"
RECEIPT_DIR.mkdir(exist_ok=True)


def get_setting(key: str, default: str = "") -> str:
    if key in os.environ:
        return os.environ[key]
    try:
        return str(st.secrets[key])
    except Exception:
        return default


def get_bool_setting(key: str, default: bool = True) -> bool:
    if key in os.environ:
        return os.environ[key].strip().lower() in ("true", "1", "yes")
    try:
        val = st.secrets[key]
        if isinstance(val, bool):
            return val
        return str(val).strip().lower() in ("true", "1", "yes")
    except Exception:
        return default


ORG_NAME = get_setting("ORG_NAME", "Your Organisation Name")
ORG_ADDRESS = get_setting("ORG_ADDRESS", "Your address, City, State - PIN")
ORG_PHONE = get_setting("ORG_PHONE", "")
RECEIPT_PREFIX = get_setting("RECEIPT_PREFIX", "RCP")
SMTP_HOST = get_setting("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT = int(get_setting("SMTP_PORT", "465"))
SMTP_USER = get_setting("SMTP_USER")
SMTP_PASSWORD = get_setting("SMTP_PASSWORD")
DEFAULT_COUNTRY_CODE = get_setting("DEFAULT_COUNTRY_CODE", "91")
ENABLE_DRIVE_UPLOAD = get_bool_setting("ENABLE_DRIVE_UPLOAD", default=True)

# ---- storage: "sheets" (Google Sheets, for deployment) or "excel" (local payments.xlsx) ----
STORAGE_BACKEND = get_setting("STORAGE_BACKEND", "excel").strip().lower()
SHEETS_SPREADSHEET_ID = get_setting("SHEETS_SPREADSHEET_ID").strip()   # optional
SHEETS_NAME = get_setting("SHEETS_NAME", "Gayatri Nagar Payments")
