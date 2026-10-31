"""Run ONCE on your own PC (needs a browser):   python authorize_google.py

It signs you in to Google (Drive + Sheets), writes token.json, and prints the line to paste
into your server's secrets so the deployed app can use Google without a browser.
"""
import json

from google_auth_oauthlib.flow import InstalledAppFlow

from drive_uploader import CREDENTIALS_FILE, SCOPES, TOKEN_FILE

if not CREDENTIALS_FILE.exists():
    raise SystemExit(f"Put your OAuth client file at {CREDENTIALS_FILE} first (see README).")

flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_FILE), SCOPES)
creds = flow.run_local_server(port=0)
TOKEN_FILE.write_text(creds.to_json(), encoding="utf-8")

print(f"\nSaved {TOKEN_FILE.name}. For a deployed app, add this to its secrets:\n")
print("GOOGLE_TOKEN_JSON = '''" + json.dumps(json.loads(creds.to_json())) + "'''\n")
print("Keep it private: anyone with it can open the app's Drive files and Google Sheet.")
