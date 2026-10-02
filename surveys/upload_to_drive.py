#!/usr/bin/env python3
"""Upload survey files to Google Drive folder (requires OAuth credentials.json)."""

import sys
from pathlib import Path

FOLDER_ID = "1fM-_x_Lva70JDIzi97QsIJPpjLpywMpD"
FILES = [
    "DUP_oprosnik_direktorov_filialov.xlsx",
    "DUP_analiz_otvetov_oprosnika.xlsx",
    "DUP_oprosnik_Google_Forms_инструкция.md",
    "ZAGRUZKA_NA_GOOGLE_DISK.md",
]

SURVEYS_DIR = Path(__file__).resolve().parent
CREDENTIALS = SURVEYS_DIR / "credentials.json"
TOKEN = SURVEYS_DIR / "token.json"


def main():
    try:
        from google.auth.transport.requests import Request
        from google.oauth2.credentials import Credentials
        from google_auth_oauthlib.flow import InstalledAppFlow
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaFileUpload
    except ImportError:
        print("Install: pip install google-api-python-client google-auth-oauthlib")
        sys.exit(1)

    if not CREDENTIALS.exists():
        print(f"Place OAuth credentials.json in: {CREDENTIALS}")
        print("See ZAGRUZKA_NA_GOOGLE_DISK.md — Способ 3")
        sys.exit(1)

    scopes = ["https://www.googleapis.com/auth/drive.file"]
    creds = None
    if TOKEN.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN), scopes)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS), scopes)
            creds = flow.run_local_server(port=0)
        TOKEN.write_text(creds.to_json())

    service = build("drive", "v3", credentials=creds)

    for name in FILES:
        path = SURVEYS_DIR / name
        if not path.exists():
            print(f"Skip (not found): {name}")
            continue
        media = MediaFileUpload(str(path), resumable=True)
        meta = {"name": name, "parents": [FOLDER_ID]}
        result = service.files().create(body=meta, media_body=media, fields="id,name,webViewLink").execute()
        print(f"Uploaded: {result['name']} → {result.get('webViewLink', result['id'])}")


if __name__ == "__main__":
    main()
