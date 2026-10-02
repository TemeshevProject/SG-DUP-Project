#!/usr/bin/env python3
"""Upload survey files to Google Drive using a service account.

Note: Service accounts cannot CREATE new files in a personal Google Drive folder
(quota limit is 0). They can read shared folders and update existing files.

Workarounds:
  1. OAuth: use upload_to_drive.py with credentials.json (Desktop OAuth client)
  2. Manual: drag files into the Drive folder in the browser
  3. Shared Drive: move the target folder to a Google Workspace Shared Drive
  4. Placeholders: create empty files with the target names in the folder, then run
     this script with --update-existing
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

FOLDER_ID = "1fM-_x_Lva70JDIzi97QsIJPpjLpywMpD"
DEFAULT_FILES = [
    "DUP_oprosnik_direktorov_filialov.xlsx",
    "DUP_analiz_otvetov_oprosnika.xlsx",
    "DUP_oprosnik_Google_Forms_инструкция.md",
    "DUP_oprosnik_paket.zip",
    "ZAGRUZKA_NA_GOOGLE_DISK.md",
]
SURVEYS_DIR = Path(__file__).resolve().parent
SA_FILE = SURVEYS_DIR / "service_account.json"
SCOPES = ["https://www.googleapis.com/auth/drive"]


def get_service():
    from google.oauth2 import service_account
    from googleapiclient.discovery import build

    if not SA_FILE.exists():
        print(f"Missing service account JSON: {SA_FILE}")
        sys.exit(1)
    creds = service_account.Credentials.from_service_account_file(str(SA_FILE), scopes=SCOPES)
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def list_folder_files(service, folder_id: str) -> dict[str, str]:
    resp = service.files().list(
        q=f"'{folder_id}' in parents and trashed=false",
        fields="files(id,name)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
        pageSize=200,
    ).execute()
    return {f["name"]: f["id"] for f in resp.get("files", [])}


def upload_create(service, folder_id: str, names: list[str]) -> None:
    from googleapiclient.http import MediaFileUpload

    for name in names:
        path = SURVEYS_DIR / name
        if not path.exists():
            print(f"SKIP (missing): {name}")
            continue
        media = MediaFileUpload(str(path), resumable=True)
        meta = {"name": name, "parents": [folder_id]}
        result = service.files().create(
            body=meta,
            media_body=media,
            fields="id,name,webViewLink",
            supportsAllDrives=True,
        ).execute()
        print(f"OK: {result['name']} -> {result.get('webViewLink', result['id'])}")


def upload_update(service, folder_id: str, names: list[str]) -> None:
    from googleapiclient.http import MediaFileUpload

    existing = list_folder_files(service, folder_id)
    for name in names:
        path = SURVEYS_DIR / name
        if not path.exists():
            print(f"SKIP (missing local): {name}")
            continue
        if name not in existing:
            print(f"SKIP (no placeholder in Drive): {name}")
            continue
        media = MediaFileUpload(str(path), resumable=True)
        result = service.files().update(
            fileId=existing[name],
            media_body=media,
            fields="id,name,webViewLink",
            supportsAllDrives=True,
        ).execute()
        print(f"UPDATED: {result['name']} -> {result.get('webViewLink', result['id'])}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Upload DUP survey files to Google Drive")
    parser.add_argument("--folder-id", default=FOLDER_ID)
    parser.add_argument("--update-existing", action="store_true", help="Update files that already exist in folder")
    parser.add_argument("files", nargs="*", default=DEFAULT_FILES)
    args = parser.parse_args()

    service = get_service()
    if args.update_existing:
        upload_update(service, args.folder_id, args.files)
    else:
        upload_create(service, args.folder_id, args.files)


if __name__ == "__main__":
    main()
