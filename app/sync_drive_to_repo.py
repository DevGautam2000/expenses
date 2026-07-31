#!/usr/bin/env python3
"""
Sync new files from a Google Drive folder into a destination folder in this repo.

How it decides "new":
  A manifest file (.gdrive-sync-manifest.json) is kept inside the destination
  folder. It maps Drive file ID -> {name, md5Checksum, modifiedTime}. On each
  run we list the Drive folder and download anything whose file ID is not yet
  in the manifest, or whose md5Checksum has changed since last sync.

  We track by Drive file ID (not filename) because filenames aren't reliable
  as a uniqueness key: two files can share a name, and a file renamed in
  Drive should count as an update, not a brand-new upload.

Native Google Docs/Sheets/Slides don't have a binary/md5 the same way — they
get exported (default: Docs->docx, Sheets->xlsx, Slides->pptx). You can change
EXPORT_MIME_MAP below if you'd rather export as PDF, etc.

Required environment variables:
  GDRIVE_SA_KEY_JSON   - full JSON contents of the service account key
  GDRIVE_FOLDER_ID      - the Drive folder ID to watch
  DEST_PATH             - path inside this repo to sync files into (e.g. "data/incoming")

Exit behavior:
  Prints a summary. Leaves changed files staged in the filesystem; the calling
  GitHub Actions workflow is responsible for `git add/commit/push`.
"""

import json
import os
import sys
from pathlib import Path

from google.oauth2 import service_account
from google.oauth2.credentials import Credentials as UserCredentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import io

SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]
MANIFEST_NAME = ".gdrive-sync-manifest.json"

# Native Google file mimeTypes -> (export mimeType, file extension)
EXPORT_MIME_MAP = {
    "application/vnd.google-apps.document": (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        ".docx",
    ),
    "application/vnd.google-apps.spreadsheet": (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".xlsx",
    ),
    "application/vnd.google-apps.presentation": (
        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        ".pptx",
    ),
}


def get_env(name: str) -> str:
    val = os.environ.get(name)
    if not val:
        print(f"ERROR: missing required environment variable {name}", file=sys.stderr)
        sys.exit(1)
    return val


def build_drive_service():
    # Prefer OAuth (personal account) if those secrets are present, otherwise
    # fall back to a service account. Only one set needs to be configured.
    client_id = os.environ.get("GDRIVE_OAUTH_CLIENT_ID")
    client_secret = os.environ.get("GDRIVE_OAUTH_CLIENT_SECRET")
    refresh_token = os.environ.get("GDRIVE_OAUTH_REFRESH_TOKEN")

    if client_id and client_secret and refresh_token:
        creds = UserCredentials(
            token=None,
            refresh_token=refresh_token,
            client_id=client_id,
            client_secret=client_secret,
            token_uri="https://oauth2.googleapis.com/token",
            scopes=SCOPES,
        )
        return build("drive", "v3", credentials=creds)

    key_json = get_env("GDRIVE_SA_KEY_JSON")
    info = json.loads(key_json)
    creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
    return build("drive", "v3", credentials=creds)


def list_drive_files(service, folder_id: str):
    files = []
    page_token = None
    query = f"'{folder_id}' in parents and trashed = false"
    fields = "nextPageToken, files(id, name, mimeType, md5Checksum, modifiedTime)"
    while True:
        resp = (
            service.files()
            .list(q=query, fields=fields, pageToken=page_token, pageSize=200)
            .execute()
        )
        files.extend(resp.get("files", []))
        page_token = resp.get("nextPageToken")
        if not page_token:
            break
    return files


def load_manifest(dest: Path) -> dict:
    manifest_path = dest / MANIFEST_NAME
    if manifest_path.exists():
        return json.loads(manifest_path.read_text())
    return {}


def save_manifest(dest: Path, manifest: dict):
    manifest_path = dest / MANIFEST_NAME
    manifest_path.write_text(json.dumps(manifest, indent=2, sort_keys=True))


def download_binary(service, file_id: str, dest_file: Path):
    request = service.files().get_media(fileId=file_id)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    dest_file.write_bytes(buf.getvalue())


def download_export(service, file_id: str, export_mime: str, dest_file: Path):
    request = service.files().export_media(fileId=file_id, mimeType=export_mime)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    dest_file.write_bytes(buf.getvalue())


def safe_filename(name: str, forced_ext: str | None = None) -> str:
    name = name.replace("/", "-").replace("\\", "-")
    if forced_ext and not name.endswith(forced_ext):
        name = f"{name}{forced_ext}"
    return name


def main():
    folder_id = get_env("GDRIVE_FOLDER_ID")
    dest = Path(get_env("DEST_PATH"))
    dest.mkdir(parents=True, exist_ok=True)

    service = build_drive_service()
    drive_files = list_drive_files(service, folder_id)
    manifest = load_manifest(dest)

    new_count = 0
    updated_count = 0

    for f in drive_files:
        file_id = f["id"]
        mime = f["mimeType"]
        prior = manifest.get(file_id)

        if mime in EXPORT_MIME_MAP:
            export_mime, ext = EXPORT_MIME_MAP[mime]
            change_key = f.get("modifiedTime")  # exported docs have no md5
            local_name = safe_filename(f["name"], ext)
        else:
            export_mime, ext = None, None
            change_key = f.get("md5Checksum")
            local_name = safe_filename(f["name"])

        is_new = prior is None
        is_updated = prior is not None and prior.get("changeKey") != change_key

        if not (is_new or is_updated):
            continue

        target_path = dest / local_name
        try:
            if export_mime:
                download_export(service, file_id, export_mime, target_path)
            else:
                download_binary(service, file_id, target_path)
        except Exception as e:
            print(f"WARN: failed to download '{f['name']}' ({file_id}): {e}", file=sys.stderr)
            continue

        manifest[file_id] = {
            "name": local_name,
            "changeKey": change_key,
            "modifiedTime": f.get("modifiedTime"),
        }

        if is_new:
            new_count += 1
            print(f"NEW: {local_name}")
        else:
            updated_count += 1
            print(f"UPDATED: {local_name}")

    save_manifest(dest, manifest)

    print(f"\nDone. {new_count} new file(s), {updated_count} updated file(s).")
    # Signal to the workflow whether there's anything to commit.
    if os.environ.get("GITHUB_OUTPUT"):
        with open(os.environ["GITHUB_OUTPUT"], "a") as gh_out:
            gh_out.write(f"changed={'true' if (new_count or updated_count) else 'false'}\n")


if __name__ == "__main__":
    main()
