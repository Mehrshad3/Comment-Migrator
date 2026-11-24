import json
import os
import re
import sys
import time

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "vendor"))

from openpyxl import load_workbook
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# ==============================================================================
# CONFIGURATION PARAMETERS (Edit these before running)
# ==============================================================================
SOURCE_XLSX_FILE = "sheet name.xlsx" # Excel file exported from source Google Sheet
TARGET_SHEET_URL = "https://docs.google.com/spreadsheets/d/Enter Google Doc address here/edit"

SOURCE_TAB_NAME = "HW1T"  # Tab name inside the exported .xlsx file
TARGET_TAB_NAME = "HW1T"  # Tab name inside the destination Google Sheet

CREDENTIALS_FILE = "credentials.json"  # OAuth Desktop Client JSON from Google Cloud
TOKEN_FILE = "token.json"              # Saved after first browser sign-in
STATE_FILE = "migrated_xlsx_comments.json"  # Tracks already-copied cell comments

DRY_RUN = False                        # Set to False to actually create comments
INCLUDE_REPLIES = True                 # Also post reply lines if Google exports threaded text
REQUEST_DELAY_SEC = 1.1                # Stays safely under Sheets 60 writes/min quota
# ==============================================================================

SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets",
]

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

BOILERPLATE_RE = re.compile(
    r"^\[Threaded comment\]\s*.*?(?:\n\s*)+",
    flags=re.DOTALL,
)


def clean_comment_body(text: str) -> str:
    cleaned = BOILERPLATE_RE.sub("", text).strip()
    if cleaned.lower().startswith("comment:"):
        cleaned = cleaned[len("comment:") :].strip()
    return cleaned


def resolve_path(filename: str) -> str:
    return filename if os.path.isabs(filename) else os.path.join(BASE_DIR, filename)


def extract_spreadsheet_id(url: str) -> str:
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", url)
    if not match:
        raise ValueError(f"Could not extract spreadsheet ID from URL: {url}")
    return match.group(1)


def get_services():
    token_path = resolve_path(TOKEN_FILE)
    creds_path = resolve_path(CREDENTIALS_FILE)

    creds = None
    if os.path.exists(token_path):
        try:
            creds = Credentials.from_authorized_user_file(token_path, SCOPES)
        except (ValueError, json.JSONDecodeError):
            creds = None

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(creds_path):
                raise FileNotFoundError(
                    f"Missing '{creds_path}'. Download OAuth Desktop credentials first."
                )
            flow = InstalledAppFlow.from_client_secrets_file(creds_path, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_path, "w", encoding="utf-8") as token:
            token.write(creds.to_json())

    drive_svc = build("drive", "v3", credentials=creds)
    sheets_svc = build("sheets", "v4", credentials=creds)
    return drive_svc, sheets_svc


def load_state() -> dict:
    state_path = resolve_path(STATE_FILE)
    if os.path.exists(state_path):
        with open(state_path, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_state(state: dict) -> None:
    state_path = resolve_path(STATE_FILE)
    with open(state_path, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=2)


def get_target_sheet_id(sheets_svc, spreadsheet_id: str, tab_name: str) -> int:
    meta = (
        sheets_svc.spreadsheets()
        .get(
            spreadsheetId=spreadsheet_id,
            fields="sheets(properties(sheetId,title))",
        )
        .execute()
    )
    available = {}
    for sheet in meta.get("sheets", []):
        props = sheet.get("properties", {})
        available[props.get("title")] = props.get("sheetId")
        if props.get("title") == tab_name:
            return props.get("sheetId")

    raise ValueError(
        f"Tab '{tab_name}' not found in destination spreadsheet. "
        f"Available tabs: {list(available.keys())}"
    )


def parse_exported_comment_text(raw_text: str):
    """
    When Google Sheets exports threaded comments to .xlsx, it often formats them as:
      Author Name:
      First message
      \t-Reply Author:
      \tReply text
    Splits the main comment from any exported replies cleanly.
    """
    text = clean_comment_body((raw_text or "").replace("\r\n", "\n"))
    if not text:
        return "", []

    lines = text.split("\n")
    main_lines = []
    replies = []
    current_reply = []
    in_reply = False

    for line in lines:
        stripped = line.lstrip("\t ")
        if line.startswith("\t-") or re.match(r"^\s*-\s*[^:\n]+:\s*$", line):
            if in_reply and current_reply:
                replies.append("\n".join(current_reply).strip())
            in_reply = True
            current_reply = [stripped.lstrip("-").strip()]
        elif in_reply:
            current_reply.append(stripped)
        else:
            main_lines.append(line)

    if in_reply and current_reply:
        replies.append("\n".join(current_reply).strip())

    main_text = "\n".join(main_lines).strip()
    return main_text, [r for r in replies if r]


def read_xlsx_comments(xlsx_path: str, tab_name: str) -> list:
    if not os.path.exists(xlsx_path):
        raise FileNotFoundError(f"Missing Excel file: '{xlsx_path}'")

    wb = load_workbook(xlsx_path, data_only=True)
    if tab_name not in wb.sheetnames:
        raise ValueError(
            f"Tab '{tab_name}' not found in '{xlsx_path}'. "
            f"Available tabs: {wb.sheetnames}"
        )

    ws = wb[tab_name]
    extracted = []
    for row in ws.iter_rows():
        for cell in row:
            if cell.comment and cell.comment.text:
                main_text, replies = parse_exported_comment_text(cell.comment.text)
                if not main_text:
                    continue
                extracted.append(
                    {
                        "coord": cell.coordinate,        # e.g., "D6"
                        "row_idx": cell.row - 1,         # 0-based for Google Sheets API
                        "col_idx": cell.column - 1,      # 0-based for Google Sheets API
                        "main_text": main_text,
                        "replies": replies,
                    }
                )
    return extracted


def migrate_xlsx_comments_to_sheets():
    xlsx_path = resolve_path(SOURCE_XLSX_FILE)
    dst_id = extract_spreadsheet_id(TARGET_SHEET_URL)

    comments = read_xlsx_comments(xlsx_path, SOURCE_TAB_NAME)
    print(
        f"Found {len(comments)} commented cell(s) in '{SOURCE_XLSX_FILE}' -> tab '{SOURCE_TAB_NAME}'.\n"
    )

    if not comments:
        return

    drive_svc, sheets_svc = get_services()
    dst_sheet_id = get_target_sheet_id(sheets_svc, dst_id, TARGET_TAB_NAME)
    migrated_map = load_state()

    for item in comments:
        state_key = f"{SOURCE_TAB_NAME}!{item['coord']}"
        if state_key in migrated_map:
            print(f"[SKIP] Already migrated {state_key}.")
            continue

        print(
            f"{'[DRY RUN] ' if DRY_RUN else ''}Cell {item['coord']} "
            f"(row={item['row_idx']}, col={item['col_idx']}): {item['main_text']!r}"
        )

        if not DRY_RUN:
            try:
                batch_resp = (
                    sheets_svc.spreadsheets()
                    .batchUpdate(
                        spreadsheetId=dst_id,
                        body={
                            "requests": [
                                {
                                    "insertComment": {
                                        "content": item["main_text"],
                                        "coordinate": {
                                            "sheetId": dst_sheet_id,
                                            "rowIndex": item["row_idx"],
                                            "columnIndex": item["col_idx"],
                                        },
                                    }
                                }
                            ]
                        },
                    )
                    .execute()
                )
                migrated_map[state_key] = batch_resp
                save_state(migrated_map)
                time.sleep(REQUEST_DELAY_SEC)

                if INCLUDE_REPLIES and item["replies"]:
                    dst_comments = (
                        drive_svc.comments()
                        .list(
                            fileId=dst_id,
                            fields="comments(id,content)",
                            includeDeleted=False,
                            pageSize=100,
                        )
                        .execute()
                        .get("comments", [])
                    )
                    newest = next(
                        (
                            dc
                            for dc in reversed(dst_comments)
                            if dc.get("content") == item["main_text"]
                        ),
                        None,
                    )
                    if newest:
                        for reply_text in item["replies"]:
                            drive_svc.replies().create(
                                fileId=dst_id,
                                commentId=newest["id"],
                                fields="id",
                                body={"content": reply_text},
                            ).execute()
                            time.sleep(REQUEST_DELAY_SEC)

                print(f"  -> Posted to '{TARGET_TAB_NAME}' at exact cell {item['coord']}.")

            except HttpError as err:
                print(f"  [ERROR] Failed for cell {item['coord']}: {err}")

        print("-" * 60)


if __name__ == "__main__":
    migrate_xlsx_comments_to_sheets()
