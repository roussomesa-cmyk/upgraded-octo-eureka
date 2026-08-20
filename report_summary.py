"""
report_summary.py
ដំណើរការរាល់ថ្ងៃម៉ោង ៥ល្ងាច៖ ដំណើរការឆ្លងកាត់ sheet គម្រោងទាំងអស់ (SHEET_NAMES) ក្នុង Google Sheet
"Report Power CHA2026" ។ sheet ណាដែលមានជួរឈរ 'Team' និង 'Result' ព្រមទាំងមានទិន្នន័យ នឹងត្រូវរាប់
Approved / Not Approved តាម Team (CHA-T01..CHA-T07) ផ្ញើសារទី ១ រួចផ្ញើ status ដែលនៅសល់ជាសារទី ២
ទៅ Telegram group របស់ Team នោះ។ sheet ណាគ្មានទិន្នន័យ/Team នឹងត្រូវរំលងចោល ទៅ sheet បន្ទាប់។
"""

import os
import time
import json
from datetime import datetime
from collections import defaultdict

from telethon.sync import TelegramClient
from telethon.sessions import StringSession

import gspread
from google.oauth2.service_account import Credentials

# ============================================================
# CONFIG
# ============================================================
API_ID = int(os.environ["TELEGRAM_API_ID"])
API_HASH = os.environ["TELEGRAM_API_HASH"]
SESSION_STRING = os.environ["TELEGRAM_SESSION"]
GOOGLE_CREDS_JSON = os.environ["GCP_SA_KEY"]

REPORT_SPREADSHEET_ID = os.environ["REPORT_SPREADSHEET_ID"]
REPORT_NOTIFY_GROUP_ID = os.environ["REPORT_NOTIFY_GROUP_ID"]

SHEET_NAMES = [
    "A1", "A2", "A4", "A5", "A6", "A7", "A8",
    "B1", "B2", "B3", "B4", "B5", "B6", "B7", "B8", "B9",
    "B11", "B12", "B13", "B14", "B16", "B17", "B18",
    "C1", "C2", "C3", "C4", "C7", "C8", "C10",
    "E1", "E2", "E3", "E4", "E5", "E6", "E13", "E15",
]

DELAY_BETWEEN_MESSAGES_SEC = 2  # ចន្លោះពេលរវាងសារនីមួយៗ ជៀសវាងការរឹតត្បិតរបស់ Telegram


def load_sheet_groups(sh):
    """អាន Sheet 'Team chat IDs' ជួរឈរ C (Sheet name) និង D (ChatID) - ត្រឡប់ dict: sheet_name -> chat_id
    (ជួរឈរ A/B ប្រើសម្រាប់គោលបំណងផ្សេង - មិនប៉ះពាល់ទេ)"""
    try:
        ws = sh.worksheet("Team chat IDs")
    except gspread.WorksheetNotFound:
        return {}
    mapping = {}
    for row in ws.get_all_values()[1:]:  # រំលងបន្ទាត់ header
        if len(row) > 3 and row[2].strip() and row[3].strip():
            mapping[row[2].strip()] = row[3].strip()
    return mapping


def get_spreadsheet():
    creds_dict = json.loads(GOOGLE_CREDS_JSON)
    scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    gc = gspread.authorize(creds)

    last_error = None
    for attempt in range(5):
        try:
            return gc.open_by_key(REPORT_SPREADSHEET_ID)
        except gspread.exceptions.APIError as e:
            last_error = e
            wait_sec = 5 * (attempt + 1)
            print(f"Google Sheets API error (សាកល្បងទី {attempt + 1}/5) - រង់ចាំ {wait_sec}s: {e}")
            time.sleep(wait_sec)
    raise last_error


def normalize_status(raw):
    s = (raw or "").strip()
    if not s or s == "-":
        return "(-)"
    return s


def find_header_row(values):
    """ស្វែងរកបន្ទាត់ header ដែលមានទាំង 'Team' និង 'Result' (ក្នុង 10 បន្ទាត់ដំបូង)។ ត្រឡប់ None បើរកមិនឃើញ"""
    for i, row in enumerate(values[:10]):
        if "Team" in row and "Result" in row:
            return i
    return None


def build_summary(values, header_idx):
    header = values[header_idx]
    team_col = header.index("Team")
    result_col = header.index("Result")

    summary = defaultdict(lambda: defaultdict(int))
    for row in values[header_idx + 1:]:
        if len(row) <= max(team_col, result_col):
            continue
        team = row[team_col].strip()
        if not team:
            continue
        status = normalize_status(row[result_col])
        summary[team][status] += 1

    return summary


def format_main_message(sheet_name, team, counts, today):
    approved = counts.get("Approved", 0)
    not_approved = counts.get("Not Approved", 0)
    return (
        f"📊 របាយការណ៍ {sheet_name} - {today}\n"
        f"👥 Team {team}\n\n"
        f"   Approved: {approved}\n"
        f"   Not Approved: {not_approved}"
    )


def format_remaining_message(sheet_name, team, counts, today):
    remaining = {k: v for k, v in counts.items() if k not in ("Approved", "Not Approved")}
    if not remaining:
        return None
    lines = [f"📋 ទិន្នន័យនៅសល់ {sheet_name} - {today}", f"👥 Team {team}", ""]
    for status, count in remaining.items():
        lines.append(f"   {status}: {count}")
    return "\n".join(lines)


def main():
    sh = get_spreadsheet()
    sheet_groups = load_sheet_groups(sh)
    today = datetime.now().strftime("%d-%m-%Y")

    with TelegramClient(StringSession(SESSION_STRING), API_ID, API_HASH) as client:
        client.get_dialogs()  # ទាញយកបញ្ជីក្រុម/ឆាតទាំងអស់ជាមុន ដើម្បីអោយ entity cache ស្គាល់ Chat ID គ្រប់ក្រុម
        for sheet_name in SHEET_NAMES:
            try:
                ws = sh.worksheet(sheet_name)
            except gspread.WorksheetNotFound:
                print(f"[{sheet_name}] រកមិនឃើញ sheet នេះ - រំលង")
                continue

            values = ws.get_all_values()
            header_idx = find_header_row(values)
            if header_idx is None:
                print(f"[{sheet_name}] គ្មានជួរឈរ 'Team'/'Result' - រំលង")
                continue

            summary = build_summary(values, header_idx)
            if not summary:
                print(f"[{sheet_name}] គ្មានទិន្នន័យ - រំលង")
                continue

            chat_id = sheet_groups.get(sheet_name, REPORT_NOTIFY_GROUP_ID)

            for team in sorted(summary.keys()):
                counts = summary[team]

                main_msg = format_main_message(sheet_name, team, counts, today)
                client.send_message(chat_id, main_msg)
                print(f"[{sheet_name}] Team {team} -> {chat_id} (Approved/Not Approved)")
                time.sleep(DELAY_BETWEEN_MESSAGES_SEC)

                remaining_msg = format_remaining_message(sheet_name, team, counts, today)
                if remaining_msg:
                    client.send_message(chat_id, remaining_msg)
                    print(f"[{sheet_name}] Team {team} -> {chat_id} (Remaining)")
                    time.sleep(DELAY_BETWEEN_MESSAGES_SEC)


if __name__ == "__main__":
    main()


