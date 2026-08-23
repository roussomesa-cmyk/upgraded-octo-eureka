"""
report_summary.py
រៀងរាល់ថ្ងៃ ដំណើរការឆ្លងកាត់ sheet គម្រោងទាំងអស់ (SHEET_NAMES) រាប់ Approved/Not Approved/ល
ពិតប្រាកដពីជួរឈរ 'Team' និង 'Result' រួចផ្ញើរបាយការណ៍ ៣ប្រភេទ៖

  ១. សរុបការងាររបស់ Team នីមួយៗ (រួមទាំងអស់ sheet) -> ក្រុមផ្ទាល់ខ្លួន Team នោះ
     (យោង Sheet "Team chat IDs" ជួរឈរ A=Team, B=ChatID)
  ២. សរុបការងារ Sheet/Task នីមួយៗ (គ្រប់ Team) -> ក្រុមទទួលខុសត្រូវ Sheet នោះ
     (យោង Sheet "Team chat IDs" ជួរឈរ C=Sheet, D=ChatID)
  ៣. សរុបទាំងអស់ (គ្រប់ Team + Sheet) -> MAIN_GROUP_ID (ក្រុម "CHA_Power Dept.")
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
MAIN_GROUP_ID = -1001853372580  # ក្រុម "CHA_Power Dept."

SHEET_NAMES = [
    "A1", "A2", "A4", "A5", "A6", "A7", "A8",
    "B1", "B2", "B3", "B4", "B5", "B6", "B7", "B8", "B9",
    "B11", "B12", "B13", "B14", "B16", "B17", "B18",
    "C1", "C2", "C3", "C4", "C7", "C8", "C10",
    "E1", "E2", "E3", "E4", "E5", "E6", "E13", "E15",
]

DELAY_BETWEEN_MESSAGES_SEC = 2


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


def load_team_and_sheet_groups(sh):
    """អាន Sheet 'Team chat IDs':
    ជួរឈរ A,B = Team (ឧ. CHA_TEAM01) -> ChatID
    ជួរឈរ C,D = Sheet name (ឧ. B4) -> ChatID
    ត្រឡប់ (team_groups, sheet_groups) ជា dict ទាំងពីរ"""
    try:
        ws = sh.worksheet("Team chat IDs")
    except gspread.WorksheetNotFound:
        return {}, {}

    team_groups = {}
    sheet_groups = {}
    for row in ws.get_all_values()[1:]:  # រំលងបន្ទាត់ header
        if len(row) > 1 and row[0].strip() and row[1].strip():
            raw_team = row[0].strip()
            team_code = raw_team.replace("_TEAM0", "-T0").replace("_TEAM", "-T")
            team_groups[team_code] = row[1].strip()
        if len(row) > 3 and row[2].strip() and row[3].strip():
            sheet_groups[row[2].strip()] = row[3].strip()

    return team_groups, sheet_groups


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


def counts_line(counts):
    approved = counts.get("Approved", 0)
    not_approved = counts.get("Not Approved", 0)
    total = sum(counts.values())
    remain = total - approved - not_approved
    return approved, not_approved, remain, total


def main():
    sh = get_spreadsheet()
    team_groups, sheet_groups = load_team_and_sheet_groups(sh)
    today = datetime.now().strftime("%d-%m-%Y")

    # team -> status -> count (ឆ្លងកាត់ sheet ទាំងអស់)
    team_totals = defaultdict(lambda: defaultdict(int))
    # status -> count (សរុបទាំងអស់)
    grand_totals = defaultdict(int)

    with TelegramClient(StringSession(SESSION_STRING), API_ID, API_HASH) as client:

        # ============================================================
        # ដំណើរការឆ្លងកាត់ sheet នីមួយៗ - ចាំបាច់ត្រូវធ្វើមុន ដើម្បីប្រមូលទិន្នន័យសរុប
        # ============================================================
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

            header = values[header_idx]
            team_col = header.index("Team")
            result_col = header.index("Result")

            sheet_summary = defaultdict(lambda: defaultdict(int))
            for row in values[header_idx + 1:]:
                if len(row) <= max(team_col, result_col):
                    continue
                team = row[team_col].strip()
                if not team:
                    continue
                status = normalize_status(row[result_col])
                sheet_summary[team][status] += 1
                team_totals[team][status] += 1
                grand_totals[status] += 1

            if not sheet_summary:
                print(f"[{sheet_name}] គ្មានទិន្នន័យ - រំលង")
                continue

            # ===== ប្រភេទទី ២៖ សរុបការងារ sheet នេះ (គ្រប់ Team) -> ក្រុមទទួលខុសត្រូវ =====
            chat_id = sheet_groups.get(sheet_name)
            if chat_id:
                lines = [f"របាយការណ៍ការងារ {sheet_name} - {today}", ""]
                for team in sorted(sheet_summary.keys()):
                    approved, not_approved, remain, total = counts_line(sheet_summary[team])
                    lines.append(
                        f"{team}: Approved {approved} | Not Approved {not_approved} | "
                        f"នៅសល់ {remain} | សរុប {total}"
                    )
                try:
                    client.send_message(chat_id, "\n".join(lines))
                    print(f"[Type2] {sheet_name} -> {chat_id}")
                except Exception as e:
                    print(f"[Type2] {sheet_name} -> {chat_id} - ERROR (គណនីប្រហែលមិនទាន់ចូលក្រុមនេះ): {e}")
                time.sleep(DELAY_BETWEEN_MESSAGES_SEC)
            else:
                print(f"[Type2] {sheet_name} - គ្មាន Chat ID កំណត់ក្នុង 'Team chat IDs' ជួរឈរ C/D - រំលង")

        # ============================================================
        # ប្រភេទទី ១៖ សរុបការងាររបស់ Team នីមួយៗ (ឆ្លងកាត់ sheet ទាំងអស់) -> ក្រុមផ្ទាល់ខ្លួន Team
        # ============================================================
        for team in sorted(team_totals.keys()):
            chat_id = team_groups.get(team)
            if not chat_id:
                print(f"[Type1] {team} - គ្មាន Chat ID កំណត់ក្នុង 'Team chat IDs' ជួរឈរ A/B - រំលង")
                continue

            approved, not_approved, remain, total = counts_line(team_totals[team])
            pct = round(approved / total * 100, 1) if total else 0
            msg = (
                f"របាយការណ៍សរុបការងារ {team} - {today}\n\n"
                f"   Approved: {approved}\n"
                f"   Not Approved: {not_approved}\n"
                f"   នៅសល់: {remain}\n"
                f"   % សម្រេច: {pct}%\n"
                f"   សរុប: {total}"
            )
            try:
                client.send_message(chat_id, msg)
                print(f"[Type1] {team} -> {chat_id}")
            except Exception as e:
                print(f"[Type1] {team} -> {chat_id} - ERROR (គណនីប្រហែលមិនទាន់ចូលក្រុមនេះ): {e}")
            time.sleep(DELAY_BETWEEN_MESSAGES_SEC)

        # ============================================================
        # ប្រភេទទី ៣៖ សរុបទាំងអស់ (គ្រប់ Team + Sheet) -> MAIN_GROUP_ID
        # ============================================================
        approved, not_approved, remain, total = counts_line(grand_totals)
        pct = round(approved / total * 100, 1) if total else 0
        msg = (
            f"របាយការណ៍សរុបរួមទាំងអស់ - {today}\n\n"
            f"   Approved: {approved}\n"
            f"   Not Approved: {not_approved}\n"
            f"   នៅសល់: {remain}\n"
            f"   % សម្រេច: {pct}%\n"
            f"   សរុប: {total}"
        )
        try:
            client.send_message(MAIN_GROUP_ID, msg)
            print(f"[Type3] Overall -> {MAIN_GROUP_ID}")
        except Exception as e:
            print(f"[Type3] Overall -> {MAIN_GROUP_ID} - ERROR (គណនីប្រហែលមិនទាន់ចូលក្រុមនេះ): {e}")


if __name__ == "__main__":
    main()
