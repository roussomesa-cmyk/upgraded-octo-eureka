"""
report_summary.py
- វេនព្រឹក (ម៉ោង < 12 PM): ផ្ញើតែ Site ដែលមិនទាន់ធ្វើ (Not yet do)
- វេនល្ងាច (ម៉ោង >= 12 PM): ផ្ញើតារាងពេញ (Approved នៅលើគេ) + តារាង Overall Summary
"""

import os
import time
import json
import random
from datetime import datetime, timedelta
from collections import defaultdict

import pandas as pd
import dataframe_image as dfi

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

TARGET_TEAMS = ["CHA-T01", "CHA-T02", "CHA-T03", "CHA-T04", "CHA-T05", "CHA-T06", "CHA-T07"]


def sleep_random_delay():
    wait_time = random.randint(30, 60)
    print(f"⏳ រង់ចាំ {wait_time} វិនាទី...")
    time.sleep(wait_time)


def load_sheet_groups(sh):
    try:
        ws = sh.worksheet("Team chat IDs")
    except gspread.WorksheetNotFound:
        return {}
    mapping = {}
    for row in ws.get_all_values()[1:]:
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
            time.sleep(wait_sec)
    raise last_error


def find_header_row(values):
    for i, row in enumerate(values[:10]):
        if "Team" in row and "Result" in row:
            return i
    return None


def parse_sheet_data(values, header_idx):
    header = values[header_idx]
    team_col = header.index("Team")
    result_col = header.index("Result")
    site_col = header.index("Site name") if "Site name" in header else None
    remark_col = header.index("Remark") if "Remark" in header else None
    qty_col = header.index("Q'ty task/site") if "Q'ty task/site" in header else None

    parsed = defaultdict(list)

    for row in values[header_idx + 1:]:
        if len(row) <= max(team_col, result_col):
            continue
        team = row[team_col].strip()
        if not team:
            continue

        result = row[result_col].strip() or "Not yet do"
        site = row[site_col].strip() if site_col is not None and len(row) > site_col else ""
        remark = row[remark_col].strip() if remark_col is not None and len(row) > remark_col else ""
        qty = row[qty_col].strip() if qty_col is not None and len(row) > qty_col else ""

        parsed[team].append({
            "No.": 1,
            "Group task": "",
            "Branch": "CHA",
            "Site name": site,
            "Q'ty task/site": qty,
            "Result": result,
            "Remark": remark,
            "Team": team
        })

    return parsed


def generate_team_image(rows_data, sheet_name, is_morning, filename="team_table.png"):
    """បង្កើតរូបភាពតារាង"""
    df = pd.DataFrame(rows_data)
    df['Group task'] = sheet_name

    if is_morning:
        # វេនព្រឹក៖ ចែកតម្រងយកតែ Site ដែលមិនទាន់ធ្វើ (Not yet do)
        df = df[df['Result'].isin(['Not yet do', '-', ''])]
        if df.empty:
            return None
    else:
        # វេនល្ងាច៖ តម្រៀប 'Approved' នៅលើគេ
        df['sort_order'] = df['Result'].apply(lambda x: 0 if x == 'Approved' else (1 if x == 'Not Approved' else 2))
        df = df.sort_values(by=['sort_order', 'Site name']).drop(columns=['sort_order'])

    df['No.'] = range(1, len(df) + 1)

    styled = df.style.set_table_styles([
        {'selector': 'th', 'props': [('background-color', '#1b8a43'), ('color', 'white'), ('font-weight', 'bold'), ('text-align', 'center'), ('border', '1px solid black')]},
        {'selector': 'td', 'props': [('text-align', 'center'), ('border', '1px solid black'), ('font-size', '12pt')]}
    ]).hide(axis='index')

    dfi.export(styled, filename, table_conversion='chrome')
    return filename


def generate_overall_summary_image(overall_summary, filename="summary_table.png"):
    """បង្កើតរូបភាពតារាងសរុបរួម (សម្រាប់តែវេនល្ងាច)"""
    summary_rows = []
    total_target, total_app, total_not_app, total_remain = 0, 0, 0, 0

    for idx, team in enumerate(TARGET_TEAMS, 1):
        data_list = overall_summary[team]
        app = sum(1 for x in data_list if x['Result'] == "Approved")
        not_app = sum(1 for x in data_list if x['Result'] == "Not Approved")
        remain = sum(1 for x in data_list if x['Result'] != "Approved")
        target = len(data_list)
        pct = f"{round((app / target * 100))}%" if target > 0 else "0%"

        total_target += target
        total_app += app
        total_not_app += not_app
        total_remain += remain

        summary_rows.append({
            "No": idx,
            "Branch": team,
            "Target Site": target,
            "Approved": app,
            "Not Approved": not_app,
            "%": pct,
            "Remain": remain
        })

    overall_pct = f"{round((total_app / total_target * 100))}%" if total_target > 0 else "0%"
    summary_rows.append({
        "No": "",
        "Branch": "TOTAL",
        "Target Site": total_target,
        "Approved": total_app,
        "Not Approved": total_not_app,
        "%": overall_pct,
        "Remain": total_remain
    })

    df = pd.DataFrame(summary_rows)

    styled = df.style.set_table_styles([
        {'selector': 'th', 'props': [('background-color', '#2d9378'), ('color', 'white'), ('font-weight', 'bold'), ('text-align', 'center'), ('border', '1px solid black')]},
        {'selector': 'td', 'props': [('text-align', 'center'), ('border', '1px solid black'), ('font-size', '12pt')]}
    ]).hide(axis='index')

    dfi.export(styled, filename, table_conversion='chrome')
    return filename


def main():
    sh = get_spreadsheet()
    sheet_groups = load_sheet_groups(sh)
    
    # ពិនិត្យមើលម៉ោងកម្ពុជា (UTC+7)
    now_ict = datetime.utcnow() + timedelta(hours=7)
    today = now_ict.strftime("%d/%b/%Y")
    
    # បើម៉ោងតិចជាង ១២ ថ្ងៃត្រង់ ចាត់ទុកជា "វេនព្រឹក"
    is_morning = now_ict.hour < 12
    shift_title = "🌅 ផែនការការងារត្រូវអនុវត្ត (Morning Plan)" if is_morning else "🟢 របាយការណ៍លទ្ធផលការងារ (Evening Progress)"

    overall_summary = defaultdict(list)

    with TelegramClient(StringSession(SESSION_STRING), API_ID, API_HASH) as client:
        client.get_dialogs()

        for sheet_name in SHEET_NAMES:
            try:
                ws = sh.worksheet(sheet_name)
            except gspread.WorksheetNotFound:
                continue

            values = ws.get_all_values()
            header_idx = find_header_row(values)
            if header_idx is None:
                continue

            parsed_data = parse_sheet_data(values, header_idx)
            if not parsed_data:
                continue

            raw_chat_id = sheet_groups.get(sheet_name, REPORT_NOTIFY_GROUP_ID)
            chat_id = int(raw_chat_id)

            for team in TARGET_TEAMS:
                if team not in parsed_data:
                    continue

                rows_data = parsed_data[team]
                overall_summary[team].extend(rows_data)

                # បង្កើតរូបភាពតារាង
                img_path = generate_team_image(rows_data, sheet_name, is_morning)
                
                # បើវេនព្រឹកគ្មាន Site ដែល Not yet do ទេ វានឹងរំលងមិនផ្ញើ
                if img_path is None:
                    continue

                caption = f"**{shift_title}**\n📍 Sheet: **{sheet_name}** | Team: **{team}** ({today})"
                client.send_file(chat_id, img_path, caption=caption)
                print(f"[{sheet_name}] Sent Image for {team}")
                
                if os.path.exists(img_path):
                    os.remove(img_path)

                sleep_random_delay()

        # ផ្ញើតារាង Overall Summary តែនៅ "វេនល្ងាច" ប៉ុណ្ណោះ
        if not is_morning:
            print("📊 កំពុងរៀបចំផ្ញើរូបភាពតារាងសរុបរួម (Evening Summary)...")
            summary_img = generate_overall_summary_image(overall_summary)
            master_chat_id = int(REPORT_NOTIFY_GROUP_ID)
            client.send_file(master_chat_id, summary_img, caption=f"🏆 **Overall Summary Report - {today}**")
            
            if os.path.exists(summary_img):
                os.remove(summary_img)
            print("✅ ផ្ញើរូបភាពតារាងសរុបរួមបានជោគជ័យ!")


if __name__ == "__main__":
    main()
