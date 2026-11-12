from datetime import date, datetime, timezone, timedelta
import io
import json
import os
import re
import time

import dataframe_image as dfi
import pandas as pd
import requests
from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from telethon.sessions import StringSession
from telethon.sync import TelegramClient

SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID") or "1PmMSqfeBWhYJe5dMv3PrLOFKc2YmLYP8BdCvf9FyZX4"
MAIN_GROUP_ID = int(os.environ.get("NOTIFY_GROUP_ID") or "-1001853372580")

VALID_TEAMS = [f"CHA-T0{i}" for i in range(1, 8)]

TEAM_LEADERS = {
    "CHA-T01": "An Vanny",
    "CHA-T02": "Thoeung Chanthan",
    "CHA-T03": "Au Chhaira",
    "CHA-T04": "Chhorn Kim Chheang",
    "CHA-T05": "Mut Chanthy",
    "CHA-T06": "Rit Borin",
    "CHA-T07": "Sam Sophal",
}

TEAM_TELEGRAM_MENTIONS = {
    "CHA-T01": "@an_vanny9",
    "CHA-T02": "@Chanthan_team2",
    "CHA-T03": "@Au_chhaira",
    "CHA-T04": "@CHHORN_KIMCHHEANG",
    "CHA-T05": "@Mutchanthy",
    "CHA-T06": "@Borin_Team06",
    "CHA-T07": "@GIS_CHA_TEAMLEADER7",
}

def get_start_date(today):
    """ថ្ងៃច័ន្ទដំបូងនៃខែ (ស្វ័យប្រវត្តិ) មិនបាច់កំណត់ខែ។ START_DATE (env) ប្រើបានបើចង់បង្ខំ។"""
    override = (os.environ.get("START_DATE") or "").strip()
    if override:
        return datetime.strptime(override, "%Y-%m-%d").date()

    def first_monday(y, m):
        d = date(y, m, 1)
        return d + timedelta(days=(7 - d.weekday()) % 7)

    start = first_monday(today.year, today.month)
    if today < start:
        y, m = (today.year - 1, 12) if today.month == 1 else (today.year, today.month - 1)
        start = first_monday(y, m)
    return start

DEFAULT_WEEKS = 4
# Task ដែលមិនទាន់មាន Chat ID ក្នុង Sheet៖ ប្រើក្រុមរបស់ Task ផ្សេង ឬ Main Group បណ្តោះអាសន្ន
TASK_CHAT_OVERRIDE = {"B11": [-1003729578688]}   # B11 មាន 2 Chat ID ក្នុង Sheet → ប្រើតែមួយនេះ
TASK_SAME_GROUP_AS = {"E4": "B11"}
TASK_TO_MAIN_GROUP = ["E1", "B16"]
# Tab ក្នុង Sheet ដែលមិនទាន់មានក្នុង "Team chat IDs" (ត្រួតពិនិត្យថាមានការងារ Team ឬអត់)
UNMAPPED_TASK_TABS = ["B8", "B16", "C7", "C8", "C10", "E1", "E2", "E3", "E4", "E5", "E9", "E13", "E15"]
TASK_WEEKS = {"A2": 3}
REMINDER_THRESHOLD = 50    # Weekly reminder if team progress is below 50%
WARNING_THRESHOLD = 80     # Warning letter if team progress is below 80% for 3 consecutive weeks
WARNING_STATE_FILE = os.environ.get("WARNING_STATE_FILE", "weekly_warning_state.json")

COMMON_CAPTION_STYLE = {
    "selector": "caption",
    "props": [
        ("caption-side", "top"),
        ("font-size", "22px"),
        ("font-weight", "normal"),
        ("text-align", "center"),
        ("background-color", "#27AE60"),
        ("color", "black"),
        ("padding", "10px"),
        ("border", "1px solid black"),
        ("font-family", "serif"),
    ],
}


def task_weeks(task_code):
    return TASK_WEEKS.get(str(task_code).strip().upper(), DEFAULT_WEEKS)


def get_current_week(today):
    return (today - get_start_date(today)).days // 7 + 1


def normalize_team(name):
    match = re.search(r"(\d+)\s*$", str(name).strip())
    return f"CHA-T{int(match.group(1)):02d}" if match else None


_RAW_CSV = {}  # cache អត្ថបទ CSV ដើម្បីកុំទាញ Sheet ធំ (A2 មាន ~50,000 ជួរ) ពីរដង


def fetch_csv(sheet_name_or_gid):
    url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet={sheet_name_or_gid}"
    headers = {"User-Agent": "Mozilla/5.0"}
    res = None
    # Sheet ធំ (A2 ~50,000 ជួរ) ត្រូវការពេលលើស 15 វិនាទី → timeout វែង + ព្យាយាមម្តងទៀត
    for attempt in range(1, 4):
        try:
            res = requests.get(url, headers=headers, timeout=120)
            if res.status_code == 200:
                break
            print(f"⚠️ Sheet '{sheet_name_or_gid}' HTTP {res.status_code} (attempt {attempt}/3)")
        except requests.RequestException as e:
            print(f"⚠️ Request failed for sheet '{sheet_name_or_gid}' (attempt {attempt}/3): {e}")
            res = None
        time.sleep(3 * attempt)
    if res is None:
        return None
    if res.status_code == 200:
        res.encoding = "utf-8"
        _RAW_CSV[str(sheet_name_or_gid)] = res.text

    if res.status_code == 200:
        lines = res.text.splitlines()
        header_row_idx = 0
        header_keywords = {"site name", "team", "sheet", "chatid", "no", "no."}

        for idx, line in enumerate(lines[:15]):
            cells = [c.strip().strip('"').lower() for c in line.split(",")]
            if any(cell in header_keywords for cell in cells):
                header_row_idx = idx
                break

        try:
            df = pd.read_csv(io.StringIO(res.text), skiprows=header_row_idx, dtype=str).dropna(how="all")
        except (pd.errors.EmptyDataError, pd.errors.ParserError) as e:
            print(f"⚠️ Could not parse sheet '{sheet_name_or_gid}': {e}")
            return None

        df.columns = df.columns.astype(str).str.strip()

        rename_dict = {}
        for col in df.columns:
            c_lower = col.lower().strip()
            if "site name" in c_lower:
                rename_dict[col] = "Site name"
            elif c_lower in ["no", "no."]:
                rename_dict[col] = "No."
            elif c_lower == "team":
                rename_dict[col] = "Team"
            elif "group task" in c_lower:
                rename_dict[col] = "Group task"
            elif "result" in c_lower:
                rename_dict[col] = "Result"
            elif "serial" in c_lower:
                rename_dict[col] = "Serial"
            elif c_lower in ("name staff", "staff name", "staff"):
                rename_dict[col] = "Name staff"

        if rename_dict:
            df = df.rename(columns=rename_dict)

        if df.columns.duplicated().any():
            dup_cols = df.columns[df.columns.duplicated()].unique().tolist()
            print(f"⚠️ Duplicate columns found and removed: {dup_cols}")
            df = df.loc[:, ~df.columns.duplicated()]

        return df

    print(f"⚠️ Failed to fetch sheet '{sheet_name_or_gid}': HTTP {res.status_code}")
    return None


def get_task_title_from_sheet(sheet_name_or_gid):
    cached = _RAW_CSV.get(str(sheet_name_or_gid))
    if cached is not None:
        lines = cached.splitlines()
        if lines:
            first_line = lines[0].replace('"', "").strip()
            first_cell = first_line.split(",")[0].strip()
            if first_cell and not first_cell.isdigit():
                return first_cell
        return f"Task {sheet_name_or_gid}"

    url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet={sheet_name_or_gid}"
    headers = {"User-Agent": "Mozilla/5.0"}
    try:
        res = requests.get(url, headers=headers, timeout=120)
    except requests.RequestException as e:
        print(f"⚠️ Request failed for title lookup '{sheet_name_or_gid}': {e}")
        return f"Task {sheet_name_or_gid}"

    if res.status_code == 200:
        res.encoding = "utf-8"
        lines = res.text.splitlines()
        if lines:
            first_line = lines[0].replace('"', "").strip()
            first_cell = first_line.split(",")[0].strip()
            if first_cell and not first_cell.isdigit():
                return first_cell

    return f"Task {sheet_name_or_gid}"


def style_detail_table(df, title):
    styler = df.style.set_caption(title).set_table_styles([
        COMMON_CAPTION_STYLE,
        {"selector": "th", "props": [("background-color", "#369388"), ("color", "black"), ("font-weight", "bold"), ("text-align", "center"), ("border", "1px solid black"), ("padding", "6px")]},
        {"selector": "td", "props": [("text-align", "center"), ("border", "1px solid black"), ("padding", "5px")]},
    ])
    return styler


def style_pending_table(df, title):
    """តារាងការងារមិនទាន់ធ្វើ: ជួរពណ៌ខៀវស្រាលតាមគំរូ"""
    styler = style_detail_table(df, title)
    styler = styler.set_table_styles(
        [{"selector": "td", "props": [("background-color", "#BDD7EE"), ("white-space", "nowrap")]}], overwrite=False)
    # Remark / Last date record វែង → កាត់ជួរចុះបន្ទាត់ កុំឱ្យរូបទទឹងពេក (Telegram បង្រួមអក្សរតូច)
    long_cols = [c for c in df.columns if str(c).strip() in ("Remark", "Last date record")]
    if long_cols:
        styler = styler.set_properties(subset=long_cols, **{
            "white-space": "normal !important", "word-break": "break-word", "max-width": "260px", "min-width": "120px"})
    return styler


def style_task_summary(df, title):
    styler = df.style.set_caption(title).set_table_styles([
        COMMON_CAPTION_STYLE,
        {"selector": "th", "props": [("background-color", "#369388"), ("color", "black"), ("font-weight", "normal"), ("text-align", "center"), ("border", "1px solid black"), ("padding", "6px")]},
        {"selector": "td", "props": [("text-align", "center"), ("border", "1px solid black"), ("padding", "5px")]},
    ])
    def apply_row_styles(row):
        if row.name == 0:
            return ["color: red; font-style: italic; font-weight: bold;" for _ in row]
        styles = [""] * len(row)
        styles[5] = "background-color: #A2D9CE; font-weight: bold; font-style: italic;"
        return styles
    return styler.apply(apply_row_styles, axis=1)


def build_pending_list_text(df_single_team, team, task_title, show_serial=False):
    if "Site name" not in df_single_team.columns:
        return None
    has_serial = "Serial" in df_single_team.columns

    lines_data = []
    for _, row in df_single_team.iterrows():
        site = str(row.get("Site name", "")).strip()
        if not site:
            continue
        serial = ""
        if show_serial:
            # Sheet A2 មិនមានជួរ "Serial" ទេ: លេខ Serial ស្ថិតក្នុងជួរ "Q'ty task/Local task"
            serial = str(row.get("Serial", "") if has_serial else row.get("Q'ty task/Local task", "")).strip()
            if serial.endswith(".0"):
                serial = serial[:-2]
        if serial and serial.lower() not in ("nan", "none", "#n/a", "n/a"):
            lines_data.append(f"{team} {site} - {serial}")
        else:
            lines_data.append(f"{team} {site}")

    if not lines_data:
        return None
    return "\n".join(lines_data)


def load_team_chat_ids_from_mapping(mapping):
    """ទាញយក Chat ID ផ្ទាល់ខ្លួនរបស់ Team នីមួយៗ (Column A & B ពី Tab Team chat IDs)"""
    chats = {}
    if mapping is None or not {"Column A", "Column B"}.issubset(mapping.columns):
        return chats

    for _, row in mapping.dropna(subset=["Column A", "Column B"]).iterrows():
        team_raw = str(row["Column A"]).strip()
        team = normalize_team(team_raw)
        if team not in VALID_TEAMS:
            continue
        try:
            chats[team] = int(float(str(row["Column B"]).strip()))
        except (ValueError, TypeError):
            continue
    return chats


def split_sites_into_weeks(team_df, task_code):
    weeks_count = task_weeks(task_code)
    sites = team_df.copy()
    sites["_site_sort"] = sites["Site name"].fillna("").astype(str).str.strip().str.casefold()
    sites = sites[sites["_site_sort"] != ""].sort_values("_site_sort", kind="stable")
    sites = sites.drop(columns=["_site_sort"]).reset_index(drop=True)

    quotient, remainder = divmod(len(sites), weeks_count)
    result = {}
    start = 0
    for week in range(1, weeks_count + 1):
        size = quotient + (1 if week <= remainder else 0)
        result[week] = sites.iloc[start:start + size].copy()
        start += size
    return result


def approved_count(df):
    if df.empty or "Result" not in df:
        return 0
    return int(df["Result"].fillna("").astype(str).str.strip().str.casefold().eq("approved").sum())


def load_warning_state():
    try:
        with open(WARNING_STATE_FILE, encoding="utf-8") as handle:
            state = json.load(handle)
        return state if isinstance(state, dict) else {}
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return {}


def save_warning_state(state):
    temp_path = WARNING_STATE_FILE + ".tmp"
    with open(temp_path, "w", encoding="utf-8") as handle:
        json.dump(state, handle, ensure_ascii=False, indent=2)
    os.replace(temp_path, WARNING_STATE_FILE)


def pct(result, target):
    """ភាគរយ បង្គត់ធម្មតា (12.5 → 13) មិនមែន round() របស់ Python (12.5 → 12)"""
    return int(result * 100 / target + 0.5) if target else 0


def weekly_team_table(items):
    rows = []
    for index, item in enumerate(items, start=1):
        row = {"No.": index, "Task": item["Task"]}
        for week in range(1, DEFAULT_WEEKS + 1):
            data = item["Stats"].get(week)
            has = bool(data and data["Target"])   # សប្តាហ៍គ្មាន Site (Target 0) បង្ហាញ "—" មិនមែន 0%
            row[f"Week {week} Target"] = data["Target"] if has else "—"   # Target រួមបញ្ចូលការងាររុញពីសប្តាហ៍មុនរួចហើយ
            row[f"Week {week} Result"] = data["Result"] if has else "—"
            row[f"Week {week} %"] = f"{data['Percent']}%" if has else "—"
        rows.append(row)

    total = {"No.": "", "Task": "TOTAL"}
    for week in range(1, DEFAULT_WEEKS + 1):
        vals = [item["Stats"][week] for item in items if week in item["Stats"]]
        target = sum(value["Target"] for value in vals)
        result = sum(value["Result"] for value in vals)
        total[f"Week {week} Target"] = target if target else "—"
        total[f"Week {week} Result"] = result if target else "—"
        total[f"Week {week} %"] = f"{pct(result, target)}%" if target else "—"

    rows.append(total)
    return pd.DataFrame(rows).fillna("")


def current_week_team_progress(items, week_number):
    target = 0
    result = 0
    for item in items:
        data = item["Stats"].get(week_number)
        if data:
            target += data["Target"]
            result += data["Result"]
    percent = pct(result, target)
    return target, result, percent


def make_warning_letter_docx(teams, progress, today, week_number, reasons=None):
    reasons = reasons or {}
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.55)
    section.bottom_margin = Inches(0.55)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)

    doc.styles["Normal"].font.name = "Times New Roman"
    doc.styles["Normal"].font.size = Pt(10)

    header = doc.add_table(rows=1, cols=2)
    def set_cell_text(cell, text, bold=False, size=9):
        cell.text = ""
        paragraph = cell.paragraphs[0]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.add_run(str(text))
        run.bold = bold
        run.font.name = "Times New Roman"
        run.font.size = Pt(size)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    set_cell_text(header.cell(0, 0), "G.I SERVICES (CAMBODIA) CO., LTD.\nKOM PONG CHAM BRANCH", True, 10)
    set_cell_text(header.cell(0, 1), "KINGDOM OF CAMBODIA\nNATION - RELIGION - KING", True, 9)
    header.cell(0, 0).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
    header.cell(0, 1).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    approval = doc.add_paragraph()
    approval.paragraph_format.space_before = Pt(10)
    approval.add_run(f"APPROVAL\nDate: {today.strftime('%d/%m/%Y')}\nDIRECTOR").bold = True

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.space_before = Pt(14)
    title_run = title.add_run("DRAFT WARNING LETTER")
    title_run.bold = True
    title_run.font.size = Pt(16)

    subject = doc.add_paragraph(
        f"Subject: Draft Warning Letter for Team Leaders whose teams have remained "
        f"below {WARNING_THRESHOLD}% of the weekly target for 3 consecutive weeks, "
        f"or have no result at all for the whole week."
    )
    subject.runs[0].bold = True

    doc.add_paragraph("To: Board of Directors, Technical Center")
    
    table = doc.add_table(rows=1, cols=6)
    table.style = "Table Grid"
    headers = ["No.", "Team", "Team Leader", "Position", "Reason for Warning", "Proposed Action"]
    for cell, label in zip(table.rows[0].cells, headers):
        set_cell_text(cell, label, True, 10)

    for index, team in enumerate(teams, start=1):
        target, result, percent = progress[team]
        values = [
            str(index), team, TEAM_LEADERS[team], "Team Leader",
            reasons.get(team) or f"Remained below {WARNING_THRESHOLD}% for 3 consecutive weeks (Week {week_number}). Total: {result}/{target} ({percent}%).",
            "Proposed: Deduct 1 KI allowance (subject to Director approval)"
        ]
        for cell, value in zip(table.add_row().cells, values):
            set_cell_text(cell, value, size=9)

    filename = f"warning_letter_draft_week_{week_number}_{today.isoformat()}.docx"
    doc.save(filename)
    return filename


def send_with_caption(client, chat_id, path, caption):
    """Caption លើស 1024 តួអក្សរ → ផ្ញើ title ជា caption ហើយបញ្ជីនៅសល់ជាសារដាច់ដោយឡែក។
    មាន try/except ដូច្នេះមួយបរាជ័យមិនបញ្ឈប់ស្គ្រីបទាំងមូល។"""
    try:
        if len(caption) <= 1000:
            client.send_file(chat_id, path, caption=caption)
        else:
            head, _, body = caption.partition("\n\n")
            client.send_file(chat_id, path, caption=head[:1000])
            lines, chunk = body.split("\n"), ""
            for line in lines:
                if len(chunk) + len(line) + 1 > 3500:
                    client.send_message(chat_id, chunk, parse_mode=None)
                    chunk = ""
                chunk += line + "\n"
            if chunk.strip():
                client.send_message(chat_id, chunk, parse_mode=None)
        time.sleep(1)
        return True
    except Exception as e:
        print(f"⚠️ Send failed for chat {chat_id}: {e}")
        return False


def main():
    api_id_raw = os.environ.get("TELEGRAM_API_ID")
    api_hash = os.environ.get("TELEGRAM_API_HASH")
    session_str = os.environ.get("TELEGRAM_SESSION")

    if not api_id_raw or not api_hash or not session_str:
        raise ValueError("Telegram API credentials are missing in environment/secrets")

    api_id = int(api_id_raw)

    df_mapping = fetch_csv("Team%20chat%20IDs")
    task_chat_ids = {}   # code -> [chat_id, ...]  (បើ Task មួយមានច្រើន Chat ID ផ្ញើទៅទាំងអស់)

    if df_mapping is not None and "Sheet" in df_mapping.columns:
        df_clean = df_mapping.dropna(subset=["Sheet", "ChatID"])
        for _, row in df_clean.iterrows():
            code = str(row["Sheet"]).strip()
            try:
                cid = int(float(str(row["ChatID"]).strip()))
            except ValueError:
                continue
            ids = task_chat_ids.setdefault(code, [])
            if cid not in ids:
                ids.append(cid)
        for code, ids in task_chat_ids.items():
            if len(ids) > 1:
                print(f"⚠️ Task {code} មាន Chat ID {len(ids)} ផ្សេងគ្នាក្នុង 'Team chat IDs' → ផ្ញើទៅទាំងអស់: {ids}")

    for code, ids in TASK_CHAT_OVERRIDE.items():          # បង្ខំ Chat ID (មានអាទិភាពលើ Sheet)
        task_chat_ids[code] = list(ids)

    # Task ដែលមិនទាន់មាន Chat ID ក្នុង Sheet (បើ Sheet មានជួរសម្រាប់ Task នោះ វាមានអាទិភាពជាង)
    for code, src in TASK_SAME_GROUP_AS.items():          # E4 → ក្រុមដូច B11
        if code not in task_chat_ids and src in task_chat_ids:
            task_chat_ids[code] = list(task_chat_ids[src])
    for code in TASK_TO_MAIN_GROUP:                       # E1, B16 → Main Group សិន
        if code not in task_chat_ids:
            task_chat_ids[code] = [MAIN_GROUP_ID]

    if not task_chat_ids:
        print("⚠️ WARNING: task_chat_ids is empty.")
        return

    # ទាញយក Chat ID សម្រាប់ Team នីមួយៗពី Column A និង Column B
    team_chat_mapping = load_team_chat_ids_from_mapping(df_mapping)

    cambodia_tz = timezone(timedelta(hours=7))
    now = datetime.now(cambodia_tz)
    today = now.date()
    shift_name = os.environ.get("SHIFT_NAME") or (
        "Morning Shift" if (os.environ.get("GITHUB_EVENT_NAME") == "schedule" or os.environ.get("RUN_SOURCE") == "cron")
        else "Afternoon Shift" if now.hour < 18 else "Evening Shift")
    current_week = get_current_week(today)
    abs_week = (today.toordinal() - 1) // 7  # លេខសប្តាហ៍ពិត (ប្រើរាប់ជាប់គ្នាឆ្លងខែ)
    warning_state = load_warning_state()

    # ===== ផ្ញើតាមម៉ោងថេរ (GitHub Actions free ចាប់ផ្តើមយឺតមិនទៀង) =====
    # ស្វ័យប្រវត្តិ = schedule របស់ GitHub ឬ ហៅពី cron ខាងក្រៅ (RUN_SOURCE=cron)។ ដំណើរការដោយដៃ = មិនមែនទាំងពីរ។
    is_scheduled = (os.environ.get("GITHUB_EVENT_NAME") == "schedule"
                    or os.environ.get("RUN_SOURCE") == "cron")
    print(f"ℹ️ Run started {now:%Y-%m-%d %H:%M:%S} (Cambodia) | event={os.environ.get('GITHUB_EVENT_NAME')} "
          f"| source={os.environ.get('RUN_SOURCE') or '-'}")
    # Cron បម្រុង: បើថ្ងៃនេះផ្ញើរួចហើយ ឱ្យរំលង (កុំផ្ញើស្ទួន)
    if is_scheduled and warning_state.get("_last_run") == today.isoformat() and os.environ.get("FORCE_RESEND") != "1":
        print(f"ℹ️ Already sent today ({today}) — skipping duplicate scheduled run")
        return
    # ដំណើរការមុនម៉ោង រួចរង់ចាំដល់ SEND_AT (ឧ. 08:00)។ ដំណើរការដោយដៃ (workflow_dispatch) មិនរង់ចាំទេ។
    send_at = (os.environ.get("SEND_AT") or "").strip()
    if is_scheduled and send_at:
        hh, mm = (int(x) for x in send_at.split(":"))
        wait_s = (now.replace(hour=hh, minute=mm, second=0, microsecond=0) - now).total_seconds()
        if 0 < wait_s <= 90 * 60:
            print(f"ℹ️ Waiting {int(wait_s)}s until {send_at} (Cambodia time) before sending")
            time.sleep(wait_s)
        else:
            print(f"ℹ️ Not waiting (now {now:%H:%M}, SEND_AT {send_at}) — sending immediately")

    team_rows = {team: [] for team in VALID_TEAMS}

    delivery = {}        # code -> {"groups": [...], "sent": n, "failed": n, "no_rows": bool, "fetch_failed": bool}
    unmapped_with_rows = []

    with TelegramClient(StringSession(session_str), api_id, api_hash) as client:
        # StringSession ថ្មីរាល់ដំណើរការ មិនស្គាល់ Group ណាទេ → ទាញបញ្ជី Dialogs ជាមុន ទើបផ្ញើចូល Group បាន
        try:
            client.get_dialogs()
        except Exception as e:
            print(f"⚠️ get_dialogs failed: {e}")
        reachable = {}
        def can_reach(cid):
            if cid not in reachable:
                try:
                    client.get_input_entity(cid)
                    reachable[cid] = True
                except Exception as e:
                    print(f"❌ រក Group {cid} មិនឃើញ (Chat ID ខុស ឬគណនីនេះមិនមែនជាសមាជិក): {e}")
                    reachable[cid] = False
            return reachable[cid]

        # ១. ស្កេនទាញយកទិន្នន័យពី Sheet ការងារនីមួយៗ (A1, A2, ...)
        # A2 ធ្វើមុនគេ (sorted មានស្ថិរភាព: Task ផ្សេងរក្សាលំដាប់ដើម)
        for task_code, chat_ids in sorted(task_chat_ids.items(), key=lambda kv: kv[0].strip().upper() != "A2"):
            info = delivery.setdefault(task_code, {"groups": list(chat_ids), "sent": 0, "failed": 0,
                                                   "no_rows": False, "fetch_failed": False})
            df_task = fetch_csv(task_code)
            sheet_sub_title = get_task_title_from_sheet(task_code)
            task_title = f"{task_code}. {sheet_sub_title}"
            if df_task is None or df_task.empty:
                info["fetch_failed"] = df_task is None
                print(f"⚠️ Task {task_code}: មិនមានទិន្នន័យ / ទាញ Sheet មិនបាន")
                continue

            cols_to_show = ["Group task", "Branch", "Site name", "Q'ty task/Local task", "Result", "Remark", "Last date record", "History Task", "Team"]
            canon = {c.lower(): c for c in cols_to_show}
            df_task = df_task.rename(columns={c: canon[c.lower().strip()] for c in df_task.columns if c.lower().strip() in canon})
            available_cols = [c for c in cols_to_show if c in df_task.columns]

            if available_cols and "Team" in df_task.columns:
                df_detail = df_task[available_cols].copy()
                df_detail["Team"] = df_detail["Team"].astype(str).str.strip()
                df_detail = df_detail[df_detail["Team"].isin(VALID_TEAMS)]

                if "Site name" in df_detail.columns:
                    df_detail = df_detail[df_detail["Site name"].notna() & (~df_detail["Site name"].astype(str).str.strip().str.lower().isin(["", "nan", "none", "#n/a", "n/a"]))]

                # Run ស្វ័យប្រវត្តិ (ព្រឹក) = យកតែការងារមិនទាន់ Approved; Run ដោយដៃ = យកទិន្នន័យទាំងអស់ (រួមទាំង Approved)
                if is_scheduled and "Result" in df_detail.columns:
                    df_detail = df_detail[df_detail["Result"].astype(str).str.strip().str.lower() != "approved"]

                df_detail = df_detail.fillna("").replace(to_replace=r"^(?i:nan|none|#n/a|n/a)$", value="", regex=True)

                # ផ្ញើការងារមិនទាន់ Approved ទៅកាន់គ្រុប Task នីមួយៗ (Task Chat ID)
                for team in VALID_TEAMS:
                    df_single_team = df_detail[df_detail["Team"] == team].copy()
                    if not df_single_team.empty:
                        styled_detail = style_pending_table(df_single_team, f"{task_title} ({team})")
                        img_path = f"detail_{task_code}_{team}.png"
                        dfi.export(styled_detail.hide(axis="index"), img_path, max_rows=-1)

                        pending_text = build_pending_list_text(df_single_team, team, task_title, show_serial=(str(task_code).strip().upper() == "A2"))
                        full_caption = (f"ការងារក្រុមមិនទាន់ធ្វើ {team} ({task_title}) - {shift_name}" if is_scheduled
                                        else f"ទិន្នន័យការងារទាំងអស់ {team} ({task_title}){(' - ' + shift_name) if shift_name else ''}")
                        if pending_text:
                            full_caption += f"\n\n{pending_text}"

                        for chat_id in chat_ids:
                            if can_reach(chat_id) and send_with_caption(client, chat_id, img_path, full_caption):
                                info["sent"] += 1
                            else:
                                info["failed"] += 1
                        if os.path.exists(img_path):
                            os.remove(img_path)

            # ប្រមូលទិន្នន័យដើម្បីគណនាសប្តាហ៍ (Weekly Split)
            for team in VALID_TEAMS:
                if "Team" in df_task.columns and "Result" in df_task.columns:
                    df_clean_task = df_task.dropna(subset=["Team"]).copy()
                    df_clean_task["Team"] = df_clean_task["Team"].astype(str).str.strip()
                    df_team = df_clean_task[df_clean_task["Team"] == team].copy()

                    if not df_team.empty and "Site name" in df_task.columns:
                        groups = split_sites_into_weeks(df_team, task_code)
                        stats = {}
                        # Result បំពេញតាមលំដាប់សប្តាហ៍: Site ដែល Approved បំពេញ Target សប្តាហ៍ទី 1 មុន
                        # ទើបលើសទៅសប្តាហ៍ទី 2, 3... (មិនរាប់តាមទីតាំង Site ក្នុងបញ្ជីអក្សរក្រម)
                        remaining = sum(approved_count(wdf) for wdf in groups.values())
                        # ការងារសល់ពីសប្តាហ៍ចាស់ (សប្តាហ៍ដែលបានចប់ហើយ) រុញបូកបន្ថែមក្នុង Target សប្តាហ៍បន្ទាប់
                        shortfall = 0
                        for week, week_df in groups.items():
                            carry = shortfall if (week - 1) < current_week else 0
                            w_target = len(week_df) + carry
                            w_result = min(remaining, w_target)
                            remaining -= w_result
                            shortfall = w_target - w_result
                            w_percent = pct(w_result, w_target)
                            stats[week] = {"Target": w_target, "Result": w_result, "Percent": w_percent, "Carry": carry}
                        team_rows[team].append({"Task": task_code, "Weeks": task_weeks(task_code), "Stats": stats})

        # ១.១ Task ដែលមានការងារ Team តែគ្មាន Chat ID ក្នុង 'Team chat IDs' (មិនដែលត្រូវបានផ្ញើ)
        extra = [c.strip() for c in (os.environ.get("CHECK_UNMAPPED_TASKS") or ",".join(UNMAPPED_TASK_TABS)).split(",") if c.strip()]
        for code in extra:
            if code in task_chat_ids:
                continue
            df_x = fetch_csv(code)
            if df_x is None or df_x.empty or "Team" not in df_x.columns or "Site name" not in df_x.columns:
                continue
            ok = df_x["Team"].astype(str).str.strip().isin(VALID_TEAMS) & df_x["Site name"].notna() \
                & ~df_x["Site name"].astype(str).str.strip().str.lower().isin(["", "nan", "none", "#n/a", "n/a"])
            if ok.sum():
                unmapped_with_rows.append((code, int(ok.sum())))

        problems = [f"❌ {c}: ផ្ញើបាន {d['sent']} / បរាជ័យ {d['failed']} (Group {d['groups']})"
                    for c, d in delivery.items() if d["failed"]]
        problems += [f"⚠️ {c}: ទាញ Sheet មិនបាន" for c, d in delivery.items() if d["fetch_failed"]]
        problems += [f"⚠️ {c}: មាន {n} ការងារ Team តែគ្មាន Chat ID ក្នុង 'Team chat IDs'" for c, n in unmapped_with_rows]
        print("===== Delivery summary =====")
        for c, d in delivery.items():
            print(f"  {c}: groups={d['groups']} sent={d['sent']} failed={d['failed']} fetch_failed={d['fetch_failed']}")
        if problems and os.environ.get("SEND_DELIVERY_REPORT", "1") != "0":
            try:
                client.send_message(MAIN_GROUP_ID, "📋 របាយការណ៍ផ្ញើ Task ទៅ Group — ចំណុចត្រូវពិនិត្យ\n" + "\n".join(problems), parse_mode=None)
            except Exception as e:
                print(f"⚠️ Could not send delivery summary: {e}")

        # ២. ផ្ញើតារាងលទ្ធផលសប្តាហ៍ និងរំលឹក (Loop ផ្ញើដាច់ដោយឡែកទៅកាន់គ្រុប Team នីមួយៗ ម្តងមួយធីម)
        for team in VALID_TEAMS:
            items = team_rows.get(team, [])
            team_chat_id = team_chat_mapping.get(team)  # យក Chat ID ផ្ទាល់ខ្លួនរបស់ Team ពី Column A & B

            if not items or not team_chat_id:
                print(f"⚠️ Missing data or chat ID for {team} — skipped weekly report.")
                continue

            # បង្កើតនិងផ្ញើរូបភាពតារាង Weekly Report
            title = f"របាយការណ៍ប្រចាំសប្តាហ៍ {team} — គោលដៅ លទ្ធផល និងភាគរយ តាមឈ្មោះ Site"
            image_path = f"team_summary_{team}.png"
            dfi.export(style_detail_table(weekly_team_table(items), title).hide(axis="index"), image_path, max_rows=-1)
            send_with_caption(client, team_chat_id, image_path, f"{title}")

            # ពិនិត្យភាគរយសប្តាហ៍បច្ចុប្បន្ន ប្រសិនបើក្រោម 50% ផ្ញើសារ Tag រំលឹក
            target, result, percent = current_week_team_progress(items, current_week)
            mention = TEAM_TELEGRAM_MENTIONS.get(team, TEAM_LEADERS[team])

            if target and percent < REMINDER_THRESHOLD:
                reminder_msg = (
                    f"⚠️ **រំលឹកប្រចាំសប្តាហ៍**\n"
                    f"Dear bong {mention} ({team} — {TEAM_LEADERS[team]})\n"
                    f"លទ្ធផលសប្តាហ៍នេះ៖ {result}/{target} ({percent}%)\n"
                    f"🔻 ភាគរយនៅក្រោម {REMINDER_THRESHOLD}% សូមបងបញ្ជាក់ពីមូលហេតុ និងវិធានការកែលម្អជាបន្ទាន់។"
                )
                try:
                    client.send_message(team_chat_id, reminder_msg)
                except Exception as e:
                    print(f"⚠️ Reminder to {team} failed: {e}")

        # ៣. ពិនិត្យ Warning Letter ធ្លាក់ចុះក្រោម 80% ជាប់គ្នា 3 សប្តាហ៍ ផ្ញើចូល Main Group
        candidates = []
        cand_prog = {}
        for team in VALID_TEAMS:
            items = team_rows.get(team, [])
            target, result, percent = current_week_team_progress(items, current_week)
            state = warning_state.setdefault(team, {"low_weeks": 0, "last_counted_week": 0, "draft_sent": False})

            # រាប់ចាប់ពីថ្ងៃសុក្រ (ចៀសវាងថ្ងៃច័ន្ទដែលភាគរយទាបដោយធម្មជាតិ)
            if target and today.weekday() >= 4:
                if percent < WARNING_THRESHOLD:
                    last = int(state.get("last_counted_week", 0))
                    if last < abs_week:
                        # ជាប់គ្នាពិតប្រាកដ: បើរំលងសប្តាហ៍ → ចាប់ផ្តើមរាប់ពី 1 វិញ
                        state["low_weeks"] = state["low_weeks"] + 1 if last == abs_week - 1 else 1
                        state["last_counted_week"] = abs_week
                    if state["low_weeks"] >= 3 and not state.get("draft_sent", False):
                        candidates.append(team)
                        cand_prog[team] = (target, result, percent)
                else:
                    state["low_weeks"] = 0
                    state["draft_sent"] = False

        # ៣.១ ទាំងសប្តាហ៍គ្មានលទ្ធផលសោះ (0%) → ចេញលិខិតព្រមានមេធីមភ្លាម (ពិនិត្យចាប់ពីថ្ងៃសុក្រ ម្តងក្នុងមួយសប្តាហ៍)
        reasons = {}
        zero_teams = []
        for team in VALID_TEAMS:
            items = team_rows.get(team, [])
            target, result, percent = current_week_team_progress(items, current_week)
            state = warning_state.setdefault(team, {"low_weeks": 0, "last_counted_week": 0, "draft_sent": False})
            if target and result == 0 and today.weekday() >= 4 and state.get("zero_sent_week") != abs_week:
                zero_teams.append(team)
                reasons[team] = f"No result at all for the whole of Week {current_week} (0/{target}, 0%)."
                cand_prog[team] = (target, result, percent)

        letter_teams = list(dict.fromkeys(candidates + zero_teams))
        if letter_teams:
            for t in candidates:
                reasons.setdefault(t, None)
            path = make_warning_letter_docx(letter_teams, cand_prog, today, current_week, reasons)
            if send_with_caption(client, MAIN_GROUP_ID, path,
                                 f"⚠️ សេចក្តីព្រាងលិខិតព្រមាន (ក្រុមក្រោម {WARNING_THRESHOLD}% ជាប់គ្នា 3 សប្តាហ៍ ឬគ្មានលទ្ធផលទាំងសប្តាហ៍)"):
                for team in candidates:
                    warning_state[team]["draft_sent"] = True
                for team in zero_teams:
                    warning_state[team]["zero_sent_week"] = abs_week

    if is_scheduled:
        warning_state["_last_run"] = today.isoformat()
    save_warning_state(warning_state)

if __name__ == "__main__":
    main()
