from datetime import date, datetime, timezone, timedelta
import csv
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


def fetch_csv(sheet_name_or_gid, spreadsheet_id=None):
    url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id or SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet={sheet_name_or_gid}"
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
        if not spreadsheet_id:
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


WARNING_TEMPLATE = os.environ.get("WARNING_TEMPLATE", "warning_template.docx")


def _set_par(par, text):
    """ប្តូរអត្ថបទប្រយោគ ដោយរក្សាទម្រង់ (font/bold) របស់ run ទី 1 (ឬ run ដែលមិនមែន bold-label)"""
    runs = par.runs
    if not runs:
        par.add_run(text)
        return
    keep = runs[-1] if len(runs) > 1 and runs[0].bold and not runs[-1].bold else runs[0]
    keep.text = text
    for r in runs:
        if r is not keep:
            r.text = ""


def make_warning_letter_from_template(teams, progress, today, week_number, reasons, monthly=None):
    """monthly = ឈ្មោះខែ (ឧ. 'November 2026') → លិខិតប្រចាំខែ (ការងារមិនចប់ក្នុងខែ រុញទៅខែក្រោយ)"""
    from copy import deepcopy
    doc = Document(WARNING_TEMPLATE)
    names = ", ".join(teams)
    for par in doc.paragraphs:
        t = par.text
        if t.startswith("Subject:"):
            _set_par(par, (f"Subject: Warning letter for Team Leaders not finishing work within the month and carrying work over to the next month ({monthly})."
                           if monthly else
                           f"Subject: Warning letter for Team Leaders whose teams remained below {WARNING_THRESHOLD}% of the weekly target "
                           f"for 3 consecutive weeks, or had no result at all for the whole week (Week {week_number})."))
            par.runs[0].bold = True
        elif t.startswith("Based on the failure"):
            _set_par(par, (f"Based on the failure to finish work within the month and leaving work to carry over to the next month ({monthly})."
                           if monthly else f"Based on the weekly target progress of Week {week_number} ({today.strftime('%d/%m/%Y')})."))
        elif t.startswith("Kom Pong Cham Branch would like to issue"):
            _set_par(par, f"Kom Pong Cham Branch would like to issue a warning to the Team Leaders of: {names}.")
        elif t.startswith("1. Details of the warning"):
            _set_par(par, ("1. Details of the warning to Team Leaders who did not finish work within the month and carried work over."
                           if monthly else "1. Details of the warning to Team Leaders whose teams did not reach the weekly target."))
        elif t.startswith("Remind Team Leaders"):
            _set_par(par, f"Remind Team Leaders {names} for the first time, based on the root causes detailed above.")
    table = next(t for t in doc.tables if len(t.columns) == 6)
    proto = deepcopy(table.rows[1]._tr)
    for row in list(table.rows)[1:]:
        table._tbl.remove(row._tr)
    for i, team in enumerate(teams, start=1):
        target, result, percent = progress[team]
        table._tbl.append(deepcopy(proto))
        cells = table.rows[-1].cells
        reason = reasons.get(team) or (f"Remained below {WARNING_THRESHOLD}% for 3 consecutive weeks (Week {week_number}). "
                                       f"Total: {result}/{target} ({percent}%).")
        for cell, val in zip(cells, [str(i), team, TEAM_LEADERS[team], "Team Leader", reason, "Deduct 1KI"]):
            _set_par(cell.paragraphs[0], val)
    filename = (f"warning_letter_monthly_{today:%Y-%m}.docx" if monthly
                else f"warning_letter_draft_week_{week_number}_{today.isoformat()}.docx")
    doc.save(filename)
    return filename


def make_monthly_letter(team_left, today):
    """team_left = {team: {task: ចំនួន Site មិនទាន់ Approved}} → លិខិតប្រចាំខែ (ទម្រង់ដូចលិខិតខែតុលា)"""
    teams = [t for t in VALID_TEAMS if team_left.get(t)]
    reasons = {}
    for t in teams:
        tasks = ", ".join(f"{k} ({v} site{'s' if v > 1 else ''})" for k, v in sorted(team_left[t].items()))
        reasons[t] = f"Did not finish the work by the deadline, and left work {tasks} to carry over to the next month."
    month_label = today.strftime("%B %Y")
    if not os.path.exists(WARNING_TEMPLATE):
        return None
    return make_warning_letter_from_template(teams, {t: (0, 0, 0) for t in teams}, today, 0, reasons, monthly=month_label)


def make_warning_letter_docx(teams, progress, today, week_number, reasons=None):
    reasons = reasons or {}
    if os.path.exists(WARNING_TEMPLATE):
        try:
            return make_warning_letter_from_template(teams, progress, today, week_number, reasons)
        except Exception as e:
            print(f"⚠️ Template letter failed ({e}) — using plain layout")
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


# ===== ម៉ោងម្តង: រកការងារ Not Approved → រកមូលហេតុ → បើគ្មាន Tag អ្នកកាន់ការងារសួរ → ផ្ញើចូលគ្រុប Team =====
# Telegram username របស់ staff (ឈ្មោះក្នុងជួរ "Name staff" → @username)។ បើគ្មាន សរសេរតែឈ្មោះ (Tag មិនបាន)។
STAFF_TELEGRAM = {
    "Oum Sopheaktra": "@sopheaktraoum",
    "Bou Soeu": "+855977689987",          # លេខទូរស័ព្ទ (គ្មាន @username)
    "Trol Kea": "@TROL_KEA",
    "Nhoeuk Pech": "@Nhoeuk_Pech",
    "Ing Sokkhemarak": "@khemarak9",
    "Chit Kanas": "@Kanasc1",
    # "Vo Hong Lam": "@username", "Oeun Samart": "@username", "Teth": "@username",
}
HOURLY_FROM, HOURLY_TO = 8, 17      # ម៉ោងកម្ពុជា (រួមទាំងម៉ោង 17)
# Sheet មេ (Master report power) មានមូលហេតុក្នុងជួរ Remark — ត្រូវចែករំលែក "Anyone with the link: Viewer"
# ម្ចាស់ Sheet មេមិនឱ្យចែករំលែក → ចម្លងទិន្នន័យចូល Sheet Report Power ខ្លួនឯងតាម IMPORTRANGE (អានតែមួយទិសដៅ មិនប៉ះពាល់ Sheet មេ)
#   Tab "MasterSUM" = IMPORTRANGE(Sheet មេ, "SUM!A1:N80") → ឈ្មោះអ្នកកាន់ការងារ
#   Tab "M_B3", "M_A1" ... (ជាជម្រើស) = IMPORTRANGE(Sheet មេ, "B3!A1:Z2000") → ជួរ Remark មូលហេតុ។ ដាក់ឈ្មោះ Task ក្នុង MASTER_REASON_TABS
MASTER_SPREADSHEET_ID = os.environ.get("MASTER_SPREADSHEET_ID") or SPREADSHEET_ID
MASTER_SUM_TAB = os.environ.get("MASTER_SUM_TAB") or ("SUM" if MASTER_SPREADSHEET_ID != SPREADSHEET_ID else "MasterSUM")
MASTER_REASON_TABS = [t.strip().upper() for t in (os.environ.get("MASTER_REASON_TABS") or "").split(",") if t.strip()]
REASON_COLUMNS = ["Remark", "Remark ELE", "History Task"]
# អ្នកកាន់ការងារដែលហៅ "អូន" (ក្រៅពីនេះហៅ "បង")។ ឧ. {"Teth"}
# អ្នកកាន់ការងារដែលមិនត្រូវសួរសោះ (ទាំងឆាតផ្ទាល់ និងក្នុងក្រុម)
ASK_FROM, ASK_TO = 9, 10      # សួរអ្នកកាន់ការងារចន្លោះម៉ោង 9:00–10:00
NEVER_ASK = {"Vo Hong Lam", "Oeun Samart", "Teth"}
CALL_AS_ADIK = {"Trol Kea", "Ing Sokkhemarak", "Oum Sopheaktra"}
JUNK_REASONS = {"-", ",", "0", "#n/a", "new target", "ok"}
_EMPTY = ("", "nan", "none", "#n/a", "n/a")


def _clean(v):
    v = str(v).strip()
    return "" if v.lower() in _EMPTY else v


def staff_tag(name):
    name = _clean(name)
    if not name:
        return "(មិនមានឈ្មោះអ្នកកាន់ការងារ)"
    try:
        extra = json.loads(os.environ.get("STAFF_TELEGRAM_JSON") or "{}")
    except json.JSONDecodeError:
        extra = {}
    handles = {k.strip().casefold(): v for k, v in {**STAFF_TELEGRAM, **extra}.items()}
    handle = handles.get(name.casefold())
    return f"{handle} ({name})" if handle else name


def _reason_from(row, cols):
    for col in cols:
        v = _clean(row.get(col, "")) if col in row.index else ""
        v = v.lstrip(" ,;.")
        if v and v.lower() not in JUNK_REASONS:
            return v
    return ""


TASK_HOLDERS_DEFAULT = {'A1': 'Oum Sopheaktra', 'A2': 'Bou Soeu', 'A3': 'Bou Soeu', 'A4': 'Chit Kanas', 'A5': 'Trol Kea', 'A6': 'Trol Kea', 'A7': 'Chit Kanas', 'A8': 'Oum Sopheaktra', 'B1': 'Oum Sopheaktra', 'B2': 'Oum Sopheaktra', 'B3': 'Bou Soeu', 'B4': 'Trol Kea', 'B5': 'Trol Kea', 'B6': 'Trol Kea', 'B7': 'Nhoeuk Pech', 'B8': 'Oum Sopheaktra', 'B9': 'Trol Kea', 'B10': 'Ing Sokkhemarak', 'B11': 'Bou Soeu', 'B12': 'Nhoeuk Pech', 'B13': 'Oum Sopheaktra', 'B14': 'Nhoeuk Pech', 'B15': 'Nhoeuk Pech', 'B16': 'Chit Kanas', 'B17': 'Trol Kea', 'B18': 'Nhoeuk Pech', 'C1': 'Vo Hong Lam', 'C2': 'Trol Kea', 'C3': 'Trol Kea', 'C4': 'Nhoeuk Pech', 'C5': 'Vo Hong Lam', 'C6': 'Vo Hong Lam', 'C7': 'Vo Hong Lam', 'C8': 'Vo Hong Lam', 'C9': 'Nhoeuk Pech', 'C10': 'Oum Sopheaktra', 'C11': 'Vo Hong Lam', 'W1': 'Oeun Samart', 'W2': 'Chit Kanas', 'W3': 'Chit Kanas', 'W4': 'Nhoeuk Pech', 'E1': 'Oum Sopheaktra', 'E2': 'Bou Soeu', 'E3': 'Teth', 'E4': 'Bou Soeu', 'E5': 'Trol Kea', 'E16': 'Teth'}   # ពី Sheet មេ Tab SUM ជួរ J (ប្រើបើគ្មាន Tab MasterSUM)


def task_holders():
    """{Task → ឈ្មោះអ្នកកាន់ការងារ} ពី Sheet មេ Tab SUM (ជួរ "Annex sheet" និង "Staff 1")"""
    df = fetch_csv(MASTER_SUM_TAB, MASTER_SPREADSHEET_ID if MASTER_SPREADSHEET_ID != SPREADSHEET_ID else None)
    out = dict(TASK_HOLDERS_DEFAULT)
    if df is None or df.empty:
        return out
    annex = next((c for c in df.columns if str(c).strip().lower().startswith("annex")), None)
    staff = next((c for c in df.columns if str(c).strip().lower().startswith("staff 1")), None)
    if not annex or not staff:
        print(f"⚠️ SUM: រកជួរ Annex sheet / Staff 1 មិនឃើញ ({list(df.columns)[:12]})")
        return out
    for _, r in df.iterrows():
        code, name = _clean(r[annex]).upper(), _clean(r[staff])
        if code and name:
            out[code] = name
    return out


def master_reasons(code):
    """{(site, serial) / site → មូលហេតុ} ពី Tab ដូចគ្នាក្នុង Sheet មេ"""
    out = {}
    if MASTER_SPREADSHEET_ID == SPREADSHEET_ID:       # Sheet មេចម្លងចូល Sheet ខ្លួនឯង → Tab M_<Task> (បើបានបង្កើត)
        if str(code).strip().upper() not in MASTER_REASON_TABS:
            return out
        df = fetch_csv(f"M_{code}")
    else:
        df = fetch_csv(code, MASTER_SPREADSHEET_ID)
    if df is None or df.empty or "Site name" not in df.columns:
        return out
    for _, r in df.iterrows():
        site = _clean(r.get("Site name", "")).casefold()
        reason = _reason_from(r, REASON_COLUMNS)
        if not site or not reason:
            continue
        serial = _clean(r.get("Q'ty task/Local task", "")).casefold()
        out.setdefault((site, serial), reason)
        out.setdefault(site, reason)
    return out


def site_staff_map():
    """{កូដ Site → ឈ្មោះអ្នកកាន់ Site} ពី Tab allteam ជួរ Q:U (Q=site name, R=Code, S=Mether, T=Staff, U=Team)"""
    url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet=allteam"
    out = {}
    try:
        res = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=60)
        if res.status_code != 200:
            print(f"⚠️ allteam HTTP {res.status_code}")
            return out
        res.encoding = "utf-8"
        for row in csv.reader(io.StringIO(res.text)):
            if len(row) < 21:
                continue
            q, code, staff = row[16].strip(), row[17].strip(), row[19].strip()
            if staff and code and code.lower() != "code":
                out[code.upper()] = staff
                if q:
                    out.setdefault(q.upper(), staff)
    except requests.RequestException as e:
        print(f"⚠️ allteam: {e}")
    return out


def collect_not_approved(task_codes):
    holders = task_holders()
    site_staff = site_staff_map()
    """ត្រឡប់ {team: [ {task, site, staff, reason} ]} សម្រាប់ជួរ Result ≠ Approved"""
    out = {t: [] for t in VALID_TEAMS}
    for code in task_codes:
        df = fetch_csv(code)
        if df is None or df.empty or "Team" not in df.columns or "Site name" not in df.columns:
            continue
        df["Team"] = df["Team"].astype(str).str.strip()
        df = df[df["Team"].isin(VALID_TEAMS)]
        if "Result" in df.columns:
            _res = df["Result"].fillna("").astype(str).str.strip().str.lower()
            # យកតែ Site ដែលមាន Result ហើយមិនមែន Approved (Not Approved) — មិនយក Site មិនទាន់ធ្វើ (Result ទទេ)
            df = df[(_res != "approved") & (~_res.isin(_EMPTY))]
        mreasons = master_reasons(code)
        holder = holders.get(str(code).strip().upper(), "")
        for _, r in df.iterrows():
            site = _clean(r.get("Site name", ""))
            if not site:
                continue
            serial = _clean(r.get("Q'ty task/Local task", "")).casefold()
            reason = (mreasons.get((site.casefold(), serial)) or mreasons.get(site.casefold())
                      or _reason_from(r, REASON_COLUMNS))     # ១) Sheet មេ ២) Sheet Report ខ្លួនឯង
            out[r["Team"]].append({"task": code, "site": site,
                                   "staff": holder,
                                   "site_staff": site_staff.get(site.upper()) or (_clean(r.get("Name staff", "")) if "Name staff" in df.columns else ""),
                                   "reason": reason})
    return out


# ---------- ជំហានទី ១ និង ២ = collect_not_approved() (រក Site Not Approved + រកមូលហេតុក្នុង Sheet មេ) ----------

def _call(holder):
    return "អូន" if _clean(holder).casefold() in {x.casefold() for x in CALL_AS_ADIK} else "បង"


# ---------- ជំហានទី ៣: គ្មានមូលហេតុ → សួរអ្នកកាន់ការងារ (ម្តងក្នុងមួយ Site ក្នុងមួយថ្ងៃ) ----------
def build_question_message(team, rows, asked_today):
    """ត្រឡប់ (text ឬ None, keys ថ្មីដែលបានសួរ)"""
    by_holder, keys = {}, []
    for r in rows:
        if r["reason"]:
            continue
        key = f"{team}|{r['task']}|{r['site']}"
        if key in asked_today:
            continue
        by_holder.setdefault(_clean(r["staff"]), []).append(r)
        keys.append(key)
    if not by_holder:
        return None, []
    blocks = []
    for holder, rs in by_holder.items():
        lines = [f"សួស្តី{_call(holder)} {staff_tag(holder)}"]
        lines += [f"ការងារ {x['task']} site {x['site']} មានមូលហេតុអី ឬខុសអី បានមិន Approved?" for x in rs]
        blocks.append("\n".join(lines))
    return "\n\n".join(blocks), keys


# ---------- ជំហានទី ៤: បានមូលហេតុហើយ → ផ្ញើចូលក្រុម (Tag អ្នកកាន់ការងារ + Site + មូលហេតុ) ----------
def _tidy_reason(v):
    """លុបសញ្ញាក្បៀសស្ទួន/ទទេ ឧ. 'with ,,,,,Air Conditioner,,,Other' → 'with Air Conditioner, Other'"""
    v = re.sub(r"\b(with|of|for)\s*,[\s,]*", r"\1 ", str(v), flags=re.I)
    v = re.sub(r"\s*,[\s,]*", ", ", v)
    return v.strip(" ,;.")


def build_reason_message(team, rows, reported):
    """មូលហេតុដែលបានហើយ → ក្រុម Team: ប្រាប់មេធីម និងកូនចៅអ្នកកាន់ Site។ ត្រឡប់ (text ឬ None, {key: reason} ថ្មី/ប្តូរ)"""
    groups, fresh = {}, {}
    for r in rows:
        if not r["reason"]:
            continue
        key = f"{team}|{r['task']}|{r['site']}"
        if reported.get(key) == r["reason"]:
            continue
        groups.setdefault(r["reason"], []).append(r)
        fresh[key] = r["reason"]
    if not groups:
        return None, {}
    mention = TEAM_TELEGRAM_MENTIONS.get(team, TEAM_LEADERS.get(team, team))
    leader = TEAM_LEADERS.get(team, team)
    blocks = []
    for reason, rs in groups.items():
        staff = [n for n in dict.fromkeys(x["site_staff"] for x in rs if x.get("site_staff")) if n.casefold() != leader.casefold()]
        head = f"បង {mention} ({leader})" + (f" & {', '.join(staff)}" if staff else "")
        many = len(staff) > 1          # បុគ្គលិកច្រើននាក់ → ដាក់ឈ្មោះក្រោយ Site នីមួយៗ
        lines = [head, "ជួយឆែកការងារនេះដែលខុស មិនទាន់ Approved"]
        lines += [f"• {x['task']} {x['site']}" + (f" ({x['site_staff']})" if many and x.get("site_staff") else "") for x in rs]
        lines.append(f"មូលហេតុ: {_tidy_reason(reason)[:300]}")
        blocks.append("\n".join(lines))
    return "\n\n".join(blocks), fresh


def _send_long(client, chat, text):
    chunk = ""
    for line in text.split("\n"):
        if len(chunk) + len(line) + 1 > 3500:
            client.send_message(chat, chunk, parse_mode=None)
            chunk = ""
        chunk += line + "\n"
    if chunk.strip():
        client.send_message(chat, chunk, parse_mode=None)


def staff_handle(name):
    """@username របស់អ្នកកាន់ការងារ (ឬ None)"""
    try:
        extra = json.loads(os.environ.get("STAFF_TELEGRAM_JSON") or "{}")
    except json.JSONDecodeError:
        extra = {}
    handles = {k.strip().casefold(): v for k, v in {**STAFF_TELEGRAM, **extra}.items()}
    return handles.get(_clean(name).casefold())


def resolve_dm_target(client, handle):
    """@username ប្រើបានភ្លាម; លេខទូរស័ព្ទ (+855…) ត្រូវនាំចូលជា Contact ជាមុនទើបផ្ញើបាន"""
    if not str(handle).startswith("+"):
        return handle
    from telethon.tl.functions.contacts import ImportContactsRequest
    from telethon.tl.types import InputPhoneContact
    res = client(ImportContactsRequest([InputPhoneContact(client_id=0, phone=handle, first_name="Staff", last_name=handle[-4:])]))
    if not res.users:
        raise RuntimeError(f"លេខ {handle} មិនមានគណនី Telegram ឬបិទការស្វែងរក")
    return res.users[0]


def run_hourly(client, task_chat_ids, team_chat_mapping, warning_state, today, now):
    day = today.isoformat()
    asked = warning_state.get("_asked", {})
    if asked.get("date") != day:
        asked = {"date": day, "keys": [], "round_done": False}
    asked_today = set(asked["keys"])
    rep = warning_state.get("_reported", {})
    if rep.get("date") != day:
        rep = {"date": day, "items": {}}          # រាល់ថ្ងៃថ្មី ផ្ញើមូលហេតុដែលនៅសល់ម្តងទៀតម្តង
    reported = rep["items"]

    data = collect_not_approved(sorted(task_chat_ids))                      # ១+២
    # ការសួរអ្នកកាន់ការងារ៖ ថ្ងៃច័ន្ទ–សៅរ៍ (មិនសួរថ្ងៃអាទិត្យ) ហើយម្តងក្នុងមួយថ្ងៃ (ជុំដំបូងនៃថ្ងៃ)
    # ម៉ោងសួរ៖ ចន្លោះ 9:00–10:00 (ម៉ោងកម្ពុជា) ប៉ុណ្ណោះ
    in_ask_hour = now is None or ASK_FROM <= now.hour < ASK_TO
    ask_today_allowed = (today.weekday() != 6 and in_ask_hour and not asked.get("round_done")) or os.environ.get("FORCE_ASK") == "1"
    # ៣. គ្មានមូលហេតុ → សួរអ្នកកាន់ការងារ។ មាន @username → ឆាតផ្ទាល់ (DM) ម្នាក់មួយសារ; គ្មាន → សួរក្នុងក្រុម Team
    dm = {}   # handle → {"name":..., "items":[(team,row,key)]}
    for team in VALID_TEAMS:
        rows, chat = data.get(team) or [], team_chat_mapping.get(team)
        if not rows:
            continue
        no_dm_rows = []
        for r in rows:
            if r["reason"] or not ask_today_allowed:
                continue
            key = f"{team}|{r['task']}|{r['site']}"
            if key in asked_today or _clean(r["staff"]).casefold() in {x.casefold() for x in NEVER_ASK}:
                continue
            h = staff_handle(r["staff"])
            if h:
                dm.setdefault(h, {"name": _clean(r["staff"]), "items": []})["items"].append((team, r, key))
            else:
                no_dm_rows.append(r)
        if no_dm_rows and chat:
            q_text, q_keys = build_question_message(team, no_dm_rows, asked_today)
            if q_text:
                try:
                    _send_long(client, chat, q_text)
                    asked_today.update(q_keys)
                except Exception as e:
                    print(f"⚠️ Question to {team} failed: {e}")
    for handle, info in dm.items():
        lines = [f"សួស្តី{_call(info['name'])} {info['name']}"]
        lines += [f"ការងារ {r['task']} site {r['site']} ({team}) មានមូលហេតុអី ឬខុសអី បានមិន Approved?" for team, r, _ in info["items"]]
        try:
            _send_long(client, resolve_dm_target(client, handle), "\n".join(lines))
            asked_today.update(k for _, _, k in info["items"])
            print(f"✉️ DM → {handle}: {len(info['items'])} Site")
        except Exception as e:
            print(f"⚠️ DM to {handle} failed: {e} — ទុកសួរក្នុងក្រុមវិញ")
            for team in {t for t, _, _ in info["items"]}:
                chat = team_chat_mapping.get(team)
                rows = [r for t, r, _ in info["items"] if t == team]
                q_text, q_keys = build_question_message(team, rows, asked_today)
                if q_text and chat:
                    try:
                        _send_long(client, chat, q_text); asked_today.update(q_keys)
                    except Exception as e2:
                        print(f"⚠️ Fallback question to {team} failed: {e2}")
    if ask_today_allowed:
        asked["round_done"] = True          # ជុំសួរនៃថ្ងៃនេះបានធ្វើរួច → ម៉ោងបន្ទាប់មិនសួរទៀត
    # ៤. បានមូលហេតុហើយ → ផ្ញើចូលក្រុម Team
    for team in VALID_TEAMS:
        rows, chat = data.get(team) or [], team_chat_mapping.get(team)
        if not rows or not chat:
            continue
        r_text, fresh = build_reason_message(team, rows, reported)
        if r_text:
            try:
                _send_long(client, chat, r_text)
                reported.update(fresh)
            except Exception as e:
                print(f"⚠️ Reason report to {team} failed: {e}")
    asked["keys"] = sorted(asked_today)
    warning_state["_asked"] = asked
    warning_state["_reported"] = rep


# ===== របាយការណ៍សរុបប្រចាំសប្តាហ៍ (ថ្ងៃសៅរ៍ព្រឹក) ជា PDF ៥ ទំព័រ =====
KH_DIGITS = str.maketrans("0123456789", "០១២៣៤៥៦៧៨៩")
KH_MONTHS = ["មករា", "កុម្ភៈ", "មីនា", "មេសា", "ឧសភា", "មិថុនា", "កក្កដា", "សីហា", "កញ្ញា", "តុលា", "វិច្ឆិកា", "ធ្នូ"]


def kh(n):
    return str(n).translate(KH_DIGITS)


def new_tally():
    return {"teams": {t: {"target": 0, "approved": 0, "notappr": 0, "pending": 0, "na_tasks": {}, "pe_tasks": {}} for t in VALID_TEAMS},
            "zero": []}


def tally_task(df, code, tally):
    """រាប់ Target / Approved / មិន Approved (មាន Result តែមិនមែន Approved) / មិនទាន់ធ្វើ (Result ទទេ) តាមធីម"""
    code = str(code).strip()
    if df is None or df.empty or "Team" not in df.columns or "Site name" not in df.columns:
        return
    d = df.copy()
    d["Team"] = d["Team"].astype(str).str.strip()
    d = d[d["Team"].isin(VALID_TEAMS) & d["Site name"].notna()
          & ~d["Site name"].astype(str).str.strip().str.lower().isin(_EMPTY)]
    if d.empty:
        tally["zero"].append(code)
        return
    res = d["Result"].fillna("").astype(str).str.strip() if "Result" in d.columns else pd.Series([""] * len(d), index=d.index)
    low = res.str.casefold()
    for team in VALID_TEAMS:
        m = d["Team"] == team
        n = int(m.sum())
        if not n:
            continue
        ap = int((low[m] == "approved").sum())
        blank = int((low[m].isin(_EMPTY)).sum())
        na = n - ap - blank
        t = tally["teams"][team]
        t["target"] += n; t["approved"] += ap; t["notappr"] += na; t["pending"] += blank
        if na:
            t["na_tasks"][code] = t["na_tasks"].get(code, 0) + na
        if blank:
            t["pe_tasks"][code] = t["pe_tasks"].get(code, 0) + blank


def weekly_records(today):
    """ចំនួនកិច្ចការ Approved ដែលបានកត់ត្រាក្នុងមួយសប្តាហ៍ (1–7, 8–14, 15–21 …) ពី Tab Daily_Count"""
    df = fetch_csv("Daily_Count")
    out = {t: [0] * ((today.day + 6) // 7) for t in VALID_TEAMS}
    if df is None or df.empty:
        return out
    dcol = df.columns[0]
    for _, r in df.iterrows():
        m = re.search(r"(\d{4})-(\d{2})-(\d{2})", str(r[dcol]))
        if not m:
            continue
        y, mo, da = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if (y, mo) != (today.year, today.month) or da > today.day:
            continue
        w = (da - 1) // 7
        for t in VALID_TEAMS:
            try:
                out[t][w] += int(float(r.get(t, 0) or 0))
            except (ValueError, TypeError):
                pass
    return out


def build_saturday_html(tally, weekly, today):
    T = tally["teams"]
    tot = {k: sum(T[t][k] for t in VALID_TEAMS) for k in ("target", "approved", "notappr", "pending")}
    pc = lambda a, b: int(a * 100 / b + 0.5) if b else 0
    weeks = (today.day + 6) // 7
    title_sub = f"សរុប {kh(weeks)}សប្ដាហ៍ ({kh(1)}–{kh(today.day)} {KH_MONTHS[today.month - 1]} {today.year}) តាមធីមនីមួយៗ"
    zero = ", ".join(sorted(set(tally["zero"])))
    css = """
@page{size:960px 540px;margin:0}
*{box-sizing:border-box}
body{margin:0;font-family:'Noto Sans Khmer','Noto Sans','Khmer OS',sans-serif;-webkit-print-color-adjust:exact}
.s{width:960px;height:540px;padding:56px 64px;page-break-after:always;overflow:hidden;position:relative}
.dark{background:#14213d;color:#f6f3ec}.light{background:#f6f1e8;color:#14213d}
h1{font-size:46px;margin:30px 0 10px}h2{font-size:28px;margin:0 0 18px}
.sub{font-size:17px;opacity:.9}.cards{display:flex;gap:18px;margin-top:34px}
.c{flex:1;background:#1f2f57;border-radius:10px;padding:18px 20px}.c b{display:block;font-size:46px}.c span{font-size:15px}
.note{font-size:12px;opacity:.8;margin-top:22px}
table{border-collapse:collapse;width:100%;font-size:15px}th{background:#14213d;color:#fff;padding:7px;border:1px solid #c9c4ba}
td{padding:6px 8px;border:1px solid #c9c4ba;text-align:center;background:#f6f1e8}tr.tot td{background:#14213d;color:#fff;font-weight:bold}
.g{color:#1f8a5b}.r{color:#c0394b}.o{color:#c98a14}
.bar{display:flex;align-items:center;margin:11px 0;font-size:14px}.bar .n{width:90px}.bar .b{display:flex;height:27px;width:560px}.bar .l{margin-left:14px}
.leg{display:flex;gap:40px;font-size:13px;margin:0 0 6px 90px}.leg i{display:inline-block;width:14px;height:14px;margin-right:6px;vertical-align:-2px}
.chart{display:flex;align-items:flex-end;gap:22px;height:300px;margin-top:10px;padding-left:10px;border-bottom:1px solid #14213d}
.grp{display:flex;align-items:flex-end;gap:5px;height:100%}.col{width:26px;text-align:center;font-size:12px}.col div{width:100%}
.lab{display:flex;gap:22px;padding-left:10px;font-size:13px}.lab div{width:88px;text-align:center}
"""
    slides = []
    slides.append(f"""<div class="s dark"><h1>របាយការណ៍ការងារ Power CHA</h1><div class="sub">{title_sub} · Sheet A1 ដល់ K5</div>
<div class="cards"><div class="c"><b>{kh(tot['target'])}</b><span>គោលដៅសរុប</span></div>
<div class="c"><b style="color:#5fd3a2">{kh(tot['approved'])}</b><span>Approved ({kh(pc(tot['approved'], tot['target']))}%)</span></div>
<div class="c"><b style="color:#ff7a90">{kh(tot['notappr'])}</b><span>មិន Approved</span></div>
<div class="c"><b style="color:#f6b94d">{kh(tot['pending'])}</b><span>មិនទាន់ធ្វើ</span></div></div>
<div class="note">ស្ថានភាព {kh(today.strftime('%d/%m/%Y'))}{(' · Sheet គ្មានគោលដៅ (០) មិនបានបង្ហាញ៖ ' + zero) if zero else ''}</div></div>""")
    bars = ""
    for t in VALID_TEAMS:
        x = T[t]; n = x["target"] or 1
        bars += (f'<div class="bar"><div class="n">{t}</div><div class="b">'
                 f'<div style="width:{x["approved"]*100/n:.1f}%;background:#1f8a5b"></div>'
                 f'<div style="width:{x["notappr"]*100/n:.1f}%;background:#c0394b"></div>'
                 f'<div style="width:{x["pending"]*100/n:.1f}%;background:#e5a11b"></div></div>'
                 f'<div class="l">{x["approved"]}/{x["target"]} · {pc(x["approved"], x["target"])}%</div></div>')
    slides.append(f"""<div class="s light"><h2>ស្ថានភាពតាមធីម៖ Approved · មិន Approved · មិនទាន់ធ្វើ</h2>
<div class="leg"><span><i style="background:#1f8a5b"></i>Approved</span><span><i style="background:#c0394b"></i>មិន Approved</span><span><i style="background:#e5a11b"></i>មិនទាន់ធ្វើ</span></div>{bars}</div>""")
    rows = "".join(f'<tr><td>{t}</td><td>{T[t]["target"]}</td><td class="g">{T[t]["approved"]}</td><td class="r">{T[t]["notappr"]}</td>'
                   f'<td class="o">{T[t]["pending"]}</td><td>{pc(T[t]["approved"], T[t]["target"])}%</td></tr>' for t in VALID_TEAMS)
    slides.append(f"""<div class="s light"><h2>តារាងសរុបតាមធីម</h2><table><tr><th>ធីម</th><th>គោលដៅ</th><th>Approved</th><th>មិន Approved</th><th>មិនទាន់ធ្វើ</th><th>%</th></tr>{rows}
<tr class="tot"><td>សរុប</td><td>{tot['target']}</td><td>{tot['approved']}</td><td>{tot['notappr']}</td><td>{tot['pending']}</td><td>{pc(tot['approved'], tot['target'])}%</td></tr></table></div>""")
    shades = ["#9fb6d6", "#4a77a8", "#14213d", "#6b8e9c", "#2c4a6b"]
    mx = max([v for t in VALID_TEAMS for v in weekly[t]] + [1])
    groups = labs = ""
    for t in VALID_TEAMS:
        cols = "".join(f'<div class="col"><span>{v}</span><div style="height:{v*250/mx:.0f}px;background:{shades[i % len(shades)]}"></div></div>' for i, v in enumerate(weekly[t]))
        groups += f'<div class="grp">{cols}</div>'
        labs += f'<div>{t[-3:]} · {sum(weekly[t])}</div>'
    legend = "".join(f'<span><i style="background:{shades[i % len(shades)]}"></i>សប្ដាហ៍ទី{kh(i+1)} ({kh(i*7+1)}–{kh(min((i+1)*7, today.day))})</span>' for i in range(weeks))
    slides.append(f"""<div class="s light"><h2>ចំនួនកិច្ចការកត់ត្រាក្នុងមួយសប្ដាហ៍</h2><div class="leg" style="margin-left:0">{legend}</div>
<div class="chart">{groups}</div><div class="lab">{labs}</div></div>""")
    fmt = lambda d: ", ".join(f"{k}:{v}" for k, v in sorted(d.items())) or "–"
    rows = "".join(f'<tr><td>{t}</td><td class="r">{fmt(T[t]["na_tasks"])}</td><td class="o">{fmt(T[t]["pe_tasks"])}</td><td>{T[t]["pending"]}</td></tr>' for t in VALID_TEAMS)
    slides.append(f"""<div class="s light"><h2>ការងារមិន Approved និងមិនទាន់ធ្វើ (ដាក់ដោយឡែក)</h2><table><tr><th>ធីម</th><th>មិន Approved</th><th>មិនទាន់ធ្វើ (Sheet : ចំនួន)</th><th>សរុប</th></tr>{rows}</table>
<div class="note" style="color:#14213d">មិនទាន់ធ្វើ = គោលដៅ − Approved − មិន Approved</div></div>""")
    return f"<!doctype html><html><head><meta charset='utf-8'><style>{css}</style></head><body>{''.join(slides)}</body></html>"


def render_pdf(html, pdf_path):
    import shutil, subprocess
    html_path = os.path.abspath(pdf_path + ".html")
    with open(html_path, "w", encoding="utf-8") as fh:
        fh.write(html)
    exe = shutil.which("google-chrome") or shutil.which("chromium") or shutil.which("chromium-browser") or os.environ.get("CHROME_BIN")
    if not exe:
        raise RuntimeError("រក Chrome មិនឃើញ")
    subprocess.run([exe, "--headless=new", "--no-sandbox", "--disable-gpu", "--no-pdf-header-footer",
                    f"--print-to-pdf={os.path.abspath(pdf_path)}", "file://" + html_path], check=True, timeout=120,
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    os.remove(html_path)
    return pdf_path


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

    if os.environ.get("EVENT_ACTION") == "daily-report-hourly":
        if not (HOURLY_FROM <= now.hour <= HOURLY_TO) and os.environ.get("FORCE_RESEND") != "1":
            print(f"ℹ️ Hourly: outside {HOURLY_FROM}-{HOURLY_TO}h (now {now:%H:%M}) — skip")
            return
        with TelegramClient(StringSession(session_str), api_id, api_hash) as client:
            try:
                client.get_dialogs()
            except Exception as e:
                print(f"⚠️ get_dialogs failed: {e}")
            run_hourly(client, task_chat_ids, team_chat_mapping, warning_state, today, now)
        save_warning_state(warning_state)
        return

    # ===== ផ្ញើតាមម៉ោងថេរ (GitHub Actions free ចាប់ផ្តើមយឺតមិនទៀង) =====
    # ស្វ័យប្រវត្តិ = schedule របស់ GitHub ឬ ហៅពី cron ខាងក្រៅ (RUN_SOURCE=cron)។ ដំណើរការដោយដៃ = មិនមែនទាំងពីរ។
    is_scheduled = (os.environ.get("GITHUB_EVENT_NAME") == "schedule"
                    or (os.environ.get("GITHUB_EVENT_NAME") == "repository_dispatch"
                        and os.environ.get("EVENT_ACTION") != "daily-report-evening")   # ល្ងាច = ទិន្នន័យទាំងអស់
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
    tally = new_tally()
    team_left = {team: {} for team in VALID_TEAMS}   # {team: {task: Site មិនទាន់ Approved}} សម្រាប់លិខិតប្រចាំខែ

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
            tally_task(df_task, task_code, tally)

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
                        styled_detail = style_pending_table(df_single_team, f"{task_title} ({team})" + ("" if is_scheduled else " — All data"))
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
                        _left = sum(len(w) for w in groups.values()) - sum(approved_count(w) for w in groups.values())
                        if _left > 0:
                            team_left[team][str(task_code).strip()] = _left
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
            tally_task(df_x, code, tally)
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

    # ៥. របាយការណ៍សរុបប្រចាំសប្តាហ៍ (PDF) ថ្ងៃសៅរ៍ព្រឹក (Run ស្វ័យប្រវត្តិ) ផ្ញើចូល Main Group
        if ((today.weekday() == 5 and is_scheduled) or os.environ.get("FORCE_SAT_REPORT") == "1") \
                and warning_state.get("_sat_report") != today.isoformat():
            try:
                pdf = render_pdf(build_saturday_html(tally, weekly_records(today), today), f"report_summary_{today.isoformat()}.pdf")
                if send_with_caption(client, MAIN_GROUP_ID, pdf, f"📊 របាយការណ៍ការងារ Power CHA — សរុបតាមធីម ({today:%d/%m/%Y})"):
                    warning_state["_sat_report"] = today.isoformat()
            except Exception as e:
                print(f"⚠️ Saturday report failed: {e}")

    # ៤. លិខិតព្រមានប្រចាំខែ: ថ្ងៃចុងខែ ក្រុមណាមានការងារមិនទាន់ចប់ក្នុងខែ រុញទៅខែក្រោយ ម្តងក្នុងមួយខែ
        # ផ្ញើតែនៅ Run ល្ងាច (ទិន្នន័យចុងថ្ងៃ) ថ្ងៃចុងខែ
        month_end = (today + timedelta(days=1)).day == 1 and (os.environ.get("EVENT_ACTION") == "daily-report-evening" or os.environ.get("MONTHLY_ANYTIME") == "1")
        if month_end and warning_state.get("_monthly_sent") != today.strftime("%Y-%m") and any(team_left.values()):
            mpath = make_monthly_letter(team_left, today)
            if mpath and send_with_caption(client, MAIN_GROUP_ID, mpath,
                                           f"⚠️ សេចក្តីព្រាងលិខិតព្រមានប្រចាំខែ {today:%B %Y} (ការងារមិនចប់ក្នុងខែ រុញទៅខែក្រោយ)"):
                warning_state["_monthly_sent"] = today.strftime("%Y-%m")

    if is_scheduled:
        warning_state["_last_run"] = today.isoformat()
    save_warning_state(warning_state)

if __name__ == "__main__":
    main()
