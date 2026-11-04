from datetime import datetime, timezone, timedelta
import io
import json
import os
import re

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
    "CHA-T01": "An Vanny", "CHA-T02": "Thoeung Chanthan",
    "CHA-T03": "Au Chhaira", "CHA-T04": "Chhorn Kim Chheang",
    "CHA-T05": "Mut Chanthy", "CHA-T06": "Rit Borin", "CHA-T07": "Sam Sophal",
}
START_DATE = datetime.strptime(os.environ.get("START_DATE") or "2026-11-02", "%Y-%m-%d").date()
DEFAULT_WEEKS = 4
TASK_WEEKS = {"A2": 3}
LOW_PROGRESS_THRESHOLD = 50
WARNING_STATE_FILE = os.environ.get("WARNING_STATE_FILE", "weekly_warning_state.json")

CAPTION_STYLE = {
    "selector": "caption",
    "props": [("caption-side", "top"), ("font-size", "22px"), ("font-weight", "bold"),
              ("text-align", "center"), ("background-color", "#27AE60"), ("color", "white"),
              ("padding", "10px"), ("border", "1px solid black")],
}


def task_weeks(task_code):
    return TASK_WEEKS.get(str(task_code).strip().upper(), DEFAULT_WEEKS)


def get_current_week(today):
    return (today - START_DATE).days // 7 + 1


def normalize_team(name):
    match = re.search(r"(\d+)\s*$", str(name).strip())
    return f"CHA-T{int(match.group(1)):02d}" if match else None


def fetch_csv(sheet_name_or_gid):
    url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet={sheet_name_or_gid}"
    try:
        response = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=15)
    except requests.RequestException as exc:
        print(f"⚠️ Request failed for sheet '{sheet_name_or_gid}': {exc}")
        return None
    if response.status_code != 200:
        print(f"⚠️ Failed to fetch sheet '{sheet_name_or_gid}': HTTP {response.status_code}")
        return None
    lines = response.text.splitlines()
    header_row_idx = 0
    keywords = {"site name", "team", "sheet", "chatid", "no", "no."}
    for idx, line in enumerate(lines[:15]):
        cells = [cell.strip().strip('"').lower() for cell in line.split(",")]
        if any(cell in keywords for cell in cells):
            header_row_idx = idx
            break
    try:
        df = pd.read_csv(io.StringIO(response.text), skiprows=header_row_idx).dropna(how="all")
    except (pd.errors.EmptyDataError, pd.errors.ParserError) as exc:
        print(f"⚠️ Could not parse sheet '{sheet_name_or_gid}': {exc}")
        return None
    df.columns = df.columns.astype(str).str.strip()
    rename = {}
    for col in df.columns:
        lower = col.lower().strip()
        if "site name" in lower:
            rename[col] = "Site name"
        elif lower in ("no", "no."):
            rename[col] = "No."
        elif lower == "team":
            rename[col] = "Team"
        elif "group task" in lower:
            rename[col] = "Group task"
        elif "result" in lower:
            rename[col] = "Result"
        elif "serial" in lower:
            rename[col] = "Serial"
        elif lower in ("name staff", "staff name"):
            rename[col] = "Name staff"
    df = df.rename(columns=rename)
    return df.loc[:, ~df.columns.duplicated()]


def get_task_title_from_sheet(sheet_name_or_gid):
    url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet={sheet_name_or_gid}"
    try:
        response = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=15)
        if response.status_code == 200 and response.text.splitlines():
            title = response.text.splitlines()[0].replace('"', "").split(",")[0].strip()
            if title and not title.isdigit():
                return title
    except requests.RequestException as exc:
        print(f"⚠️ Title lookup failed for '{sheet_name_or_gid}': {exc}")
    return f"Task {sheet_name_or_gid}"


def style_table(df, title):
    return df.style.set_caption(title).set_table_styles([
        CAPTION_STYLE,
        {"selector": "th", "props": [("background-color", "#369388"), ("color", "black"),
            ("font-weight", "bold"), ("text-align", "center"), ("border", "1px solid black"), ("padding", "6px")]},
        {"selector": "td", "props": [("text-align", "center"), ("border", "1px solid black"), ("padding", "5px")]},
    ])


def load_team_chat_ids(mapping):
    chats = {}
    if mapping is None or not {"Column A", "Column B"}.issubset(mapping.columns):
        print("⚠️ 'Column A' / 'Column B' not found; team-group reports disabled.")
        return chats
    for _, row in mapping.dropna(subset=["Column A", "Column B"]).iterrows():
        team = normalize_team(row["Column A"])
        if team not in VALID_TEAMS:
            continue
        try:
            chats[team] = int(float(str(row["Column B"]).strip()))
        except (ValueError, TypeError):
            continue
    return chats


def split_sites_into_weeks(team_df, task_code):
    """Sort by Site name, then divide sites as evenly as possible among weeks."""
    count_weeks = task_weeks(task_code)
    sites = team_df.copy()
    sites["_site_sort"] = sites["Site name"].fillna("").astype(str).str.strip().str.casefold()
    sites = sites[sites["_site_sort"] != ""].sort_values("_site_sort", kind="stable")
    sites = sites.drop(columns=["_site_sort"]).reset_index(drop=True)
    quotient, remainder = divmod(len(sites), count_weeks)
    result, start = {}, 0
    for week in range(1, count_weeks + 1):
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
    except (FileNotFoundError, json.JSONDecodeError, OSError) as exc:
        print(f"ℹ️ Could not load warning history: {exc}; starting a new tracking file.")
        return {}


def save_warning_state(state):
    with open(WARNING_STATE_FILE, "w", encoding="utf-8") as handle:
        json.dump(state, handle, ensure_ascii=False, indent=2)


def weekly_team_table(items):
    rows = []
    for index, item in enumerate(items, start=1):
        row = {"No.": index, "Task": item["Task"]}
        for week in range(1, DEFAULT_WEEKS + 1):
            data = item["Stats"].get(week)
            row[f"Week {week} Target"] = data["Target"] if data else "—"
            row[f"Week {week} Result"] = data["Result"] if data else "—"
            row[f"Week {week} %"] = f"{data['Percent']}%" if data else "—"
        rows.append(row)
    total = {"No.": "", "Task": "TOTAL"}
    for week in range(1, DEFAULT_WEEKS + 1):
        vals = [item["Stats"][week] for item in items if week in item["Stats"]]
        target = sum(value["Target"] for value in vals)
        result = sum(value["Result"] for value in vals)
        total[f"Week {week} Target"] = target
        total[f"Week {week} Result"] = result
        total[f"Week {week} %"] = f"{round(result * 100 / target) if target else 0}%"
    rows.append(total)
    return pd.DataFrame(rows).fillna("")


def current_week_team_progress(items, week_number):
    target = result = 0
    for item in items:
        data = item["Stats"].get(week_number)
        if data:
            target += data["Target"]
            result += data["Result"]
    return target, result, round(result * 100 / target) if target else 0


def set_cell_text(cell, text, bold=False, size=9):
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run(str(text))
    run.bold = bold
    run.font.name = "Times New Roman"
    run.font.size = Pt(size)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def make_warning_letter_docx(teams, progress, today, week_number):
    """Create an unsigned director-review DOCX using the supplied warning-letter layout."""
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.55)
    section.bottom_margin = Inches(0.55)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)
    doc.styles["Normal"].font.name = "Times New Roman"
    doc.styles["Normal"].font.size = Pt(10)

    header = doc.add_table(rows=1, cols=2)
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
    title.add_run("DRAFT — WARNING LETTER").bold = True
    title.runs[0].font.size = Pt(16)
    doc.add_paragraph(
        "Subject: Warning letter for Team Leaders whose teams have remained below 50% "
        "of the weekly target for the third time."
    ).runs[0].bold = True
    doc.add_paragraph("To: Board of Directors, Technical Center")
    doc.add_paragraph(
        "Based on the weekly progress reports, the teams listed below achieved less than "
        "50% of their assigned target for the third recorded week. This is a draft for review "
        "and is not effective until approved and signed by the Director."
    )
    doc.add_paragraph("+ Warning to:").runs[0].bold = True
    doc.add_paragraph("Kompong Cham Branch would like to submit this draft warning to the Team Leaders listed below.")
    doc.add_paragraph("+ Reasons for warning").runs[0].bold = True
    doc.add_paragraph(
        "1. The team has remained below 50% of its weekly target for the third recorded time. "
        "The team leader is requested to explain the root cause and corrective actions."
    )

    table = doc.add_table(rows=1, cols=6)
    table.style = "Table Grid"
    for cell, label in zip(table.rows[0].cells, ["N.o", "Team", "Name", "Position", "Reason for warning", "Penalty"]):
        set_cell_text(cell, label, True, 9)
    for index, team in enumerate(teams, start=1):
        target, result, percent = progress[team]
        cells = table.add_row().cells
        values = [index, team, TEAM_LEADERS[team], "Team Leader",
                  f"Below 50% for the third recorded week (week {week_number}); weekly result "
                  f"{result}/{target} ({percent}%). Please explain the cause and corrective action.",
                  "Proposed: Deduct 1KI, subject to Director approval"]
        for cell, value in zip(cells, values):
            set_cell_text(cell, value, size=8)

    doc.add_paragraph("+ Reminders for next time").runs[0].bold = True
    doc.add_paragraph("- Remind the listed Team Leaders to submit root causes and corrective actions.")
    doc.add_paragraph("- Request all team members to recheck tasks and finish before the required deadline.")
    doc.add_paragraph("- Kindly request the Board of Directors to review and approve this draft.")
    signature = doc.add_paragraph()
    signature.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    signature.add_run(
        f"Date: {today.strftime('%d/%m/%Y')}\nPrepared by\nPOWER STAFF\n\n\n"
        "____________________________\nDirector (Name and Signature)"
    )
    doc.add_paragraph("Recipients:").runs[0].bold = True
    doc.add_paragraph("- BOD of GIS (for review);\n- HR Dept.;\n- Archive: Admin, Kompong Cham Branch.")
    filename = f"warning_letter_draft_week_{week_number}_{today.isoformat()}.docx"
    doc.save(filename)
    return filename


def send_weekly_team_reports(client, team_rows, team_chats, today, week_number, warning_state):
    candidates, candidate_progress = [], {}
    for team in VALID_TEAMS:
        items = team_rows.get(team, [])
        chat_id = team_chats.get(team)
        if not items or not chat_id:
            if items and not chat_id:
                print(f"⚠️ No group chat ID for {team}; weekly report not sent.")
            continue
        target, result, percent = current_week_team_progress(items, week_number)
        title = f"Weekly report {team} — Target, Result and % by Site name"
        image_path = f"team_summary_{team}.png"
        dfi.export(style_table(weekly_team_table(items), title).hide(axis="index"), image_path, max_rows=-1)
        try:
            client.send_file(chat_id, image_path, caption=f"{title} — week {week_number}")
            text = (f"Weekly reminder {week_number}: {team} ({TEAM_LEADERS[team]})\n"
                    f"This week's result: {result}/{target} ({percent}%).\n")
            text += ("Below 50%. Please state the reason and corrective action in this group."
                     if target and percent < LOW_PROGRESS_THRESHOLD else
                     "Please continue tracking the target and complete the work on schedule.")
            client.send_message(chat_id, text)
        except Exception as exc:
            print(f"⚠️ Failed to send weekly report/reminder to {team}: {exc}")

        state = warning_state.setdefault(team, {"low_weeks": 0, "last_counted_week": 0, "draft_sent": False})
        if target and percent < LOW_PROGRESS_THRESHOLD:
            if int(state.get("last_counted_week", 0)) < week_number:
                state["low_weeks"] = int(state.get("low_weeks", 0)) + 1
                state["last_counted_week"] = week_number
            if state["low_weeks"] >= 3 and not state.get("draft_sent", False):
                candidates.append(team)
                candidate_progress[team] = (target, result, percent)
        else:
            state["last_counted_week"] = max(int(state.get("last_counted_week", 0)), week_number)

    if candidates:
        try:
            path = make_warning_letter_docx(candidates, candidate_progress, today, week_number)
            client.send_file(MAIN_GROUP_ID, path, caption=(
                "DRAFT WARNING LETTER — for Director review and signature only. "
                "No penalty is effective before approval."
            ))
            for team in candidates:
                warning_state[team]["draft_sent"] = True
        except Exception as exc:
            print(f"⚠️ Failed to create/send warning-letter draft: {exc}")


def main():
    api_id_raw = os.environ.get("TELEGRAM_API_ID")
    api_hash = os.environ.get("TELEGRAM_API_HASH")
    session_string = os.environ.get("TELEGRAM_SESSION")
    if not api_id_raw:
        raise ValueError("TELEGRAM_API_ID is missing or empty in environment/secrets")
    try:
        api_id = int(api_id_raw)
    except ValueError as exc:
        raise ValueError(f"TELEGRAM_API_ID is not a valid integer: '{api_id_raw}'") from exc
    if not api_hash:
        raise ValueError("TELEGRAM_API_HASH is missing or empty in environment/secrets")
    if not session_string:
        raise ValueError("TELEGRAM_SESSION is missing or empty in environment/secrets")

    mapping = fetch_csv("Team%20chat%20IDs")
    task_chats = {}
    if mapping is not None and {"Sheet", "ChatID"}.issubset(mapping.columns):
        for _, row in mapping.dropna(subset=["Sheet", "ChatID"]).iterrows():
            try:
                task_chats[str(row["Sheet"]).strip()] = int(float(str(row["ChatID"]).strip()))
            except (TypeError, ValueError):
                continue
    if not task_chats:
        print("⚠️ No task-to-chat mappings found in 'Team chat IDs'; nothing to send.")
        return

    team_chats = load_team_chat_ids(mapping)
    now = datetime.now(timezone(timedelta(hours=7)))
    today = now.date()
    current_week = get_current_week(today)
    warning_state = load_warning_state()
    team_rows = {team: [] for team in VALID_TEAMS}
    overall = {team: {"Target": 0, "Result": 0} for team in VALID_TEAMS}

    with TelegramClient(StringSession(session_string), api_id, api_hash) as client:
        for task_code, task_chat_id in task_chats.items():
            task_df = fetch_csv(task_code)
            if task_df is None or task_df.empty:
                print(f"⚠️ Skipping '{task_code}': sheet is empty or failed to load.")
                continue
            if not {"Team", "Site name", "Result"}.issubset(task_df.columns):
                print(f"⚠️ Skipping '{task_code}': required Team/Site name/Result column is missing.")
                continue
            task_df = task_df.copy()
            task_df["Team"] = task_df["Team"].fillna("").astype(str).str.strip()
            task_df["Site name"] = task_df["Site name"].fillna("").astype(str).str.strip()
            task_df = task_df[task_df["Team"].isin(VALID_TEAMS) & task_df["Site name"].ne("")]
            title = f"{task_code}. {get_task_title_from_sheet(task_code)}"

            for team in VALID_TEAMS:
                team_df = task_df[task_df["Team"] == team].copy()
                if team_df.empty:
                    continue
                groups = split_sites_into_weeks(team_df, task_code)
                stats = {}
                for week, week_df in groups.items():
                    target = len(week_df)
                    result = approved_count(week_df)
                    stats[week] = {"Target": target, "Result": result,
                                   "Percent": round(result * 100 / target) if target else 0}
                team_rows[team].append({"Task": task_code, "Weeks": task_weeks(task_code), "Stats": stats})
                overall[team]["Target"] += len(team_df)
                overall[team]["Result"] += approved_count(team_df)

            columns = [col for col in ["No.", "Group task", "Branch", "Site name", "Serial",
                "Q'ty task/Local task", "Result", "Remark", "Last date record", "History Task",
                "Team", "Name staff"] if col in task_df.columns]
            details = task_df[columns].copy()
            if now.hour < 12:
                details = details[details["Result"].astype(str).str.strip().str.casefold() != "approved"]
            if not details.empty:
                image_path = f"detail_{task_code}.png"
                dfi.export(style_table(details, title).hide(axis="index"), image_path, max_rows=-1)
                try:
                    client.send_file(task_chat_id, image_path, caption=title)
                except Exception as exc:
                    print(f"⚠️ Failed to send task detail for {task_code}: {exc}")

        if today.weekday() == 5 and 1 <= current_week <= DEFAULT_WEEKS:
            send_weekly_team_reports(client, team_rows, team_chats, today, current_week, warning_state)

        if now.hour >= 12:
            rows, total_target, total_result = [], 0, 0
            for index, team in enumerate(VALID_TEAMS, start=1):
                target, result = overall[team]["Target"], overall[team]["Result"]
                percent = round(result * 100 / target) if target else 0
                rows.append({"No.": index, "Team": team, "Target": target, "Result": result, "%": f"{percent}%"})
                total_target += target
                total_result += result
            rows.append({"No.": "TOTAL", "Team": "", "Target": total_target, "Result": total_result,
                         "%": f"{round(total_result * 100 / total_target) if total_target else 0}%"})
            summary = pd.DataFrame(rows)
            image_path = "overall_report.png"
            dfi.export(style_table(summary, f"Overall report M{now.month}").hide(axis="index"), image_path, max_rows=-1)
            client.send_file(MAIN_GROUP_ID, image_path, caption=f"Overall report M{now.month}")

    save_warning_state(warning_state)


if __name__ == "__main__":
    main()
