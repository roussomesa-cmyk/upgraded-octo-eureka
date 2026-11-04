from datetime import datetime, timezone, timedelta
import io
import json
import os
import re

import dataframe_image as dfi
import pandas as pd
import requests
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
START_DATE = datetime.strptime(os.environ.get("START_DATE") or "2026-11-02", "%Y-%m-%d").date()
DEFAULT_WEEKS = 4
TASK_WEEKS = {"A2": 3}
LOW_PROGRESS_THRESHOLD = 50
WARNING_STATE_FILE = os.environ.get("WARNING_STATE_FILE", "weekly_warning_state.json")

CAPTION_STYLE = {
    "selector": "caption",
    "props": [
        ("caption-side", "top"), ("font-size", "22px"),
        ("font-weight", "bold"), ("text-align", "center"),
        ("background-color", "#27AE60"), ("color", "white"),
        ("padding", "10px"), ("border", "1px solid black"),
    ],
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
                                        ("font-weight", "bold"), ("text-align", "center"),
                                        ("border", "1px solid black"), ("padding", "6px")]},
        {"selector": "td", "props": [("text-align", "center"), ("border", "1px solid black"),
                                        ("padding", "5px")]},
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
    weeks_count = task_weeks(task_code)
    sites = team_df.copy()
    sites["_site_sort"] = sites["Site name"].fillna("").astype(str).str.strip().str.casefold()
    sites = sites[sites["_site_sort"] != ""].sort_values("_site_sort", kind="stable")
    sites = sites.drop(columns=["_site_sort"]).reset_index(drop=True)
    quotient, remainder = divmod(len(sites), weeks_count)
    result, start = {}, 0
    for week in range(1, weeks_count + 1):
        count = quotient + (1 if week <= remainder else 0)
        result[week] = sites.iloc[start:start + count].copy()
        start += count
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
        print(f"ℹ️ Could not load warning history; starting new tracking file: {exc}")
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
        stats = [item["Stats"][week] for item in items if week in item["Stats"]]
        target = sum(value["Target"] for value in stats)
        result = sum(value["Result"] for value in stats)
        total[f"Week {week} Target"] = target
        total[f"Week {week} Result"] = result
        total[f"Week {week} %"] = f"{round(result * 100 / target) if target else 0}%"
    rows.append(total)
    return pd.DataFrame(rows).fillna("")


def cumulative_team_progress(items, current_week):
    target = result = 0
    for item in items:
        for week in range(1, min(current_week, item["Weeks"]) + 1):
            stats = item["Stats"].get(week, {"Target": 0, "Result": 0})
            target += stats["Target"]
            result += stats["Result"]
    return target, result, round(result * 100 / target) if target else 0


def make_warning_draft(teams, today, week_number):
    entries = [
        f"{i}. {team} | {TEAM_LEADERS[team]} | Team Leader | "
        "លទ្ធផលក្រោម 50% ជាលើកទី 3 | ស្នើពិន័យ៖ កាត់ 1KI (រង់ចាំការអនុម័ត)"
        for i, team in enumerate(teams, start=1)
    ]
    return (
        "សេចក្តីព្រាងលិខិតព្រមាន — មិនទាន់មានសុពលភាព រង់ចាំនាយកពិនិត្យ និងចុះហត្ថលេខា\n\n"
        "ប្រធានបទ៖ សេចក្តីព្រាងលិខិតព្រមានចំពោះក្រុមដែលលទ្ធផលការងារក្រោម 50% ចំនួន 3 សប្តាហ៍\n"
        "ជូន៖ ក្រុមប្រឹក្សាភិបាល / នាយកពិនិត្យ និងអនុម័ត\n"
        f"កាលបរិច្ឆេទ៖ {today.strftime('%d/%m/%Y')} | សប្តាហ៍ទី {week_number}\n\n"
        "ក្រុមដែលដល់លក្ខខណ្ឌ៖\n" + "\n".join(entries) + "\n\n"
        "មូលហេតុ៖ លទ្ធផលការងារដែលបានអនុម័តស្ថិតក្រោម 50% នៃ Target ក្នុងសប្តាហ៍ដែលបានរាយ។\n"
        "សំណើ៖ សូមនាយកពិនិត្យមូលហេតុ និងសម្រេចលើការចេញលិខិតព្រមាន និងសំណើកាត់ 1KI។\n\n"
        "សេចក្តីព្រាងនេះមិនមែនជាលិខិតផ្លូវការ និងមិនអនុវត្តការកាត់ 1KI ដោយស្វ័យប្រវត្តិទេ។\n"
        "ហត្ថលេខានាយក៖ ____________________    កាលបរិច្ឆេទ៖ ______________"
    )


def send_weekly_team_reports(client, team_rows, team_chats, today, week_number, warning_state):
    warning_candidates = []
    for team in VALID_TEAMS:
        items = team_rows.get(team, [])
        chat_id = team_chats.get(team)
        if not items or not chat_id:
            if items and not chat_id:
                print(f"⚠️ No group chat ID for {team}; weekly report not sent.")
            continue

        target, result, percent = cumulative_team_progress(items, week_number)
        title = f"របាយការណ៍ {team} — Target, Result និង % តាម Site name"
        image_path = f"team_summary_{team}.png"
        dfi.export(style_table(weekly_team_table(items), title).hide(axis="index"), image_path, max_rows=-1)
        try:
            client.send_file(chat_id, image_path, caption=f"{title} — សប្តាហ៍ទី {week_number}")
            message = (
                f"សាររំលឹកប្រចាំសប្តាហ៍ទី {week_number}: {team} ({TEAM_LEADERS[team]})\n"
                f"លទ្ធផលសរុប៖ {result}/{target} ({percent}%).\n"
            )
            if target and percent < LOW_PROGRESS_THRESHOLD:
                message += "លទ្ធផលក្រោម 50%។ សូមបញ្ជាក់មូលហេតុ និងផែនការដោះស្រាយក្នុងក្រុម។"
            else:
                message += "សូមបន្តតាមដានគោលដៅ និងបញ្ចប់ការងារតាមផែនការ។"
            client.send_message(chat_id, message)
        except Exception as exc:
            print(f"⚠️ Failed to send weekly report/reminder to {team}: {exc}")

        if target and percent < LOW_PROGRESS_THRESHOLD:
            state = warning_state.setdefault(team, {"low_weeks": 0, "last_counted_week": 0, "draft_sent": False})
            if int(state.get("last_counted_week", 0)) < week_number:
                state["low_weeks"] = int(state.get("low_weeks", 0)) + 1
                state["last_counted_week"] = week_number
            if state["low_weeks"] >= 3 and not state.get("draft_sent", False):
                warning_candidates.append(team)
        else:
            state = warning_state.setdefault(team, {"low_weeks": 0, "last_counted_week": 0, "draft_sent": False})
            state["last_counted_week"] = max(int(state.get("last_counted_week", 0)), week_number)

    if warning_candidates:
        try:
            client.send_message(MAIN_GROUP_ID, make_warning_draft(warning_candidates, today, week_number))
            for team in warning_candidates:
                warning_state[team]["draft_sent"] = True
            print(f"✅ Sent unsigned warning-letter draft to main group for: {warning_candidates}")
        except Exception as exc:
            print(f"⚠️ Failed to send warning-letter draft: {exc}")


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
    cambodia_tz = timezone(timedelta(hours=7))
    now = datetime.now(cambodia_tz)
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
            required = {"Team", "Site name", "Result"}
            if not required.issubset(task_df.columns):
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

            columns = [column for column in [
                "No.", "Group task", "Branch", "Site name", "Serial", "Q'ty task/Local task",
                "Result", "Remark", "Last date record", "History Task", "Team", "Name staff",
            ] if column in task_df.columns]
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

        # Send weekly reports and reminders each Saturday, only within the four-week period.
        if today.weekday() == 5 and 1 <= current_week <= DEFAULT_WEEKS:
            send_weekly_team_reports(client, team_rows, team_chats, today, current_week, warning_state)

        if now.hour >= 12:
            rows = []
            total_target = total_result = 0
            for index, team in enumerate(VALID_TEAMS, start=1):
                target, result = overall[team]["Target"], overall[team]["Result"]
                percent = round(result * 100 / target) if target else 0
                rows.append({"No.": index, "Team": team, "Target": target,
                             "Result": result, "%": f"{percent}%"})
                total_target += target
                total_result += result
            rows.append({"No.": "TOTAL", "Team": "", "Target": total_target,
                         "Result": total_result,
                         "%": f"{round(total_result * 100 / total_target) if total_target else 0}%"})
            image_path = "overall_report.png"
            summary = pd.DataFrame(rows)
            dfi.export(style_table(summary, f"របាយការណ៍សរុប M{now.month}").hide(axis="index"), image_path, max_rows=-1)
            client.send_file(MAIN_GROUP_ID, image_path, caption=f"របាយការណ៍សរុប M{now.month}")

    save_warning_state(warning_state)


if __name__ == "__main__":
    main()
