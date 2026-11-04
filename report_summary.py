from datetime import datetime, timezone, timedelta
import io
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

COMMON_CAPTION_STYLE = {
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
    url = (
        f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}"
        f"/gviz/tq?tqx=out:csv&sheet={sheet_name_or_gid}"
    )
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
    header_keywords = {"site name", "team", "sheet", "chatid", "no", "no."}
    for idx, line in enumerate(lines[:15]):
        cells = [cell.strip().strip('"').lower() for cell in line.split(",")]
        if any(cell in header_keywords for cell in cells):
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
    if df.columns.duplicated().any():
        df = df.loc[:, ~df.columns.duplicated()]
    return df


def get_task_title_from_sheet(sheet_name_or_gid):
    url = (
        f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}"
        f"/gviz/tq?tqx=out:csv&sheet={sheet_name_or_gid}"
    )
    try:
        response = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=15)
        if response.status_code == 200 and response.text.splitlines():
            first = response.text.splitlines()[0].replace('"', "").split(",")[0].strip()
            if first and not first.isdigit():
                return first
    except requests.RequestException as exc:
        print(f"⚠️ Title lookup failed for '{sheet_name_or_gid}': {exc}")
    return f"Task {sheet_name_or_gid}"


def style_table(df, title):
    return df.style.set_caption(title).set_table_styles([
        COMMON_CAPTION_STYLE,
        {"selector": "th", "props": [("background-color", "#369388"), ("color", "black"),
                                        ("font-weight", "bold"), ("text-align", "center"),
                                        ("border", "1px solid black"), ("padding", "6px")]},
        {"selector": "td", "props": [("text-align", "center"), ("border", "1px solid black"),
                                        ("padding", "5px")]},
    ])


def load_team_chat_ids(df_mapping):
    result = {}
    if df_mapping is None or not {"Column A", "Column B"}.issubset(df_mapping.columns):
        print("⚠️ Team chat columns 'Column A' / 'Column B' not found; team reports disabled.")
        return result
    for _, row in df_mapping.dropna(subset=["Column A", "Column B"]).iterrows():
        team = normalize_team(row["Column A"])
        if team not in VALID_TEAMS:
            continue
        try:
            result[team] = int(float(str(row["Column B"]).strip()))
        except (TypeError, ValueError):
            continue
    return result


def split_sites_into_weeks(df_team, task_code):
    """Sort Site name and split its rows into equal, consecutive weekly target groups.

    Any remainder is assigned one row at a time to the earliest weeks. Thus, for
    example, 10 sites over four weeks are divided 3, 3, 2, 2. A2 is divided into
    three groups. The original sheet order breaks ties between identical names.
    """
    weeks_count = task_weeks(task_code)
    if "Site name" not in df_team.columns:
        return {week: df_team.iloc[0:0].copy() for week in range(1, weeks_count + 1)}

    sites = df_team.copy()
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


def weekly_result(week_df):
    if week_df.empty or "Result" not in week_df.columns:
        return 0
    return int(week_df["Result"].fillna("").astype(str).str.strip().str.casefold().eq("approved").sum())


def build_weekly_team_rows(team_task_rows):
    rows = []
    for idx, item in enumerate(team_task_rows, start=1):
        row = {"No.": idx, "Task": item["Task"]}
        for week in range(1, item["Weeks"] + 1):
            stats = item["Stats"][week]
            row[f"សប្តាហ៍ទី {week} Target"] = stats["Target"]
            row[f"សប្តាហ៍ទី {week} Result"] = stats["Result"]
            row[f"សប្តាហ៍ទី {week} %"] = f"{stats['Percent']}%"
        rows.append(row)

    if not rows:
        return pd.DataFrame()
    total = {"No.": "", "Task": "TOTAL"}
    for week in range(1, DEFAULT_WEEKS + 1):
        stats = [item["Stats"][week] for item in team_task_rows if week in item["Stats"]]
        target = sum(value["Target"] for value in stats)
        result = sum(value["Result"] for value in stats)
        total[f"សប្តាហ៍ទី {week} Target"] = target
        total[f"សប្តាហ៍ទី {week} Result"] = result
        total[f"សប្តាហ៍ទី {week} %"] = f"{round(result * 100 / target) if target else 0}%"
    rows.append(total)
    return pd.DataFrame(rows).fillna("")


def send_team_reports(client, team_task_rows, team_chat_ids, today, week_now):
    for team in VALID_TEAMS:
        items = team_task_rows.get(team, [])
        chat_id = team_chat_ids.get(team)
        if not items or not chat_id:
            if items and not chat_id:
                print(f"⚠️ No group chat ID for {team}; weekly report not sent.")
            continue
        table = build_weekly_team_rows(items)
        title = f"របាយការណ៍ {team} បែងចែកតាម Site name — សប្តាហ៍ទី {week_now}"
        image_path = f"team_summary_{team}.png"
        dfi.export(style_table(table, title).hide(axis="index"), image_path, max_rows=-1)
        try:
            client.send_file(chat_id, image_path, caption=title)
            print(f"✅ Sent weekly target/result/% report to {team} group ({chat_id}).")
        except Exception as exc:
            print(f"⚠️ Failed to send weekly report to {team}: {exc}")

        # On Saturdays, ask the team leader to explain progress below 50%.
        if today.weekday() == 5 and 1 <= week_now <= DEFAULT_WEEKS:
            targets = results = 0
            for item in items:
                for week in range(1, min(week_now, item["Weeks"]) + 1):
                    stats = item["Stats"].get(week, {"Target": 0, "Result": 0})
                    targets += stats["Target"]
                    results += stats["Result"]
            percent = round(results * 100 / targets) if targets else 0
            if targets and percent < LOW_PROGRESS_THRESHOLD:
                leader = TEAM_LEADERS.get(team, "មេក្រុម")
                alert = (
                    f"ការរំលឹកលើកទី១៖ {team} ({leader}) មានលទ្ធផល {percent}% "
                    f"ក្រោមគោលដៅ 50% គិតត្រឹមសប្តាហ៍ទី {week_now}.\n"
                    "សូមក្រុមបញ្ជាក់មូលហេតុ និងផែនការដោះស្រាយនៅក្នុងក្រុមនេះ។ "
                    "សូមអ្នកគ្រប់គ្រងពិនិត្យ និងសម្រេចជំហានបន្ទាប់។"
                )
                try:
                    client.send_message(chat_id, alert)
                except Exception as exc:
                    print(f"⚠️ Failed to send Saturday progress reminder to {team}: {exc}")


def main():
    api_id_raw = os.environ.get("TELEGRAM_API_ID")
    api_hash = os.environ.get("TELEGRAM_API_HASH")
    session_str = os.environ.get("TELEGRAM_SESSION")
    if not api_id_raw:
        raise ValueError("TELEGRAM_API_ID is missing or empty in environment/secrets")
    try:
        api_id = int(api_id_raw)
    except ValueError as exc:
        raise ValueError(f"TELEGRAM_API_ID is not a valid integer: '{api_id_raw}'") from exc
    if not api_hash:
        raise ValueError("TELEGRAM_API_HASH is missing or empty in environment/secrets")
    if not session_str:
        raise ValueError("TELEGRAM_SESSION is missing or empty in environment/secrets")

    mapping = fetch_csv("Team%20chat%20IDs")
    task_chat_ids = {}
    if mapping is not None and {"Sheet", "ChatID"}.issubset(mapping.columns):
        for _, row in mapping.dropna(subset=["Sheet", "ChatID"]).iterrows():
            code = str(row["Sheet"]).strip()
            try:
                task_chat_ids[code] = int(float(str(row["ChatID"]).strip()))
            except (TypeError, ValueError):
                continue
    if not task_chat_ids:
        print("⚠️ No task-to-chat mappings found in 'Team chat IDs'; nothing to send.")
        return

    team_chat_ids = load_team_chat_ids(mapping)
    cambodia_tz = timezone(timedelta(hours=7))
    now = datetime.now(cambodia_tz)
    today = now.date()
    week_now = get_current_week(today)
    is_morning = now.hour < 12
    print(f"ℹ️ Current week from START_DATE {START_DATE}: {week_now}")

    team_task_rows = {team: [] for team in VALID_TEAMS}
    overall_stats = {team: {"Target": 0, "Approved": 0} for team in VALID_TEAMS}

    with TelegramClient(StringSession(session_str), api_id, api_hash) as client:
        for task_code, task_chat_id in task_chat_ids.items():
            df_task = fetch_csv(task_code)
            if df_task is None or df_task.empty:
                print(f"⚠️ Skipping '{task_code}': sheet is empty or failed to load.")
                continue
            if not {"Team", "Site name"}.issubset(df_task.columns):
                print(f"⚠️ Skipping '{task_code}': 'Team' or 'Site name' column is missing.")
                continue
            if "Result" not in df_task.columns:
                print(f"⚠️ Skipping '{task_code}': 'Result' column is missing; cannot calculate progress.")
                continue

            df_task = df_task.copy()
            df_task["Team"] = df_task["Team"].fillna("").astype(str).str.strip()
            df_task["Site name"] = df_task["Site name"].fillna("").astype(str).str.strip()
            df_task = df_task[df_task["Team"].isin(VALID_TEAMS) & df_task["Site name"].ne("")]
            title = f"{task_code}. {get_task_title_from_sheet(task_code)}"

            for team in VALID_TEAMS:
                df_team = df_task[df_task["Team"] == team].copy()
                if df_team.empty:
                    continue
                week_groups = split_sites_into_weeks(df_team, task_code)
                stats = {}
                for week, week_df in week_groups.items():
                    target = len(week_df)
                    approved = weekly_result(week_df)
                    stats[week] = {
                        "Target": target,
                        "Result": approved,
                        "Percent": round(approved * 100 / target) if target else 0,
                    }
                team_task_rows[team].append({
                    "Task": task_code,
                    "Weeks": task_weeks(task_code),
                    "Stats": stats,
                })

                total_target = len(df_team)
                total_approved = weekly_result(df_team)
                overall_stats[team]["Target"] += total_target
                overall_stats[team]["Approved"] += total_approved

            # Keep sending the familiar task detail image to its configured task group.
            detail_columns = [c for c in [
                "No.", "Group task", "Branch", "Site name", "Serial",
                "Q'ty task/Local task", "Result", "Remark", "Last date record",
                "History Task", "Team", "Name staff",
            ] if c in df_task.columns]
            details = df_task[detail_columns].copy()
            if is_morning:
                details = details[details["Result"].astype(str).str.strip().str.casefold() != "approved"]
            if not details.empty:
                image_path = f"detail_{task_code}.png"
                dfi.export(style_table(details, title).hide(axis="index"), image_path, max_rows=-1)
                try:
                    client.send_file(task_chat_id, image_path, caption=title)
                except Exception as exc:
                    print(f"⚠️ Failed to send detail report for {task_code}: {exc}")

        if is_morning:
            if 1 <= week_now <= DEFAULT_WEEKS:
                send_team_reports(client, team_task_rows, team_chat_ids, today, week_now)
            else:
                print(f"ℹ️ Week {week_now} is outside the four-week reporting period.")
        else:
            summary_rows = []
            total_target = total_result = 0
            for index, team in enumerate(VALID_TEAMS, start=1):
                target = overall_stats[team]["Target"]
                result = overall_stats[team]["Approved"]
                percent = round(result * 100 / target) if target else 0
                summary_rows.append({"No.": index, "Team": team, "Target": target,
                                     "Result": result, "%": f"{percent}%"})
                total_target += target
                total_result += result
            summary_rows.append({"No.": "TOTAL", "Team": "", "Target": total_target,
                                 "Result": total_result,
                                 "%": f"{round(total_result * 100 / total_target) if total_target else 0}%"})
            summary = pd.DataFrame(summary_rows)
            image_path = "overall_report.png"
            dfi.export(style_table(summary, f"របាយការណ៍សរុប M{now.month}").hide(axis="index"), image_path, max_rows=-1)
            client.send_file(MAIN_GROUP_ID, image_path, caption=f"របាយការណ៍សរុប M{now.month} — Evening Shift")


if __name__ == "__main__":
    main()
