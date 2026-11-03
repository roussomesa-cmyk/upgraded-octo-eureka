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

# ===== ការផ្ញើអូតូពេលព្រឹកទៅគ្រុបធីមនីមួយៗ =====
# START_DATE : ថ្ងៃចាប់ផ្តើម (ទម្រង់ YYYY-MM-DD) អាចកំណត់តាម env var START_DATE
# DEFAULT_WEEKS : រយៈពេលធម្មតា (សប្តាហ៍)
# TASK_WEEKS : ការងារណាដែលមានរយៈពេលខុសពីធម្មតា (ឧ. A2 = 3 សប្តាហ៍)
START_DATE = datetime.strptime(os.environ.get("START_DATE") or "2026-11-02", "%Y-%m-%d").date()
DEFAULT_WEEKS = 4
TASK_WEEKS = {"A2": 3}

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


def is_team_send_active(task_code, today):
    """true បើថ្ងៃនេះនៅក្នុងរយៈពេលផ្ញើទៅគ្រុបធីម (A2 = 3 សប្តាហ៍, ផ្សេងទៀត = 4 សប្តាហ៍)."""
    weeks = TASK_WEEKS.get(str(task_code).strip().upper(), DEFAULT_WEEKS)
    days_passed = (today - START_DATE).days
    return 0 <= days_passed < weeks * 7


def normalize_team(name):
    """CHA_TEAM01 / CHA-T01 / CHA-TEAM1 -> CHA-T01"""
    m = re.search(r"(\d+)\s*$", str(name).strip())
    return f"CHA-T{int(m.group(1)):02d}" if m else None


KHMER_DIGITS = str.maketrans("០១២៣៤៥៦៧៨៩", "0123456789")


def parse_week(value):
    """'1' / 'W2' / 'Week 3' / 'សប្តាហ៍ទី ៤' -> 1..4 (បើអានមិនបាន = None)."""
    m = re.search(r"\d+", str(value).translate(KHMER_DIGITS))
    if not m:
        return None
    n = int(m.group())
    return n if 1 <= n <= DEFAULT_WEEKS else None


def get_current_week(today):
    """គណនាសប្តាហ៍ថ្ងៃនេះដោយអូតូពី START_DATE (សប្តាហ៍ទី 1 = 7 ថ្ងៃដំបូង)."""
    return (today - START_DATE).days // 7 + 1


def select_rows_for_week(df_team, task_code, task_week_map, today):
    """ជ្រើសជួរការងាររបស់ធីមដែលត្រូវនឹងសប្តាហ៍ថ្ងៃនេះ ដោយអូតូ តាមលំដាប់អាទិភាព៖
    1) Column 'Week' ក្នុង Sheet ការងារនោះ (ជួរនីមួយៗមានសប្តាហ៍ផ្ទាល់ខ្លួន)
    2) Column 'Week' ក្នុង Sheet 'Team chat IDs' (មួយការងារ = មួយសប្តាហ៍)
    3) គ្មានព័ត៌មានសប្តាហ៍ទេ → ប្រើច្បាប់ចាស់ (A2 = 3 សប្តាហ៍, ផ្សេងទៀត = 4 សប្តាហ៍)
    """
    week_now = get_current_week(today)
    if "Week" in df_team.columns:
        weeks = df_team["Week"].map(parse_week)
        if weeks.notna().any():
            return df_team[weeks == week_now]
    if task_code in task_week_map:
        return df_team if task_week_map[task_code] == week_now else df_team.iloc[0:0]
    return df_team if is_team_send_active(task_code, today) else df_team.iloc[0:0]


def fetch_csv(sheet_name_or_gid):
    url =f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet={sheet_name_or_gid}"
    headers = {"User-Agent": "Mozilla/5.0"}
    try:
        res = requests.get(url, headers=headers, timeout=15)
    except requests.RequestException as e:
        print(f"⚠️ Request failed for sheet '{sheet_name_or_gid}': {e}")
        return None

    if res.status_code == 200:
        lines = res.text.splitlines()
        header_row_idx = 0
        header_keywords = {"site name", "team", "sheet", "chatid", "no", "no."}

        for idx, line in enumerate(lines[:15]):
            cells = [c.strip().strip('"').lower() for c in line.split(",")]
            if any(cell in header_keywords for cell in cells):
                header_row_idx = idx
                break

        df = pd.read_csv(io.StringIO(res.text), skiprows=header_row_idx)
        df = df.dropna(how="all")
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
            elif c_lower in ("week", "សប្តាហ៍", "សប្តាហ៍ទី"):
                rename_dict[col] = "Week"

        if rename_dict:
            df = df.rename(columns=rename_dict)

        if df.columns.duplicated().any():
            dup_cols = df.columns[df.columns.duplicated()].unique().tolist()
            print(f"⚠️ Duplicate columns found and removed (keeping first occurrence): {dup_cols}")
            df = df.loc[:, ~df.columns.duplicated()]

        return df

    print(f"⚠️ Failed to fetch sheet '{sheet_name_or_gid}': HTTP {res.status_code}")
    return None


def get_task_title_from_sheet(sheet_name_or_gid):
    url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet={sheet_name_or_gid}"
    headers = {"User-Agent": "Mozilla/5.0"}
    try:
        res = requests.get(url, headers=headers, timeout=15)
    except requests.RequestException as e:
        print(f"⚠️ Request failed for title lookup '{sheet_name_or_gid}': {e}")
        return f"Task {sheet_name_or_gid}"

    if res.status_code == 200:
        lines = res.text.splitlines()
        if lines:
            first_line = lines[0].replace('"', "").strip()
            first_cell = first_line.split(",")[0].strip()
            if first_cell and not first_cell.isdigit():
                return first_cell

    print(f"⚠️ Could not extract title for '{sheet_name_or_gid}', using fallback.")
    return f"Task {sheet_name_or_gid}"


def style_detail_table(df, title):
    styler = df.style.set_caption(title).set_table_styles([
        COMMON_CAPTION_STYLE,
        {"selector": "th", "props": [("background-color", "#369388"), ("color", "black"), ("font-weight", "bold"), ("text-align", "center"), ("border", "1px solid black"), ("padding", "6px")]},
        {"selector": "td", "props": [("text-align", "center"), ("border", "1px solid black"), ("padding", "5px")]},
    ])
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


def style_overall_summary(df, title):
    styler = df.style.set_caption(title).set_table_styles([
        COMMON_CAPTION_STYLE,
        {"selector": "caption", "props": [("caption-side", "top"), ("font-size", "22px"), ("font-weight", "bold"), ("text-align", "center"), ("background-color", "#2EA44E"), ("color", "white"), ("padding", "10px"), ("border", "1px solid black")]},
        {"selector": "th", "props": [("background-color", "#2EA44E"), ("color", "white"), ("font-weight", "bold"), ("text-align", "center"), ("border", "1px solid black"), ("padding", "6px")]},
        {"selector": "td", "props": [("text-align", "center"), ("border", "1px solid black"), ("padding", "5px")]},
    ])

    def apply_total_style(row):
        if row.name == len(df) - 1:
            return ["font-weight: bold; background-color: #F2F2F2;"] * len(row)
        return [""] * len(row)

    return styler.apply(apply_total_style, axis=1)


def build_pending_list_text(df_single_team, team, task_title):
    """រាយ Site name (+ Serial បើមាន) នៃការងារមិនទាន់ធ្វើ ជាអត្ថបទ (គ្មានចំណងជើងស្ទួន)."""
    if "Site name" not in df_single_team.columns:
        return None
    has_serial = "Serial" in df_single_team.columns

    lines_data = []
    for _, row in df_single_team.iterrows():
        site = str(row.get("Site name", "")).strip()
        if not site:
            continue
        if has_serial:
            serial = str(row.get("Serial", "")).strip()
            if serial and serial.lower() not in ("nan", "none", "#n/a", "n/a"):
                lines_data.append(f"{team} {site} - {serial}")
            else:
                lines_data.append(f"{team} {site}")
        else:
            lines_data.append(f"{team} {site}")

    if not lines_data:
        return None

    return "\n".join(lines_data)


def load_team_chat_ids(df_mapping):
    """អាន ChatID គ្រុបធីមពី Sheet 'Team chat IDs' (Column A = ឈ្មោះធីម, Column B = ChatID)."""
    team_chat_ids = {}
    if df_mapping is None:
        return team_chat_ids
    if "Column A" not in df_mapping.columns or "Column B" not in df_mapping.columns:
        print("⚠️ 'Column A' / 'Column B' not found in 'Team chat IDs' — team group sending disabled.")
        return team_chat_ids

    for _, row in df_mapping.dropna(subset=["Column A", "Column B"]).iterrows():
        team = normalize_team(row["Column A"])
        if team not in VALID_TEAMS:
            continue
        try:
            team_chat_ids[team] = int(float(str(row["Column B"]).strip()))
        except ValueError:
            continue
    return team_chat_ids


def count_pending_by_week(df_team, task_code, task_week_map, today):
    """រាប់ site មិនទាន់ approve របស់ធីមមួយក្នុងការងារមួយ ចែកតាមសប្តាហ៍ទី 1..4.
    ប្រភពសប្តាហ៍ (តាមអាទិភាព): column 'Week' ក្នុង Sheet ការងារ -> 'Week' ក្នុង 'Team chat IDs' -> សប្តាហ៍បច្ចុប្បន្ន"""
    counts = {w: 0 for w in range(1, DEFAULT_WEEKS + 1)}
    pending = df_team[df_team["Result"].astype(str).str.strip().str.lower() != "approved"]
    if pending.empty:
        return counts
    if "Week" in pending.columns and pending["Week"].map(parse_week).notna().any():
        for w in pending["Week"].map(parse_week).dropna():
            counts[int(w)] += 1
    elif task_code in task_week_map:
        counts[task_week_map[task_code]] += len(pending)
    elif is_team_send_active(task_code, today):
        wk = get_current_week(today)
        if wk in counts:
            counts[wk] += len(pending)
    return counts


def send_team_task_tables(client, team_task_rows, team_chat_ids, today, shift_title, week_now):
    """ពេលព្រឹក៖ ផ្ញើ 'របាយការណ៍ CHA-Txx សប្តាហ៍ទី N' មួយសារទៅគ្រុបធីមនីមួយៗ.
    ជួរ = ការងារ (No., Task), ជួរឈរ = សប្តាហ៍ទី 1..4 និង សរុបការងារទាំងអស់ (site មិនទាន់ធ្វើ)"""
    for team in VALID_TEAMS:
        rows = team_task_rows.get(team, [])
        if not rows:
            print(f"ℹ️ {team}: no tasks for team-group table — skipped.")
            continue
        team_gid = team_chat_ids.get(team)
        if not team_gid:
            print(f"⚠️ No team group ID for {team} — skipped.")
            continue

        table_rows = []
        for idx, r in enumerate(rows, start=1):
            wk = r["Weeks"]
            row = {"No.": idx, "Task": r["Task"]}
            for w in range(1, DEFAULT_WEEKS + 1):
                row[f"សប្តាហ៍ទី {w}"] = wk[w]
            row["សរុបការងារទាំងអស់"] = sum(wk.values())
            table_rows.append(row)

        total_row = {"No.": "", "Task": "TOTAL"}
        for w in range(1, DEFAULT_WEEKS + 1):
            total_row[f"សប្តាហ៍ទី {w}"] = sum(r["Weeks"][w] for r in rows)
        total_row["សរុបការងារទាំងអស់"] = sum(sum(r["Weeks"].values()) for r in rows)
        table_rows.append(total_row)

        df_team = pd.DataFrame(table_rows)
        title = f"របាយការណ៍ {team} សប្តាហ៍ទី {week_now}"
        styled = style_overall_summary(df_team, title)
        img_path = f"team_summary_{team}.png"
        dfi.export(styled.hide(axis="index"), img_path, max_rows=-1)

        try:
            client.send_file(team_gid, img_path, caption=f"{title} - {shift_title}")
            print(f"✅ Sent team table to {team} group ({team_gid}).")
        except Exception as e:
            print(f"⚠️ Failed to send team table to {team} ({team_gid}): {e}")


def main():
    api_id_raw = os.environ.get("TELEGRAM_API_ID")
    api_hash = os.environ.get("TELEGRAM_API_HASH")
    session_str = os.environ.get("TELEGRAM_SESSION")

    if not api_id_raw:
        raise ValueError("TELEGRAM_API_ID is missing or empty in environment/secrets")
    try:
        api_id = int(api_id_raw)
    except ValueError:
        raise ValueError(f"TELEGRAM_API_ID is not a valid integer: '{api_id_raw}'")

    if not api_hash:
        raise ValueError("TELEGRAM_API_HASH is missing or empty in environment/secrets")
    if not session_str:
        raise ValueError("TELEGRAM_SESSION is missing or empty in environment/secrets")

    df_mapping = fetch_csv("Team%20chat%20IDs")
    task_chat_ids = {}
    task_week_map = {}  # ជម្រើស៖ ការងារ -> សប្តាហ៍ (ពី column 'Week' ក្នុង Sheet 'Team chat IDs')

    if df_mapping is not None and "Sheet" in df_mapping.columns:
        df_clean = df_mapping.dropna(subset=["Sheet", "ChatID"])
        for _, row in df_clean.iterrows():
            code = str(row["Sheet"]).strip()
            try:
                task_chat_ids[code] = int(float(str(row["ChatID"]).strip()))
            except ValueError:
                continue
            if "Week" in df_clean.columns:
                wk = parse_week(row["Week"])
                if wk:
                    task_week_map[code] = wk

    if not task_chat_ids:
        print("⚠️ WARNING: task_chat_ids is empty — check the 'Team chat IDs' sheet. No reports will be sent.")
        return

    print(f"ℹ️ Loaded {len(task_chat_ids)} task->chat mappings: {task_chat_ids}")

    team_chat_ids = load_team_chat_ids(df_mapping)
    print(f"ℹ️ Loaded {len(team_chat_ids)} team group chat IDs: {team_chat_ids}")

    cambodia_tz = timezone(timedelta(hours=7))
    now = datetime.now(cambodia_tz)
    today = now.date()
    is_morning = now.hour < 12
    shift_title = "Morning Shift" if is_morning else "Evening Shift"
    print(f"ℹ️ Shift: {shift_title} — Summary/Overall reports will be sent to MAIN_GROUP_ID only in the Evening Shift.")

    overall_stats = {team: {"Target": 0, "Approved": 0, "NotApproved": 0} for team in VALID_TEAMS}

    # សម្រាប់តារាងព្រឹកដែលផ្ញើទៅគ្រុបធីម៖ ការងាររបស់ធីមនីមួយៗក្នុងសប្តាហ៍បច្ចុប្បន្ន (រកឃើញដោយអូតូ)
    week_now = get_current_week(today)
    team_task_rows = {team: [] for team in VALID_TEAMS}
    print(f"ℹ️ Current week (auto from START_DATE {START_DATE}): {week_now}")

    with TelegramClient(StringSession(session_str), api_id, api_hash) as client:
        for task_code, chat_id in task_chat_ids.items():
            sheet_sub_title = get_task_title_from_sheet(task_code)
            task_title = f"{task_code}. {sheet_sub_title}"

            df_task = fetch_csv(task_code)
            if df_task is None or df_task.empty:
                print(f"⚠️ Skipping '{task_code}': sheet is empty or failed to load.")
                continue

            cols_to_show = ["No.", "Group task", "Branch", "Site name", "Serial", "Q'ty task/Local task", "Result", "Remark", "Last date record", "History Task", "Team"]
            available_cols = [c for c in cols_to_show if c in df_task.columns]

            if available_cols and "Team" in df_task.columns:
                df_detail = df_task[available_cols].copy()

                df_detail["Team"] = df_detail["Team"].astype(str).str.strip()
                df_detail = df_detail[df_detail["Team"].isin(VALID_TEAMS)]

                if "Group task" in df_detail.columns:
                    df_detail = df_detail[df_detail["Group task"].notna() & (~df_detail["Group task"].astype(str).str.strip().str.lower().isin(["", "nan", "none", "#n/a", "n/a"]))]

                if "Site name" in df_detail.columns:
                    df_detail = df_detail[df_detail["Site name"].notna() & (~df_detail["Site name"].astype(str).str.strip().str.lower().isin(["", "nan", "none", "#n/a", "n/a"]))]

                # ពេលព្រឹក៖ ត្រូវការ column "Result" ដើម្បីដឹងច្បាស់ថា site ណាមិនទាន់ approve
                # បើ column នេះបាត់ (header ទទេក្នុង Sheet) កុំផ្ញើទិន្នន័យខុស — ស្រង់ចេញឲ្យទទេវិញ ហើយ log ព្រមាន
                if is_morning:
                    if "Result" in df_detail.columns:
                        df_detail = df_detail[df_detail["Result"].astype(str).str.strip().str.lower() != "approved"]
                    else:
                        print(f"⚠️ '{task_code}': 'Result' column is missing/empty header — cannot determine approval status. Skipping morning detail images for this sheet to avoid sending incorrect data. Please add a 'Result' header in this sheet.")
                        df_detail = df_detail.iloc[0:0]

                df_detail = df_detail.fillna("")
                df_detail = df_detail.replace(to_replace=r"^(?i:nan|none|#n/a|n/a)$", value="", regex=True)

                for team in VALID_TEAMS:
                    df_single_team = df_detail[df_detail["Team"] == team].copy()
                    if not df_single_team.empty:
                        styled_detail = style_detail_table(df_single_team, f"{task_title} ({team})")
                        img_detail_path = f"detail_{task_code}_{team}.png"
                        dfi.export(styled_detail.hide(axis="index"), img_detail_path, max_rows=-1)

                        caption_text = f"ការងារមិនទាន់ធ្វើ {team} ({task_title})" if is_morning else f"ការងារសរុប {team} ({task_title})"
                        full_caption = f"{caption_text} - {shift_title}"

                        if is_morning:
                            pending_text = build_pending_list_text(df_single_team, team, task_title)
                            if pending_text:
                                full_caption = f"{full_caption}\n\n{pending_text}"

                        # 1) ផ្ញើទៅគ្រុបការងារ (ដូចដើម)
                        client.send_file(chat_id, img_detail_path, caption=full_caption)
                    else:
                        print(f"ℹ️ '{task_code}' / team '{team}': no pending items (all approved or no data) — skipped, no image sent.")
            else:
                print(f"⚠️ '{task_code}': missing 'Team' column or no recognizable columns — no detail images sent.")

            rows = []
            tot_target = tot_approved = tot_not_approved = tot_remain = 0

            for idx, team in enumerate(VALID_TEAMS, start=1):
                target_site = approved = not_approved = 0
                if "Team" in df_task.columns and "Result" in df_task.columns:
                    df_clean_task = df_task.dropna(subset=["Team"]).copy()
                    df_clean_task["Team"] = df_clean_task["Team"].astype(str).str.strip()
                    df_team = df_clean_task[df_clean_task["Team"] == team].copy()

                    target_site = len(df_team)
                    approved = len(df_team[df_team["Result"].astype(str).str.strip().str.lower() == "approved"])
                    not_approved = len(df_team[df_team["Result"].astype(str).str.strip().str.lower() == "not approved"])

                    # ពេលព្រឹក៖ រៀបចំជួរតារាងសម្រាប់គ្រុបធីម (site មិនទាន់ធ្វើ ចែកតាមសប្តាហ៍ 1..4)
                    if is_morning:
                        wk_counts = count_pending_by_week(df_team, task_code, task_week_map, today)
                        if target_site > 0:
                            team_task_rows[team].append({"Task": task_code, "Weeks": wk_counts})

                remain = target_site - (approved + not_approved)
                pct_val = f"{int(round((approved / target_site) * 100))}%" if target_site > 0 else "0%"

                overall_stats[team]["Target"] += target_site
                overall_stats[team]["Approved"] += approved
                overall_stats[team]["NotApproved"] += not_approved

                tot_target += target_site
                tot_approved += approved
                tot_not_approved += not_approved
                tot_remain += remain

                rows.append({"No": idx, "Team": team, "Target Site": target_site, "Approved": approved, "Not Approved": not_approved, "%": pct_val, "Remain": remain, "Remark": ""})

            if tot_target == 0:
                print(f"⚠️ Skipping summary for '{task_code}': no rows matched any VALID_TEAMS ({VALID_TEAMS}). Check 'Team' column values in this sheet.")
                continue

            if is_morning:
                print(f"ℹ️ '{task_code}': Morning shift — summary image NOT sent to MAIN_GROUP_ID.")
            else:
                tot_pct = f"{int(round((tot_approved / tot_target) * 100))}%" if tot_target > 0 else "0%"
                total_row = {"No": "", "Team": "", "Target Site": tot_target, "Approved": tot_approved, "Not Approved": tot_not_approved, "%": tot_pct, "Remain": tot_remain, "Remark": ""}

                df_summary = pd.DataFrame([total_row] + rows)
                df_summary.columns = pd.MultiIndex.from_tuples([("", "No"), ("", "Team"), ("", "Target Site"), ("Result", "Approved"), ("Result", "Not Approved"), ("", "%"), ("", "Remain"), ("", "Remark")])

                styled_summary = style_task_summary(df_summary, task_title)
                img_summary_path = f"summary_{task_code}.png"
                dfi.export(styled_summary.hide(axis="index"), img_summary_path, max_rows=-1)

                client.send_file(MAIN_GROUP_ID, img_summary_path, caption=f"របាយការណ៍សង្ខេប {task_title} ({shift_title})")

        if is_morning:
            if 1 <= week_now <= DEFAULT_WEEKS:
                send_team_task_tables(client, team_task_rows, team_chat_ids, today, shift_title, week_now)
            else:
                print(f"ℹ️ Week {week_now} is outside 1..{DEFAULT_WEEKS} — team-group tables NOT sent.")
            print("ℹ️ Morning shift — overall report NOT sent to MAIN_GROUP_ID.")
        else:
            title_3 = f"Report Plan Power M{now.month}"
            overall_rows = []
            sum_target = sum_approved = sum_not_approved = sum_remain = 0

            for idx, team in enumerate(VALID_TEAMS, start=1):
                t = overall_stats[team]["Target"]
                a = overall_stats[team]["Approved"]
                na = overall_stats[team]["NotApproved"]
                r = t - (a + na)
                pct = f"{int(round((a / t) * 100))}%" if t > 0 else "0%"

                sum_target += t
                sum_approved += a
                sum_not_approved += na
                sum_remain += r

                overall_rows.append({"No": idx, "Branch": team, "Target Site": t, "Approved": a, "Not Approved": na, "%": pct, "Remain": r})

            if sum_target > 0:
                sum_pct = f"{int(round((sum_approved / sum_target) * 100))}%" if sum_target > 0 else "0%"
                overall_rows.append({"No": "TOTAL", "Branch": "", "Target Site": sum_target, "Approved": sum_approved, "Not Approved": sum_not_approved, "%": sum_pct, "Remain": sum_remain})

                df_overall = pd.DataFrame(overall_rows)
                styled_overall = style_overall_summary(df_overall, title_3)
                img_overall_path = "overall_report.png"
                dfi.export(styled_overall.hide(axis="index"), img_overall_path, max_rows=-1)

                client.send_file(MAIN_GROUP_ID, img_overall_path, caption=f"របាយការណ៍សរុបរួម {title_3} - {shift_title}")


if __name__ == "__main__":
    main()
