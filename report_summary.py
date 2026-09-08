from datetime import datetime, timezone, timedelta
import io
import os
import dataframe_image as dfi
import pandas as pd
import requests
from telethon.sessions import StringSession
from telethon.sync import TelegramClient

# ទាញយកតម្លៃពី GitHub Secrets (បើគ្មាន ឬទទេ វានឹងប្រើតម្លៃ Default)
SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID") or "1PmMSqfeBWhYJe5dMv3PrLOFKc2YmLYP8BdCvf9FyZX4"
MAIN_GROUP_ID = int(os.environ.get("NOTIFY_GROUP_ID") or "-1001853372580")

VALID_TEAMS = [f"CHA-T0{i}" for i in range(1, 8)]

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

def fetch_csv(sheet_name_or_gid):
    url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet={sheet_name_or_gid}"
    headers = {"User-Agent": "Mozilla/5.0"}
    try:
        res = requests.get(url, headers=headers, timeout=15)
    except requests.RequestException as e:
        print(f"⚠️ Request failed for sheet '{sheet_name_or_gid}': {e}")
        return None

    if res.status_code == 200:
        lines = res.text.splitlines()
        header_row_idx = 0

        for idx, line in enumerate(lines[:15]):
            line_lower = line.lower()
            if "site name" in line_lower or "team" in line_lower:
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

        if rename_dict:
            df = df.rename(columns=rename_dict)

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
        for line in lines[:4]:
            clean_line = line.replace('"', "").strip()
            clean_lower = clean_line.lower()

            if (
                clean_line
                and not clean_lower.startswith("no")
                and "site name" not in clean_lower
                and "team" not in clean_lower
                and len(clean_line) > 3
            ):
                parts = [p.strip() for p in clean_line.split(",") if p.strip()]
                if len(parts) == 1 and not parts[0].isdigit():
                    return parts[0]
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

def main():
    # -- validate Telegram credentials up front --
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

    if df_mapping is not None and "Sheet" in df_mapping.columns:
        df_clean = df_mapping.dropna(subset=["Sheet", "ChatID"])
        for _, row in df_clean.iterrows():
            code = str(row["Sheet"]).strip()
            try:
                task_chat_ids[code] = int(float(str(row["ChatID"]).strip()))
            except ValueError:
                continue

    if not task_chat_ids:
        print("⚠️ WARNING: task_chat_ids is empty — check the 'Team chat IDs' sheet (columns 'Sheet' / 'ChatID', or the sheet failed to load). No reports will be sent.")
        return

    cambodia_tz = timezone(timedelta(hours=7))
    now = datetime.now(cambodia_tz)
    is_morning = now.hour < 12
    shift_title = "Morning Shift" if is_morning else "Evening Shift"

    overall_stats = {team: {"Target": 0, "Approved": 0, "NotApproved": 0} for team in VALID_TEAMS}

    with TelegramClient(StringSession(session_str), api_id, api_hash) as client:
        for task_code, chat_id in task_chat_ids.items():
            sheet_sub_title = get_task_title_from_sheet(task_code)
            task_title = f"{task_code}. {sheet_sub_title}"

            df_task = fetch_csv(task_code)
            if df_task is None or df_task.empty:
                continue

            cols_to_show = ["No.", "Group task", "Branch", "Site name", "Q'ty task/Local task", "Result", "Remark", "Last date record", "History Task", "Team"]
            available_cols = [c for c in cols_to_show if c in df_task.columns]

            if available_cols and "Team" in df_task.columns:
                df_detail = df_task[available_cols].copy()

                df_detail["Team"] = df_detail["Team"].astype(str).str.strip()
                df_detail = df_detail[df_detail["Team"].isin(VALID_TEAMS)]

                if "Group task" in df_detail.columns:
                    df_detail = df_detail[df_detail["Group task"].notna() & (~df_detail["Group task"].astype(str).str.strip().str.lower().isin(["", "nan", "none", "#n/a", "n/a"]))]

                if "Site name" in df_detail.columns:
                    df_detail = df_detail[df_detail["Site name"].notna() & (~df_detail["Site name"].astype(str).str.strip().str.lower().isin(["", "nan", "none", "#n/a", "n/a"]))]

                if is_morning and "Result" in df_detail.columns:
                    df_detail = df_detail[df_detail["Result"].astype(str).str.strip().str.lower() != "approved"]

                df_detail = df_detail.fillna("")
                df_detail = df_detail.replace(to_replace=r"^(?i:nan|none|#n/a|n/a)$", value="", regex=True)

                for team in VALID_TEAMS:
                    df_single_team = df_detail[df_detail["Team"] == team].copy()
                    if not df_single_team.empty:
                        styled_detail = style_detail_table(df_single_team, f"{task_title} ({team})")
                        img_detail_path = f"detail_{task_code}_{team}.png"
                        dfi.export(styled_detail.hide(axis="index"), img_detail_path, max_rows=-1)

                        caption_text = f"ការងារត្រូវមិនទាន់ធ្វើ {team} ({task_title})" if is_morning else f"ការងារសរុប {team} ({task_title})"
                        client.send_file(chat_id, img_detail_path, caption=f"{caption_text} - {shift_title}")

            rows = []
            tot_target = tot_approved = tot_not_approved = tot_remain = 0

            for idx, team in enumerate(VALID_TEAMS, start=1):
                target_site = approved = not_approved = 0
                if "Team" in df_task.columns and "Result" in df_task.columns:
                    df_clean_task = df_task.dropna(subset=["Team"])
                    df_clean_task["Team"] = df_clean_task["Team"].astype(str).str.strip()
                    df_team = df_clean_task[df_clean_task["Team"] == team].copy()

                    target_site = len(df_team)
                    approved = len(df_team[df_team["Result"].astype(str).str.strip().str.lower() == "approved"])
                    not_approved = len(df_team[df_team["Result"].astype(str).str.strip().str.lower() == "not approved"])

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
                continue

            tot_pct = f"{int(round((tot_approved / tot_target) * 100))}%" if tot_target > 0 else "0%"
            total_row = {"No": "", "Team": "", "Target Site": tot_target, "Approved": tot_approved, "Not Approved": tot_not_approved, "%": tot_pct, "Remain": tot_remain, "Remark": ""}

            df_summary = pd.DataFrame([total_row] + rows)
            df_summary.columns = pd.MultiIndex.from_tuples([("", "No"), ("", "Team"), ("", "Target Site"), ("Result", "Approved"), ("Result", "Not Approved"), ("", "%"), ("", "Remain"), ("", "Remark")])

            styled_summary = style_task_summary(df_summary, task_title)
            img_summary_path = f"summary_{task_code}.png"
            dfi.export(styled_summary.hide(axis="index"), img_summary_path, max_rows=-1)

            client.send_file(MAIN_GROUP_ID, img_summary_path, caption=f"របាយការណ៍សង្ខេប {task_title} ({shift_title})")

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
