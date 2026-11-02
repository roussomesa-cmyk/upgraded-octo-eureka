import io
import os
import re
from datetime import date, datetime, timedelta, timezone
from functools import lru_cache

import dataframe_image as dfi
import pandas as pd
import requests
from telethon.sessions import StringSession
from telethon.sync import TelegramClient

SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID") or "1PmMSqfeBWhYJe5dMv3PrLOFKc2YmLYP8BdCvf9FyZX4"
MAIN_GROUP_ID = int(os.environ.get("NOTIFY_GROUP_ID") or "-1001853372580")

VALID_TEAMS = [f"CHA-T0{i}" for i in range(1, 8)]
NA_VALUES = ("", "nan", "none", "#n/a", "n/a")

# ============================================================
# កំណត់ការផ្ញើអូតូនូវរបាយការណ៍ Pass/Not Pass (កូតប្រចាំសប្តាហ៍)
# ============================================================
# ៤សប្តាហ៍ គណនាដោយស្វ័យប្រវត្តិរាល់ខែ (មិនចាំបាច់កែដៃ)
# ច្បាប់៖ សប្តាហ៍ទី 1 ចាប់ផ្តើមថ្ងៃអាទិត្យដំបូងនៃខែ រួចបន្តជា 4 សប្តាហ៍ (អាទិត្យ→សៅរ៍)
# ឧ. តុលា 2026 → 4 តុលា ដល់ 31 តុលា | វិច្ឆិកា 2026 → 1 វិច្ឆិកា ដល់ 28 វិច្ឆិកា
CAMBODIA_TZ = timezone(timedelta(hours=7))


def first_sunday(year, month):
    d = date(year, month, 1)
    return d + timedelta(days=(6 - d.weekday()) % 7)


def build_weeks(start):
    return [
        {"label": f"សប្តាហ៍ទី {i + 1}",
         "start": start + timedelta(days=7 * i),
         "end": start + timedelta(days=7 * i + 6)}
        for i in range(4)
    ]


def detect_period_start(today):
    """ជ្រើសរយៈពេល 4 សប្តាហ៍ចុងក្រោយបំផុតដែលបានចាប់ផ្តើមរួចហើយ"""
    override = (os.environ.get("PERIOD_START") or "").strip()  # ឧ. 2026-10-04 (សម្រាប់ផ្ញើសរុបខែមុន)
    if override:
        return date.fromisoformat(override)
    prev_y, prev_m = (today.year - 1, 12) if today.month == 1 else (today.year, today.month - 1)
    candidates = [first_sunday(prev_y, prev_m), first_sunday(today.year, today.month)]
    started = [c for c in candidates if c <= today]
    return max(started) if started else candidates[0]


WEEKS = build_weeks(detect_period_start(datetime.now(CAMBODIA_TZ).date()))
print(f"ℹ️ Period: {WEEKS[0]['start']} → {WEEKS[-1]['end']}")

# Task ដែលចែកតិចជាង 4 សប្តាហ៍ (Task ផ្សេងទៀតចែក 4 សប្តាហ៍ធម្មតា)
TASK_WEEK_COUNT = {"A2": 3}

AUTO_DAILY = True        # ផ្ញើរបាយការណ៍ប្រចាំថ្ងៃទៅធីមរាល់ល្ងាច (ក្នុងអំឡុង 4 សប្តាហ៍)
MAX_SITES_LISTED = 8

# ប្រើសម្រាប់រត់ដោយដៃ/ផ្ញើបន្ថែម (ឧ. នៅក្នុង GitHub Actions ដាក់ env FORCE_REPORT)
#   FORCE_REPORT=final | weekly:2 | daily | daily:2026-10-20
#   បើចង់ផ្ញើរបាយការណ៍ខែមុន ដាក់ PERIOD_START=YYYY-MM-DD (ថ្ងៃអាទិត្យដំបូងនៃខែនោះ)
FORCE_REPORT = (os.environ.get("FORCE_REPORT") or "").strip().lower()

GREEN, RED = "background-color:#C6EFCE;font-weight:bold;", "background-color:#FFC7CE;font-weight:bold;"
KH_FONT = "Khmer OS Battambang, sans-serif"

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


@lru_cache(maxsize=None)
def cached_task_title(task_code):
    return get_task_title_from_sheet(task_code)[:40]


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


# ============================================================
# ផ្នែកកូតប្រចាំសប្តាហ៍ + Pass / Not Pass
# ============================================================
def parse_date(val, lo, hi):
    s = str(val).strip()
    if s.lower() in NA_VALUES:
        return None
    cands = []
    m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})", s)
    if m:
        y, a, b = map(int, m.groups())
        cands = [(y, a, b), (y, b, a)]
    else:
        m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})", s)
        if m:
            d, mo, y = map(int, m.groups())
            cands = [(y, mo, d), (y, d, mo)]
        elif re.match(r"^\d{5}(\.\d+)?$", s):
            dt = date(1899, 12, 30) + timedelta(days=int(float(s)))
            return dt if lo <= dt <= hi else None
    for y, mo, d in cands:
        try:
            dt = date(y, mo, d)
        except ValueError:
            continue
        if lo <= dt <= hi:
            return dt
    return None


def load_log_dates(lo, hi):
    df = fetch_csv("Approved_Log")
    out = {}
    if df is None or df.empty:
        return out
    date_col = "Date text" if "Date text" in df.columns else ("Detected at" if "Detected at" in df.columns else None)
    need = {"Team", "Task code", "Site"}
    if date_col is None or not need <= set(df.columns):
        return out
    for _, r in df.iterrows():
        key = (str(r["Task code"]).strip(), str(r["Site"]).strip(), str(r["Team"]).strip())
        out.setdefault(key, []).append(parse_date(r[date_col], lo, hi))
    return out


def collect_tasks(task_codes, lo, hi, log_dates=None):
    frames = []
    for code in task_codes:
        df = fetch_csv(code)
        if df is None or df.empty or "Team" not in df.columns or "Site name" not in df.columns:
            continue
        d = df.copy()
        d["Team"] = d["Team"].astype(str).str.strip()
        d = d[d["Team"].isin(VALID_TEAMS)]
        d = d[~d["Site name"].astype(str).str.strip().str.lower().isin(NA_VALUES)]
        if d.empty:
            continue
        if "Result" in d.columns:
            approved = d["Result"].astype(str).str.strip().str.lower() == "approved"
        else:
            approved = pd.Series(False, index=d.index)
        if "Last date record" in d.columns:
            dates = d["Last date record"].map(lambda v: parse_date(v, lo, hi))
        else:
            dates = pd.Series([None] * len(d), index=d.index)
        sites = d["Site name"].astype(str).str.strip().values
        final_dates = []
        for site, team, appr, d0 in zip(sites, d["Team"].values, approved.values, dates.values):
            dt = d0
            if appr and log_dates is not None:
                lst = log_dates.get((code, site, team))
                if lst:
                    log_dt = lst.pop(0)
                    if dt is None:
                        dt = log_dt
            final_dates.append(dt)
        frames.append(pd.DataFrame({
            "Task": code,
            "Team": d["Team"].values,
            "Site": sites,
            "Approved": approved.values,
            "Date": final_dates,
        }))
    if not frames:
        return pd.DataFrame(columns=["Task", "Team", "Site", "Approved", "Date"])
    return pd.concat(frames, ignore_index=True)


def split_quota(total, n=None):
    """ចែក total ជា n សប្តាហ៍ដំបូង ហើយសប្តាហ៍ដែលនៅសល់ = 0"""
    full = len(WEEKS)
    n = min(n or full, full)
    base, rem = divmod(total, n)
    return [base + (1 if i < rem else 0) for i in range(n)] + [0] * (full - n)


def team_quota(df, team):
    """កូតប្រចាំសប្តាហ៍របស់ធីម = ផលបូកកូតរបស់ Task នីមួយៗ (A2 ចែកតែ 3 សប្តាហ៍)"""
    quotas = [0] * len(WEEKS)
    d = df[df["Team"] == team]
    for task, g in d.groupby("Task"):
        q = split_quota(len(g), TASK_WEEK_COUNT.get(str(task).strip().upper()))
        quotas = [a + b for a, b in zip(quotas, q)]
    return quotas


def week_index(day):
    for i, w in enumerate(WEEKS):
        if w["start"] <= day <= w["end"]:
            return i
    return -1


def week_done(df, team, wi):
    w = WEEKS[wi]
    m = (df["Team"] == team) & df["Approved"] & df["Date"].map(lambda x: x is not None and w["start"] <= x <= w["end"])
    return int(m.sum())


def team_totals(df, team):
    d = df[df["Team"] == team]
    return len(d), int(d["Approved"].sum())


def style_pass_report(df, title, flags):
    st = df.style.set_caption(title).set_table_styles([
        {"selector": "caption", "props": [("caption-side", "top"), ("font-size", "20px"), ("font-weight", "bold"),
                                          ("text-align", "center"), ("background-color", "#2EA44E"), ("color", "white"),
                                          ("padding", "10px"), ("border", "1px solid black"), ("font-family", KH_FONT)]},
        {"selector": "th", "props": [("background-color", "#2EA44E"), ("color", "white"), ("font-weight", "bold"),
                                     ("text-align", "center"), ("border", "1px solid black"), ("padding", "6px"), ("font-family", KH_FONT)]},
        {"selector": "td", "props": [("text-align", "center"), ("border", "1px solid black"), ("padding", "5px"), ("font-family", KH_FONT)]},
    ])

    def row_style(row):
        out = []
        for col in df.columns:
            f = flags.get(col)
            if f is None or f[row.name] is None:
                out.append("font-weight:bold;background-color:#F2F2F2;" if row.name == len(df) - 1 else "")
            else:
                out.append(GREEN if f[row.name] else RED)
        return out

    return st.apply(row_style, axis=1)


def build_daily_reports(df, day):
    wi = week_index(day)
    reports = {}
    for team in VALID_TEAMS:
        rows = df[(df["Team"] == team) & df["Approved"] & (df["Date"] == day)]
        table_rows = []
        if rows.empty:
            table_rows.append({
                "ប្រភេទការងារ (Task Code)": "គ្មានការងារ Approved ថ្ងៃនេះទេ",
                "ចំនួន": 0,
                "បញ្ជី Site": "-",
            })
        else:
            for task, g in rows.groupby("Task"):
                title = cached_task_title(task)
                sites = list(g["Site"])
                shown = ", ".join(sites[:MAX_SITES_LISTED])
                if len(sites) > MAX_SITES_LISTED:
                    shown += f" ... +{len(sites) - MAX_SITES_LISTED}"
                task_label = f"{task}" + (f"\n({title})" if title else "")
                table_rows.append({
                    "ប្រភេទការងារ (Task Code)": task_label,
                    "ចំនួន": len(g),
                    "បញ្ជី Site": shown,
                })
        df_daily = pd.DataFrame(table_rows)
        title_text = f"របាយការណ៍ប្រចាំថ្ងៃ {team} ({day.isoformat()})\nApproved ថ្ងៃនេះ៖ {len(rows)} ការងារ"
        if wi >= 0:
            quota = team_quota(df, team)[wi]
            done_cnt = week_done(df, team, wi)
            title_text += f" | {WEEKS[wi]['label']}៖ Approved {done_cnt}/{quota}"

        styler = df_daily.style.set_caption(title_text).set_table_styles([
            {"selector": "caption", "props": [("caption-side", "top"), ("font-size", "16px"), ("font-weight", "bold"),
                                              ("text-align", "center"), ("background-color", "#1b5e20"), ("color", "white"),
                                              ("padding", "8px"), ("border", "1px solid black"), ("font-family", KH_FONT)]},
            {"selector": "th", "props": [("background-color", "#2e7d32"), ("color", "white"), ("font-weight", "bold"),
                                         ("text-align", "center"), ("border", "1px solid black"), ("padding", "6px"), ("font-family", KH_FONT)]},
            {"selector": "td", "props": [("text-align", "center"), ("border", "1px solid black"), ("padding", "6px"), ("font-family", KH_FONT)]},
        ])
        caption_text = f"របាយការណ៍ប្រចាំថ្ងៃ {team} - {day.isoformat()}"
        reports[team] = (styler, caption_text)
    return reports


def build_weekly(df, wi):
    w = WEEKS[wi]
    rows, flags_ok = [], []
    t_done = t_quota = 0
    for i, team in enumerate(VALID_TEAMS, start=1):
        quota = team_quota(df, team)[wi]
        done = week_done(df, team, wi)
        ok = done >= quota
        t_done += done
        t_quota += quota
        pct = f"{int(round(done / quota * 100))}%" if quota > 0 else "0%"
        rows.append({"No": i, "Team": team, "ត្រូវធ្វើ": quota, "Approved": done, "%": pct,
                     "Result": "Pass" if ok else "Not Pass"})
        flags_ok.append(ok)
    n_pass = sum(flags_ok)
    pct = f"{int(round(t_done / t_quota * 100))}%" if t_quota else "0%"
    rows.append({"No": "", "Team": "TOTAL", "ត្រូវធ្វើ": t_quota, "Approved": t_done, "%": pct,
                 "Result": f"Pass {n_pass}/{len(VALID_TEAMS)}"})
    out = pd.DataFrame(rows)
    flags = {"Result": flags_ok + [None]}
    title = f"សរុប{w['label']} ({w['start']} → {w['end']})"
    caption = f"{title}\n" + "\n".join(
        f"{r['Team']}: {r['Approved']}/{r['ត្រូវធ្វើ']}  {'Pass' if r['Result'] == 'Pass' else 'Not Pass'}"
        for r in rows[:-1]) + f"\n\nPass {n_pass} ធីម | Not Pass {len(VALID_TEAMS) - n_pass} ធីម"
    team_texts = {}
    for r in rows[:-1]:
        ok = r["Result"] == "Pass"
        miss = max(r["ត្រូវធ្វើ"] - r["Approved"], 0)
        status_text = "Pass" if ok else f"Not Pass (ខ្វះ {miss})"
        team_texts[r["Team"]] = (
            f"{r['Team']} - {title}\n\n"
            f"• ត្រូវធ្វើ៖ {r['ត្រូវធ្វើ']}\n"
            f"• Approved៖ {r['Approved']} ({r['%']})\n"
            f"• ស្ថានភាព៖ {status_text}"
        )
    return out, flags, title, caption, team_texts


def build_final(df):
    rows, flag_cols, res_flags, team_texts = [], {f"ស{i + 1}": [] for i in range(len(WEEKS))}, [], {}
    for team in VALID_TEAMS:
        total, approved = team_totals(df, team)
        quotas = team_quota(df, team)
        row = {"Team": team}
        wk_lines = []
        for i in range(len(WEEKS)):
            done = week_done(df, team, i)
            row[f"ស{i + 1}"] = f"{done}/{quotas[i]}"
            is_ok = done >= quotas[i]
            flag_cols[f"ស{i + 1}"].append(is_ok)
            wk_lines.append(f"• {WEEKS[i]['label']}៖ {done}/{quotas[i]} {'(Pass)' if is_ok else '(Not Pass)'}")

        ok = total > 0 and approved >= total
        miss_total = max(total - approved, 0)
        status_final = "Pass" if ok else f"Not Pass (ខ្វះ {miss_total} ការងារ)"

        team_texts[team] = (
            f"{team} - សរុបរួម {len(WEEKS)} សប្តាហ៍\n\n"
            + "\n".join(wk_lines) + "\n\n"
            f"• Approved សរុប៖ {approved}/{total}\n"
            f"• លទ្ធផល៖ {status_final}"
        )
        row["Approved សរុប"] = f"{approved}/{total}"
        row["Result"] = "Pass" if ok else "Not Pass"
        res_flags.append(ok)
        rows.append(row)

    out = pd.DataFrame(rows)
    flags = dict(flag_cols)
    flags["Result"] = res_flags
    title = f"សរុបរួម {len(WEEKS)} សប្តាហ៍ ({WEEKS[0]['start']} → {WEEKS[-1]['end']})"
    caption = f"{title}\n" + "\n".join(
        f"{r['Team']}: {r['Approved សរុប']}  {'Pass' if r['Result'] == 'Pass' else 'Not Pass'}" for r in rows
    ) + f"\n\nPass {sum(res_flags)} ធីម | Not Pass {len(VALID_TEAMS) - sum(res_flags)} ធីម" \
        "\n(Approved សរុប រាប់ទាំងការងារដែលគ្មានកាលបរិច្ឆេទ)"
    return out, flags, title, caption, team_texts


def pick_pass_jobs(today):
    """សម្រេចថាថ្ងៃនេះត្រូវផ្ញើរបាយការណ៍អ្វីខ្លះ (អូតូ ឬតាម FORCE_REPORT)"""
    jobs = []
    if FORCE_REPORT:
        kind, _, arg = FORCE_REPORT.partition(":")
        if kind == "daily":
            jobs.append(("daily", date.fromisoformat(arg) if arg else today))
        elif kind == "weekly":
            wi = (int(arg) - 1) if arg else week_index(today)
            if 0 <= wi < len(WEEKS):
                jobs.append(("weekly", wi))
            else:
                print("⚠️ FORCE_REPORT=weekly ត្រូវការលេខសប្តាហ៍ 1-4 (ឧ. weekly:2)")
        elif kind == "final":
            jobs.append(("final", None))
        else:
            print(f"⚠️ FORCE_REPORT មិនស្គាល់៖ {FORCE_REPORT}")
        return jobs

    wi = week_index(today)
    if wi >= 0:
        if AUTO_DAILY:
            jobs.append(("daily", today))
        if today == WEEKS[wi]["end"]:
            jobs.append(("weekly", wi))
    if today == WEEKS[-1]["end"]:
        jobs.append(("final", None))
    return jobs


def run_pass_reports(client, today, task_chat_ids, team_chat_ids):
    jobs = pick_pass_jobs(today)
    if not jobs:
        print("ℹ️ Pass report: ថ្ងៃនេះគ្មានរបាយការណ៍ត្រូវផ្ញើទេ។")
        return

    lo, hi = WEEKS[0]["start"], min(WEEKS[-1]["end"], today)
    log_dates = load_log_dates(lo, hi)
    df = collect_tasks(list(dict.fromkeys(task_chat_ids)), lo, hi, log_dates)

    for kind, arg in jobs:
        print(f"ℹ️ Pass report: {kind} {arg if arg is not None else ''}")
        if kind == "daily":
            for team, (styler, caption) in build_daily_reports(df, arg).items():
                chat = team_chat_ids.get(team)
                if not chat:
                    print(f"⚠️ ខ្វះ Chat ID របស់ {team} រំលង")
                    continue
                path = f"daily_{team}.png"
                dfi.export(styler.hide(axis="index"), path, max_rows=-1)
                client.send_file(chat, path, caption=caption[:1000])
            continue

        if kind == "weekly":
            table, flags, title, caption, team_texts = build_weekly(df, arg)
            path = f"weekly_{arg + 1}.png"
        else:
            table, flags, title, caption, team_texts = build_final(df)
            path = "final_report.png"

        for team in VALID_TEAMS:
            chat = team_chat_ids.get(team)
            if chat:
                client.send_message(chat, team_texts[team][:4000])
            else:
                print(f"⚠️ ខ្វះ Chat ID របស់ {team} រំលង")

        styled = style_pass_report(table, title, flags)
        dfi.export(styled.hide(axis="index"), path, max_rows=-1)
        client.send_file(MAIN_GROUP_ID, path, caption=caption[:1000])


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
    team_chat_ids = {}

    if df_mapping is not None and "Sheet" in df_mapping.columns:
        df_clean = df_mapping.dropna(subset=["Sheet", "ChatID"])
        for _, row in df_clean.iterrows():
            code = str(row["Sheet"]).strip()
            try:
                task_chat_ids[code] = int(float(str(row["ChatID"]).strip()))
            except ValueError:
                continue

    if df_mapping is not None and {"Column A", "Column B"} <= set(df_mapping.columns):
        for _, row in df_mapping.dropna(subset=["Column A", "Column B"]).iterrows():
            m = re.search(r"(\d+)", str(row["Column A"]))
            if not m:
                continue
            team = f"CHA-T{int(m.group(1)):02d}"
            try:
                team_chat_ids[team] = int(float(str(row["Column B"]).strip()))
            except ValueError:
                continue

    if not task_chat_ids:
        print("⚠️ WARNING: task_chat_ids is empty — check the 'Team chat IDs' sheet. No reports will be sent.")
        return

    print(f"ℹ️ Loaded {len(task_chat_ids)} task->chat mappings: {task_chat_ids}")
    print(f"ℹ️ Loaded {len(team_chat_ids)} team->chat mappings: {team_chat_ids}")

    cambodia_tz = timezone(timedelta(hours=7))
    now = datetime.now(cambodia_tz)
    is_morning = now.hour < 12
    shift_title = "Morning Shift" if is_morning else "Evening Shift"
    print(f"ℹ️ Shift: {shift_title} — Summary/Overall reports will be sent to MAIN_GROUP_ID only in the Evening Shift.")

    overall_stats = {team: {"Target": 0, "Approved": 0, "NotApproved": 0} for team in VALID_TEAMS}

    with TelegramClient(StringSession(session_str), api_id, api_hash) as client:
        # រត់ដោយដៃ៖ ផ្ញើតែរបាយការណ៍ Pass/Not Pass ដែលបានកំណត់ រួចបញ្ចប់
        if FORCE_REPORT:
            run_pass_reports(client, now.date(), task_chat_ids, team_chat_ids)
            return

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

        # របាយការណ៍ Pass/Not Pass (កូតប្រចាំសប្តាហ៍) — ផ្ញើអូតូតែពេលល្ងាច
        # បើមានបញ្ហា កុំឱ្យប៉ះពាល់របាយការណ៍ខាងលើដែលបានផ្ញើរួច
        if not is_morning:
            try:
                run_pass_reports(client, now.date(), task_chat_ids, team_chat_ids)
            except Exception as e:
                print(f"⚠️ Pass report failed: {e}")


if __name__ == "__main__":
    main()
