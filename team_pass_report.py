"""
team_pass_report.py - Report Power CHA2026 (រត់ដោយដៃ ដាច់ដោយឡែកពីគ្នា)

File នេះមិនប៉ះពាល់កូដ Morning/Evening Shift ចាស់ទេ។ រត់ម្តងមួយ Mode ដោយដៃ៖

  python team_pass_report.py daily  [--date 2026-10-17]   # របាយការណ៍ប្រចាំថ្ងៃ ជារូបភាពតារាង ទៅក្រុមធីមនីមួយៗ
  python team_pass_report.py weekly [--week 2]            # សរុបសប្តាហ៍ ៧ធីម Pass / Not Pass
  python team_pass_report.py final                        # សរុបរួមទាំង ៤សប្តាហ៍

ជម្រើសរួម៖
  --dry-run          បង្ហាញលទ្ធផលក្នុង Terminal តែមិនផ្ញើ Telegram (មិនត្រូវការ TELEGRAM_* ទេ)
  --test-chat ID      ផ្ញើសារទាំងអស់ទៅ Chat ID នេះតែមួយ (សម្រាប់សាកល្បង)
  --no-teams          មិនផ្ញើទៅក្រុមធីម
  --no-main           មិនផ្ញើទៅ MAIN_GROUP_ID (weekly/final)
"""
import argparse
import io
import os
import re
import sys
from datetime import date, datetime, timedelta, timezone

import pandas as pd
import requests

SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID") or "1PmMSqfeBWhYJe5dMv3PrLOFKc2YmLYP8BdCvf9FyZX4"
MAIN_GROUP_ID = int(os.environ.get("NOTIFY_GROUP_ID") or "-1001853372580")
VALID_TEAMS = [f"CHA-T0{i}" for i in range(1, 8)]
CAMBODIA_TZ = timezone(timedelta(hours=7))

# ៤សប្តាហ៍ (កែបាន)
WEEKS = [
    {"label": "សប្តាហ៍ទី 1", "start": date(2026, 10, 1), "end": date(2026, 10, 7)},
    {"label": "សប្តាហ៍ទី 2", "start": date(2026, 10, 8), "end": date(2026, 10, 14)},
    {"label": "សប្តាហ៍ទី 3", "start": date(2026, 10, 15), "end": date(2026, 10, 21)},
    {"label": "សប្តាហ៍ទី 4", "start": date(2026, 10, 22), "end": date(2026, 10, 31)},
]

GREEN, RED = "background-color:#C6EFCE;font-weight:bold;", "background-color:#FFC7CE;font-weight:bold;"
MAX_SITES_LISTED = 8
NA_VALUES = ("", "nan", "none", "#n/a", "n/a")


# ---------------------------------------------------------------- ទាញទិន្នន័យ
def fetch_csv(sheet_name_or_gid):
    url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet={sheet_name_or_gid}"
    try:
        res = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=20)
    except requests.RequestException as e:
        print(f"⚠️ Request failed for sheet '{sheet_name_or_gid}': {e}")
        return None
    if res.status_code != 200:
        print(f"⚠️ Failed to fetch sheet '{sheet_name_or_gid}': HTTP {res.status_code}")
        return None

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
        c = col.lower().strip()
        if "site name" in c:
            rename_dict[col] = "Site name"
        elif c in ("no", "no."):
            rename_dict[col] = "No."
        elif c == "team":
            rename_dict[col] = "Team"
        elif "group task" in c:
            rename_dict[col] = "Group task"
        elif "result" in c:
            rename_dict[col] = "Result"
        elif "serial" in c:
            rename_dict[col] = "Serial"
    if rename_dict:
        df = df.rename(columns=rename_dict)
    if df.columns.duplicated().any():
        df = df.loc[:, ~df.columns.duplicated()]
    return df


def get_task_title(task_code, cache={}):
    if task_code in cache:
        return cache[task_code]
    url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet={task_code}"
    title = ""
    try:
        res = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=20)
        if res.status_code == 200 and res.text.splitlines():
            first = res.text.splitlines()[0].replace('"', "").strip().split(",")[0].strip()
            if first and not first.isdigit():
                title = first[:40]
    except requests.RequestException:
        pass
    cache[task_code] = title
    return title


def load_mappings():
    """Return (task_chat_ids, team_chat_ids). team_chat_ids ពី Column A/B ក្នុង Tab 'Team chat IDs'."""
    df = fetch_csv("Team%20chat%20IDs")
    task_chat, team_chat = {}, {}
    if df is None:
        return task_chat, team_chat
    if {"Sheet", "ChatID"} <= set(df.columns):
        for _, r in df.dropna(subset=["Sheet", "ChatID"]).iterrows():
            try:
                task_chat[str(r["Sheet"]).strip()] = int(float(str(r["ChatID"]).strip()))
            except ValueError:
                continue
    if {"Column A", "Column B"} <= set(df.columns):
        for _, r in df.dropna(subset=["Column A", "Column B"]).iterrows():
            m = re.search(r"(\d+)", str(r["Column A"]))
            if not m:
                continue
            team = f"CHA-T{int(m.group(1)):02d}"
            try:
                team_chat[team] = int(float(str(r["Column B"]).strip()))
            except ValueError:
                continue
    return task_chat, team_chat


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
        print("ℹ️ មិនមាន Tab Approved_Log ឬទទេ")
        return out
    date_col = "Date text" if "Date text" in df.columns else ("Detected at" if "Detected at" in df.columns else None)
    need = {"Team", "Task code", "Site"}
    if date_col is None or not need <= set(df.columns):
        print("⚠️ Approved_Log ខ្វះ Column ចាំបាច់ — រំលង")
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


# ---------------------------------------------------------------- គណនា
def split_quota(total):
    n = len(WEEKS)
    base, rem = divmod(total, n)
    return [base + (1 if i < rem else 0) for i in range(n)]


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


# ---------------------------------------------------------------- រចនា Table
def style_report(df, title, flags):
    st = df.style.set_caption(title).set_table_styles([
        {"selector": "caption", "props": [("caption-side", "top"), ("font-size", "20px"), ("font-weight", "bold"),
                                          ("text-align", "center"), ("background-color", "#2EA44E"), ("color", "white"),
                                          ("padding", "10px"), ("border", "1px solid black"), ("font-family", "Khmer OS Battambang, sans-serif")]},
        {"selector": "th", "props": [("background-color", "#2EA44E"), ("color", "white"), ("font-weight", "bold"),
                                     ("text-align", "center"), ("border", "1px solid black"), ("padding", "6px"), ("font-family", "Khmer OS Battambang, sans-serif")]},
        {"selector": "td", "props": [("text-align", "center"), ("border", "1px solid black"), ("padding", "5px"), ("font-family", "Khmer OS Battambang, sans-serif")]},
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


def export_image(styler, path):
    import dataframe_image as dfi
    dfi.export(styler.hide(axis="index"), path, max_rows=-1)


# ---------------------------------------------------------------- Mode: daily (ជារូបភាពតារាង)
def build_daily_reports(df, day):
    """បង្កើតតារាងជារូបភាពសម្រាប់ Daily Report របស់ក្រុមនីមួយៗ"""
    wi = week_index(day)
    reports = {}
    for team in VALID_TEAMS:
        rows = df[(df["Team"] == team) & df["Approved"] & (df["Date"] == day)]
        
        table_rows = []
        if rows.empty:
            table_rows.append({
                "ប្រភេទការងារ (Task Code)": "គ្មានការងារ Approved ថ្ងៃនេះទេ",
                "ចំនួន": 0,
                "បញ្ជី Site": "-"
            })
        else:
            for task, g in rows.groupby("Task"):
                title = get_task_title(task)
                sites = list(g["Site"])
                shown = ", ".join(sites[:MAX_SITES_LISTED])
                if len(sites) > MAX_SITES_LISTED:
                    shown += f" ... +{len(sites) - MAX_SITES_LISTED}"
                
                task_label = f"{task}" + (f"\n({title})" if title else "")
                table_rows.append({
                    "ប្រភេទការងារ (Task Code)": task_label,
                    "ចំនួន": len(g),
                    "បញ្ជី Site": shown
                })
        
        df_daily = pd.DataFrame(table_rows)
        
        # ក្បាលតារាង & Summary Progress
        title_text = f"📋 របាយការណ៍ប្រចាំថ្ងៃ {team} ({day.isoformat()})\n✅ Approved ថ្ងៃនេះ៖ {len(rows)} ការងារ"
        if wi >= 0:
            total, _ = team_totals(df, team)
            quota = split_quota(total)[wi]
            done_cnt = week_done(df, team, wi)
            title_text += f" | 📊 {WEEKS[wi]['label']}៖ Approved {done_cnt}/{quota}"

        # ឌីហ្សាញតារាង Daily
        styler = df_daily.style.set_caption(title_text).set_table_styles([
            {"selector": "caption", "props": [("caption-side", "top"), ("font-size", "16px"), ("font-weight", "bold"),
                                              ("text-align", "center"), ("background-color", "#1b5e20"), ("color", "white"),
                                              ("padding", "8px"), ("border", "1px solid black"), ("font-family", "Khmer OS Battambang, sans-serif")]},
            {"selector": "th", "props": [("background-color", "#2e7d32"), ("color", "white"), ("font-weight", "bold"),
                                         ("text-align", "center"), ("border", "1px solid black"), ("padding", "6px"), ("font-family", "Khmer OS Battambang, sans-serif")]},
            {"selector": "td", "props": [("text-align", "center"), ("border", "1px solid black"), ("padding", "6px"), ("font-family", "Khmer OS Battambang, sans-serif")]},
        ])
        
        caption_text = f"📋 របាយការណ៍ប្រចាំថ្ងៃ {team} - {day.isoformat()}"
        reports[team] = (styler, df_daily, caption_text)
        
    return reports


# ---------------------------------------------------------------- Mode: weekly
def build_weekly(df, wi):
    w = WEEKS[wi]
    rows, flags_ok = [], []
    t_done = t_quota = 0
    for i, team in enumerate(VALID_TEAMS, start=1):
        total, _ = team_totals(df, team)
        quota = split_quota(total)[wi]
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
    title = f"{w['label']} ({w['start']} → {w['end']})"
    caption = f"🏁 សរុប{title}\n" + "\n".join(
        f"{r['Team']}: {r['Approved']}/{r['ត្រូវធ្វើ']}  {'✅ Pass' if r['Result'] == 'Pass' else '❌ Not Pass'}"
        for r in rows[:-1]) + f"\n\nPass {n_pass} ធីម | Not Pass {len(VALID_TEAMS) - n_pass} ធីម"
    team_texts = {}
    for r in rows[:-1]:
        ok = r["Result"] == "Pass"
        miss = max(r["ត្រូវធ្វើ"] - r["Approved"], 0)
        team_texts[r["Team"]] = (
            f"🏁 {r['Team']} - សរុប{title}\n\nត្រូវធ្វើ៖ {r['ត្រូវធ្វើ']}\nApproved៖ {r['Approved']} ({r['%']})\n\n"
            + ("✅ Pass" if ok else f"❌ Not Pass (នៅខ្វះ {miss} ការងារ)"))
    return out, flags, title, caption, team_texts


# ---------------------------------------------------------------- Mode: final
def build_final(df):
    rows, flag_cols, res_flags, team_texts = [], {f"ស{i + 1}": [] for i in range(len(WEEKS))}, [], {}
    for team in VALID_TEAMS:
        total, approved = team_totals(df, team)
        quotas = split_quota(total)
        row = {"Team": team}
        wk_lines = []
        for i in range(len(WEEKS)):
            done = week_done(df, team, i)
            row[f"ស{i + 1}"] = f"{done}/{quotas[i]}"
            flag_cols[f"ស{i + 1}"].append(done >= quotas[i])
            wk_lines.append(f"{WEEKS[i]['label']}៖ {done}/{quotas[i]} " + ("✅" if done >= quotas[i] else "❌"))
        ok = total > 0 and approved >= total
        team_texts[team] = (f"🏆 {team} - សរុបរួម {len(WEEKS)} សប្តាហ៍\n\n" + "\n".join(wk_lines)
                            + f"\n\nApproved សរុប៖ {approved}/{total}\n" + ("✅ Pass" if ok else f"❌ Not Pass (នៅខ្វះ {max(total - approved, 0)} ការងារ)"))
        row["Approved សរុប"] = f"{approved}/{total}"
        row["Result"] = "Pass" if ok else "Not Pass"
        res_flags.append(ok)
        rows.append(row)
    out = pd.DataFrame(rows)
    flags = {k: v for k, v in flag_cols.items()}
    flags["Result"] = res_flags
    title = f"សរុបរួម {len(WEEKS)} សប្តាហ៍ ({WEEKS[0]['start']} → {WEEKS[-1]['end']})"
    caption = f"🏆 {title}\n" + "\n".join(
        f"{r['Team']}: {r['Approved សរុប']}  {'✅ Pass' if r['Result'] == 'Pass' else '❌ Not Pass'}" for r in rows
    ) + f"\n\nPass {sum(res_flags)} ធីម | Not Pass {len(VALID_TEAMS) - sum(res_flags)} ធីម" \
        "\n(Approved សរុប រាប់ទាំងការងារដែលគ្មានកាលបរិច្ឆេទ)"
    return out, flags, title, caption, team_texts


# ---------------------------------------------------------------- Telegram
class Sender:
    def __init__(self, dry_run, test_chat):
        self.dry, self.test_chat, self.client = dry_run, test_chat, None

    def __enter__(self):
        if self.dry:
            return self
        from telethon.sessions import StringSession
        from telethon.sync import TelegramClient
        api_id, api_hash, session = (os.environ.get("TELEGRAM_API_ID"), os.environ.get("TELEGRAM_API_HASH"),
                                     os.environ.get("TELEGRAM_SESSION"))
        if not (api_id and api_hash and session):
            raise ValueError("ខ្វះ TELEGRAM_API_ID / TELEGRAM_API_HASH / TELEGRAM_SESSION")
        self.client = TelegramClient(StringSession(session), int(api_id), api_hash).__enter__()
        return self

    def __exit__(self, *exc):
        if self.client:
            self.client.__exit__(*exc)

    def text(self, chat_id, msg):
        target = self.test_chat or chat_id
        if self.dry:
            print(f"\n--- [dry-run] ទៅ {target} ---\n{msg}")
            return
        self.client.send_message(target, msg[:4000])

    def image(self, chat_id, styler, path, caption, df):
        target = self.test_chat or chat_id
        if self.dry:
            print(f"\n--- [dry-run] រូបភាព ទៅ {target} ---\n{df.to_string(index=False)}\n\n{caption}")
            return
        export_image(styler, path)
        self.client.send_file(target, path, caption=caption[:1000])


# ---------------------------------------------------------------- Main
def main():
    ap = argparse.ArgumentParser(description="Team Pass report (រត់ដោយដៃ)")
    ap.add_argument("mode", choices=["daily", "weekly", "final"])
    ap.add_argument("--date", help="daily៖ YYYY-MM-DD (លំនាំដើម = ថ្ងៃនេះ)")
    ap.add_argument("--week", type=int, help="weekly៖ 1-4 (លំនាំដើម = សប្តាហ៍ដែលមានថ្ងៃនេះ)")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--test-chat", type=int)
    ap.add_argument("--no-teams", action="store_true")
    ap.add_argument("--no-main", action="store_true")
    args = ap.parse_args()

    today = datetime.now(CAMBODIA_TZ).date()
    task_chat, team_chat = load_mappings()
    if not task_chat:
        sys.exit("⚠️ រកមិនឃើញ Sheet→ChatID ក្នុង Tab 'Team chat IDs'")
    print(f"ℹ️ Chat ID ធីម៖ {team_chat}")

    lo, hi = WEEKS[0]["start"], min(WEEKS[-1]["end"], today)
    log_dates = load_log_dates(lo, hi)
    df = collect_tasks(list(dict.fromkeys(task_chat)), lo, hi, log_dates)
    print(f"ℹ️ ការងារសរុប {len(df)} ជួរ | Approved {int(df['Approved'].sum())} | "
          f"Approved មានកាលបរិច្ឆេទ {int((df['Approved'] & df['Date'].notna()).sum())}")

    with Sender(args.dry_run, args.test_chat) as s:
        if args.mode == "daily":
            day = date.fromisoformat(args.date) if args.date else today
            daily_reports = build_daily_reports(df, day)
            for team, (styler, df_daily, caption) in daily_reports.items():
                img_path = f"daily_{team}.png"
                if team in team_chat:
                    s.image(team_chat[team], styler, img_path, caption, df_daily)
                else:
                    print(f"⚠️ ខ្វះ Chat ID របស់ {team} រំលង")
            return

        if args.mode == "weekly":
            wi = (args.week - 1) if args.week else week_index(today)
            if not 0 <= wi < len(WEEKS):
                sys.exit("ថ្ងៃនេះមិនស្ថិតក្នុងសប្តាហ៍ណាទេ សូមប្រើ --week 1..4")
            table, flags, title, caption, team_texts = build_weekly(df, wi)
            path = f"weekly_{wi + 1}.png"
        else:
            table, flags, title, caption, team_texts = build_final(df)
            path = "final_report.png"

        # ក្រុមធីមនីមួយៗ ទទួលលទ្ធផលរបស់ធីមខ្លួនឯង
        if not args.no_teams:
            for team in VALID_TEAMS:
                if team in team_chat:
                    s.text(team_chat[team], team_texts[team])
                else:
                    print(f"⚠️ ខ្វះ Chat ID របស់ {team} រំលង")
        # មែនគ្រុប (MAIN_GROUP_ID) ទទួលតារាងសរុបទាំង ៧ធីម
        if not args.no_main:
            s.image(MAIN_GROUP_ID, style_report(table, title, flags), path, caption, table)


if __name__ == "__main__":
    main()
