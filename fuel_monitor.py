import asyncio
import calendar
import io
import json
import os
import re
from datetime import datetime, timezone, timedelta

import gspread
import pandas as pd
import requests
from google.oauth2.service_account import Credentials
from gspread.utils import rowcol_to_a1, ValueInputOption
from telethon import TelegramClient
from telethon.sessions import StringSession

# ============================ CONFIG ============================
SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID") or "1PmMSqfeBWhYJe5dMv3PrLOFKc2YmLYP8BdCvf9FyZX4"
STATION_GROUP_ID = int(os.environ.get("STATION_GROUP_ID") or "0")

FUEL_ALERT_GROUP_IDS = [
    int(x) for x in re.split(r"[,\s]+", os.environ.get("FUEL_ALERT_GROUP_ID") or "") if x
]
FUEL_ALERT_THRESHOLD = float(os.environ.get("FUEL_ALERT_THRESHOLD") or "15")

STATION_LIST_SHEET = "List Site"
STATION_CODE_COLUMN = "Site name"
OUTPUT_SHEET_NAME = os.environ.get("OUTPUT_SHEET_NAME") or "Fuel Monitor Log"

COMMAND_PREFIX = "/mn"
RESPONSE_TIMEOUT_SEC = 20
DELAY_BETWEEN_CODES_SEC = 180

# Report
REPORT_SHEET_NAME = "Weekly Report"
SUMMARY_SHEET_NAME = "Summary Report"
TEAM_LIST_SHEET = "allteam"
TEAM_RANGE = "allteam!$A:$B"      # A = Site Name, B = Team
TEAM_LIST_COL = 2                  # allteam column B = Team
LOW_FUEL_THRESHOLD = 15           # Remark ក្នុង Weekly Report
SUMMARY_THRESHOLD = 20            # Summary: ក្រោម 20%
TELEGRAM_MAX_LEN = 3800           # Telegram អនុញ្ញាត 4096 តួអក្សរ ទុកគម្លាតបន្តិច
REPORT_HEADERS = ["ថ្ងៃខែ", "Site Name", "Team", "Fuel Level(%)",
                  "Battery Voltage of DG (V)", "Remark"]

OUTPUT_HEADERS = [
    "Timestamp", "Station Code (Requested)", "Station Code (Reply)", "Vendor",
    "Cooling Water Temperature (°C)", "Oil Pressure (kPa)", "Fuel Level(%)",
    "Battery Voltage of DG (V)", "Total Runtime of DG (hour)", "Addition Runtime of DG (min)",
    "Phase Voltage 1 (V)", "Phase Voltage 2 (V)", "Phase Voltage 3 (V)",
    "Phase Current 1 (A)", "Phase Current 2 (A)", "Phase Current 3 (A)",
    "Status",
]

FIELD_MAP = {
    "station code": "Station Code (Reply)",
    "vendor": "Vendor",
    "cooling water temperature": "Cooling Water Temperature (°C)",
    "oil pressure": "Oil Pressure (kPa)",
    "fuel level": "Fuel Level(%)",
    "battery voltage of dg": "Battery Voltage of DG (V)",
    "total runtime of dg": "Total Runtime of DG (hour)",
    "addition runtime of dg": "Addition Runtime of DG (min)",
    "phase voltage 1": "Phase Voltage 1 (V)",
    "phase voltage 2": "Phase Voltage 2 (V)",
    "phase voltage 3": "Phase Voltage 3 (V)",
    "phase curent 1": "Phase Current 1 (A)",
    "phase curent 2": "Phase Current 2 (A)",
    "phase curent 3": "Phase Current 3 (A)",
    "phase current 1": "Phase Current 1 (A)",
    "phase current 2": "Phase Current 2 (A)",
    "phase current 3": "Phase Current 3 (A)",
}

# សញ្ញា Emoji / Pictograph / ព្រួញ / Bullet ដែលត្រូវដកចេញពីសារ Telegram
SYMBOL_RE = re.compile("[\U0001F000-\U0001FAFF\u2600-\u27BF\u2B00-\u2BFF\u2190-\u21FF\u2022\uFE0F\u200D]+")


def strip_symbols(text):
    return SYMBOL_RE.sub("", str(text)).strip()


# ====================== LIST SITE / TEAM INFO ======================
def _fetch_list_site_df():
    url = (f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq"
           f"?tqx=out:csv&sheet={STATION_LIST_SHEET.replace(' ', '%20')}")
    try:
        res = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=15)
    except requests.RequestException as e:
        print(f"[WARN] Failed to fetch station list: {e}")
        return None
    if res.status_code != 200:
        print(f"[WARN] Failed to fetch station list: HTTP {res.status_code}")
        return None
    df = pd.read_csv(io.StringIO(res.text))
    df.columns = df.columns.astype(str).str.strip()
    return df


def fetch_station_codes():
    df = _fetch_list_site_df()
    if df is None:
        return []
    if STATION_CODE_COLUMN not in df.columns:
        print(f"[WARN] Column '{STATION_CODE_COLUMN}' not found. Columns found: {list(df.columns)}")
        return []
    codes = df[STATION_CODE_COLUMN].dropna().astype(str).str.strip()
    return [c for c in codes if c and c.lower() not in ("nan", "none")]


def fetch_team_info():
    """Return {site_name: {"team": ..., "leader": ...}} ពី List Site."""
    df = _fetch_list_site_df()
    if df is None:
        return {}
    leader_col = next((c for c in df.columns
                       if c.lower().replace(" ", "").startswith("teamlea")), None)
    team_col = "Stock code" if "Stock code" in df.columns else None

    info = {}
    for _, r in df.iterrows():
        site = str(r.get(STATION_CODE_COLUMN, "")).strip()
        if not site or site.lower() == "nan":
            continue
        leader = str(r[leader_col]).strip() if leader_col and pd.notna(r[leader_col]) else ""
        team = str(r[team_col]).strip() if team_col and pd.notna(r[team_col]) else ""
        info.setdefault(site, {"team": team, "leader": leader})
    return info


# ============================ PARSER ============================
def parse_bot_reply(text):
    parsed = {}
    for line in text.splitlines():
        if ":" not in line:
            continue
        label_raw, _, value_raw = line.partition(":")
        label_clean = re.sub(r"\(.*?\)", "", label_raw).strip().lower()
        value_clean = value_raw.strip()
        for key_pattern, column_name in FIELD_MAP.items():
            if label_clean == key_pattern:
                parsed[column_name] = value_clean
                break
    return parsed


# ============================ SHEETS ============================
def get_sheets_client():
    sa_key_raw = os.environ.get("GCP_SA_KEY")
    if not sa_key_raw:
        raise ValueError("GCP_SA_KEY is missing - required to write to Google Sheet.")
    sa_info = json.loads(sa_key_raw)
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_info(sa_info, scopes=scopes)
    return gspread.authorize(creds)


def get_or_create_output_worksheet(gc):
    sh = gc.open_by_key(SPREADSHEET_ID)
    try:
        ws = sh.worksheet(OUTPUT_SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=OUTPUT_SHEET_NAME, rows=1000, cols=len(OUTPUT_HEADERS))
        ws.append_row(OUTPUT_HEADERS)
        print(f"[INFO] Created new worksheet '{OUTPUT_SHEET_NAME}' with headers.")
    return ws


async def append_row_with_retry(ws, row, max_attempts=3):
    for attempt in range(1, max_attempts + 1):
        try:
            ws.append_row(row, value_input_option="USER_ENTERED")
            return True
        except Exception as e:
            print(f"[WARN] Sheet write attempt {attempt} failed: {e}")
            if attempt < max_attempts:
                wait_sec = 10 * attempt
                await asyncio.sleep(wait_sec)
            else:
                print(f"[ERROR] Sheet write failed after {max_attempts} attempts - skipping row.")
                return False


# ============================ FUEL ALERT & TELEGRAM REPORTS ============================
def _pct(v):
    return f"{v:g}%"  # 0.0 -> 0%, 11.0 -> 11%


def build_fuel_alert_message(code, fuel_val, reply_text, team_info):
    info = team_info.get(code, {})
    leader = info.get("leader", "")
    team = info.get("team", "")
    lines = [f"Fuel ទាប! Site: {code}",
             f"Fuel Level: {_pct(fuel_val)}"]
    if team:
        lines.append(f"Team: {team}")
    if leader:
        lines.append(f"Bong {leader}")
    msg = "\n".join(lines) + "\n\n" + strip_symbols(reply_text)
    return msg[:TELEGRAM_MAX_LEN]


async def send_fuel_alert(client, code, fuel_val, reply_text, team_info):
    msg = build_fuel_alert_message(code, fuel_val, reply_text, team_info)
    if not FUEL_ALERT_GROUP_IDS:
        return
    for gid in FUEL_ALERT_GROUP_IDS:
        try:
            await client.send_message(gid, msg, parse_mode=None, link_preview=False)
            print(f"[ALERT] Fuel alert sent for '{code}' to {gid}")
        except Exception as e:
            print(f"[WARN] Failed to send alert for '{code}' to {gid}: {e}")


def group_low_fuel(data, team_info):
    """ប្រើតម្លៃចុងក្រោយរបស់ Site នីមួយៗ ហើយដាក់ជាក្រុមតាម Team។
    Return (low_sites, groups) ដែល groups = {team: {"leaders": [...], "sites": [...]}}"""
    latest = {}
    for x in data:  # data តម្រៀបតាមថ្ងៃរួចហើយ => ជួរក្រោយ = តម្លៃថ្មីបំផុត
        latest[x["site"]] = x
    low = [x for x in latest.values() if x["fuel"] is not None and x["fuel"] < SUMMARY_THRESHOLD]
    groups = {}
    for x in low:
        info = team_info.get(x["site"], {})
        team = info.get("team") or "N/A"
        g = groups.setdefault(team, {"leaders": [], "sites": []})
        leader = info.get("leader", "")
        if leader and leader not in g["leaders"]:
            g["leaders"].append(leader)
        g["sites"].append(x)
    return low, groups


def _status_label(fuel):
    if fuel <= 0:
        return "អស់"
    if fuel <= FUEL_ALERT_THRESHOLD:
        return "ទាប"
    return "ត្រូវតាមដាន"


def build_summary_messages(data, team_info):
    """សារសង្ខេបប្រេងទាប (អត្ថបទសុទ្ធ គ្មានសញ្ញា) ដាក់ជាក្រុមតាម Team (ប្រើជា fallback ពេលបង្កើតតារាងរូបភាពមិនបាន)។"""
    low, groups = group_low_fuel(data, team_info)
    if not low:
        return []

    stamp = datetime.now(timezone(timedelta(hours=7))).strftime("%Y-%m-%d %H:%M")
    header = (f"របាយការណ៍ប្រេងទាប (ក្រោម {SUMMARY_THRESHOLD}%)\n"
              f"ថ្ងៃទី {stamp}  សរុប {len(low)} Site")

    blocks = []
    for team in sorted(groups, key=lambda t: (t == "N/A", t)):  # N/A ដាក់ចុងក្រោយ
        g = groups[team]
        sites = sorted(g["sites"], key=lambda x: (x["fuel"], x["site"]))
        title = "Team មិនស្គាល់" if team == "N/A" else f"Team {team}"
        lines = [f"{title} ({len(sites)} Site)"]
        if g["leaders"]:
            lines.append("Bong " + ", ".join(g["leaders"]))
        for x in sites:
            lines.append(f"{x['site']}: {_pct(x['fuel'])}" + (" (អស់)" if x["fuel"] <= 0 else ""))
        blocks.append("\n".join(lines))

    footer = "សូមចាក់ប្រេងជាបន្ទាន់!"
    messages, current = [], header
    for b in blocks:
        if len(current) + len(b) + 2 > TELEGRAM_MAX_LEN:
            messages.append(current)
            current = "របាយការណ៍ប្រេងទាប (បន្ត)"
        current += "\n\n" + b
    messages.append(current + "\n\n" + footer)
    return messages


# ---------------------- តារាងរូបភាព (Table image) ----------------------
TABLE_COLUMNS = ["No", "Team", "Site", "Fuel", "Status", "Team Leader"]
ROW_FILL = {"អស់": "#FFC7CE", "ទាប": "#FFE0B2", "ត្រូវតាមដាន": "#FFF2CC"}
MAX_ROWS_PER_IMAGE = 40  # បំបែកជាច្រើនរូបបើច្រើនជាងនេះ (បំបែកតាមក្រុម Team)


def build_summary_tables(data, team_info):
    """Return (pages, low_count, groups)។ page = (DataFrame, first_row_flags)។
    ក្នុងក្រុមនីមួយៗ Team និង Leader បង្ហាញតែជួរដំបូង ដើម្បីឲ្យតារាងស្អាត។"""
    low, groups = group_low_fuel(data, team_info)
    if not low:
        return [], 0, groups

    team_rows = []  # [[row dict, ...] ក្នុងមួយ Team]
    no = 0
    for team in sorted(groups, key=lambda t: (t == "N/A", t)):
        g = groups[team]
        sites = sorted(g["sites"], key=lambda x: (x["fuel"], x["site"]))
        rows = []
        for i, x in enumerate(sites):
            no += 1
            rows.append({
                "No": no,
                "Team": ("មិនស្គាល់" if team == "N/A" else team) if i == 0 else "",
                "Site": x["site"],
                "Fuel": _pct(x["fuel"]),
                "Status": _status_label(x["fuel"]),
                "Team Leader": ", ".join(g["leaders"]) if i == 0 else "",
                "_first": i == 0,
            })
        team_rows.append(rows)

    pages, current = [], []
    for rows in team_rows:
        if current and len(current) + len(rows) > MAX_ROWS_PER_IMAGE:
            pages.append(current)
            current = []
        current += rows
    if current:
        pages.append(current)
    return pages, len(low), groups


def style_summary_table(rows, title):
    df = pd.DataFrame(rows)
    first = list(df["_first"])
    status = list(df["Status"])
    df = df[TABLE_COLUMNS]

    st = df.style.set_caption(title).set_table_styles([
        {"selector": "caption", "props": [("caption-side", "top"), ("font-size", "22px"), ("font-weight", "bold"),
                                          ("text-align", "center"), ("background-color", "#2EA44E"), ("color", "white"),
                                          ("padding", "10px"), ("border", "1px solid black")]},
        {"selector": "th", "props": [("background-color", "#369388"), ("color", "white"), ("font-weight", "bold"),
                                     ("text-align", "center"), ("border", "1px solid black"), ("padding", "6px 10px")]},
        {"selector": "td", "props": [("text-align", "center"), ("border", "1px solid #888"), ("padding", "5px 10px"),
                                     ("font-size", "16px")]},
    ])

    def row_style(row):
        i = row.name
        out = []
        for col in TABLE_COLUMNS:
            css = f"background-color: {ROW_FILL.get(status[i], '')};" if status[i] in ROW_FILL else ""
            if first[i]:
                css += "border-top: 3px solid black;"
            if col in ("Team", "Team Leader") and first[i]:
                css += "font-weight: bold;"
            if col == "Status":
                css += "font-weight: bold;"
            out.append(css)
        return out

    return st.apply(row_style, axis=1).hide(axis="index")


def export_table_image(rows, title, path):
    import dataframe_image as dfi  # import យឺត: ប្រើតែពេលចាំបាច់
    dfi.export(style_summary_table(rows, title), path, max_rows=-1)


def build_table_caption(groups, low_count, stamp, page_no, page_total):
    """Caption របស់រូប៖ ចំណងជើង + Mention មេធីម (Telegram ជូនដំណឹងពី Caption បាន)។"""
    head = f"របាយការណ៍ប្រេងទាប (ក្រោម {SUMMARY_THRESHOLD}%)\n{stamp}  សរុប {low_count} Site"
    if page_total > 1:
        head += f"  (រូបទី {page_no}/{page_total})"
    lines = []
    for team in sorted(groups, key=lambda t: (t == "N/A", t)):
        if groups[team]["leaders"]:
            lines.append(f"{team}: Bong " + ", ".join(groups[team]["leaders"]))
    caption = head + "\n\n" + "\n".join(lines) + "\n\nសូមចាក់ប្រេងជាបន្ទាន់!"
    return caption[:1000]


async def send_summary_telegram_report(client, data, team_info):
    """ផ្ញើរបាយការណ៍ Summary ប្រេងទាប ជាតារាងរូបភាព (ដាក់ជាក្រុម Team) ចូល Main Group។
    បើបង្កើតរូបមិនបាន ប្រើសារអត្ថបទជំនួស។"""
    if not FUEL_ALERT_GROUP_IDS or not data:
        return
    stamp = datetime.now(timezone(timedelta(hours=7))).strftime("%Y-%m-%d %H:%M")
    try:
        pages, low_count, groups = build_summary_tables(data, team_info)
        if not pages:
            return
        paths = []
        for i, rows in enumerate(pages, start=1):
            title = f"ប្រេងទាប ក្រោម {SUMMARY_THRESHOLD}%  {stamp}"
            path = f"low_fuel_summary_{i}.png"
            export_table_image(rows, title, path)
            paths.append(path)
        for gid in FUEL_ALERT_GROUP_IDS:
            for i, path in enumerate(paths, start=1):
                # Caption (Mention មេធីម) ភ្ជាប់ជាមួយរូបដំបូង រូបបន្តបន្ទាប់មាន Caption ខ្លី
                cap = build_table_caption(groups, low_count, stamp, i, len(paths)) if i == 1 else f"(រូបទី {i}/{len(paths)})"
                await client.send_file(gid, path, caption=cap, parse_mode=None)
            print(f"[OK] Summary table image sent to Main Group ({gid}) - {len(paths)} image(s)")
        return
    except Exception as e:
        print(f"[WARN] Table image failed ({type(e).__name__}: {e}) - fallback to text message")

    messages = build_summary_messages(data, team_info)
    if not messages:
        return
    for gid in FUEL_ALERT_GROUP_IDS:
        for msg in messages:
            try:
                await client.send_message(gid, msg, parse_mode=None, link_preview=False)
            except Exception as e:
                print(f"[WARN] Failed to send summary report to {gid}: {e}")
                break
        else:
            print(f"[OK] Summary report sent to Main Group ({gid}) - {len(messages)} message(s)")


# ============================ REPORTS ============================
def _load_team_map(sh, team_info):
    tmap = {}
    try:
        for r in sh.worksheet(TEAM_LIST_SHEET).get_all_values()[1:]:
            if len(r) > 1 and r[0].strip() and r[1].strip():
                tmap.setdefault(r[0].strip(), r[1].strip())
    except Exception as e:
        print(f"[WARN] មិនអាចអាន sheet '{TEAM_LIST_SHEET}': {e}")
    for site, i in (team_info or {}).items():
        if i.get("team"):
            tmap.setdefault(site, i["team"])
    return tmap


def _parse_date(text):
    text = text.lstrip("'").strip()
    for fmt, n in (("%Y-%m-%d", 10), ("%m/%d/%Y", None), ("%d/%m/%Y", None)):
        try:
            return datetime.strptime(text[:10] if n else text.split(" ")[0], fmt)
        except ValueError:
            continue
    return None


def _read_log(sh, year, month):
    rows = sh.worksheet(OUTPUT_SHEET_NAME).get_all_values()
    out = []
    for r in rows:
        r = r + [""] * (17 - len(r))
        if r[16].strip() != "OK":
            continue
        d = _parse_date(r[0])
        if not d or (d.year, d.month) != (year, month):
            continue
        try:
            fuel = float(r[6])
        except ValueError:
            fuel = None
        out.append({"date": d, "site": r[1].strip(), "fuel": fuel,
                    "fuel_raw": r[6], "batt": r[7]})
    out.sort(key=lambda x: (x["date"], x["site"]))
    return out


def _num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return v


def _get_ws(sh, name, rows=1000, cols=30):
    try:
        return sh.worksheet(name)
    except gspread.WorksheetNotFound:
        return sh.add_worksheet(title=name, rows=rows, cols=cols)


def _write_grid(ws, grid):
    width = max(len(r) for r in grid)
    grid = [r + [""] * (width - len(r)) for r in grid]
    if ws.row_count < len(grid) + 5 or ws.col_count < width:
        ws.resize(rows=max(len(grid) + 20, 100), cols=max(width, 10))
    ws.clear()
    ws.update(values=grid, range_name="A1", value_input_option="RAW")


def build_weekly_report(sh, data, tmap, year, month):
    last_day = calendar.monthrange(year, month)[1]
    ranges = [(1, 7), (8, 14), (15, 21), (22, last_day)]
    tables = []
    for w, (d1, d2) in enumerate(ranges):
        rows = [x for x in data if min((x["date"].day - 1) // 7, 3) == w]
        t = [[f"សប្ដាហ៍ទី {w + 1}"],
             ["From", f"{year}-{month:02d}-{d1:02d}", "To", f"{year}-{month:02d}-{d2:02d}"],
             REPORT_HEADERS]
        for x in rows:
            if x["fuel"] is None:
                remark = "គ្មានទិន្នន័យ" if x["fuel_raw"].strip() else ""
            elif x["fuel"] == 0:
                remark = "Fuel អស់"
            elif x["fuel"] <= LOW_FUEL_THRESHOLD:
                remark = "Fuel ទាប"
            else:
                remark = ""
            t.append([x["date"].strftime("%Y-%m-%d"), x["site"], tmap.get(x["site"], "N/A"),
                      x["fuel"] if x["fuel"] is not None else x["fuel_raw"],
                      _num(x["batt"]), remark])
        tables.append(t)

    height = max(len(t) for t in tables)
    grid = []
    for i in range(height):
        row = []
        for t in tables:
            cells = (t[i] if i < len(t) else [])
            cells = cells + [""] * (6 - len(cells))
            row += cells + [""]
        grid.append(row)

    ws = _get_ws(sh, REPORT_SHEET_NAME, rows=max(height + 20, 200), cols=28)
    _write_grid(ws, grid)
    print(f"[OK] Weekly Report បានបង្កើត ({len(data)} ជួរ, {year}-{month:02d})")


def build_summary_report(sh, data, tmap):
    teams = sorted(set(tmap.values()))
    for x in data:
        t = tmap.get(x["site"], "N/A")
        if t not in teams:
            teams.append(t)

    low = [x for x in data if x["fuel"] is not None and x["fuel"] < SUMMARY_THRESHOLD]
    grid = [["សរុបចំនួនម៉ាស៊ីនដែលប្រេងទាប តាម Team"],
            ["ក្រោម (%)", SUMMARY_THRESHOLD],
            [],
            ["Team", "សប្ដាហ៍ទី 1", "សប្ដាហ៍ទី 2", "សប្ដាហ៍ទី 3", "សប្ដាហ៍ទី 4",
             f"សរុប (Site មិនស្ទួន)", f"បញ្ជី Site ក្រោម {SUMMARY_THRESHOLD}%"]]
    sums = [0, 0, 0, 0, 0]
    for team in teams:
        mine = [x for x in low if tmap.get(x["site"], "N/A") == team]
        weeks = [len({x["site"] for x in mine if min((x["date"].day - 1) // 7, 3) == w})
                 for w in range(4)]
        sites = sorted({x["site"] for x in mine})
        grid.append([team] + weeks + [len(sites), ", ".join(sites)])
        for i, v in enumerate(weeks + [len(sites)]):
            sums[i] += v
    grid.append(["សរុបទាំងអស់"] + sums + [""])

    ws = _get_ws(sh, SUMMARY_SHEET_NAME, rows=200, cols=10)
    _write_grid(ws, grid)
    print(f"[OK] Summary Report បានបង្កើត ({len(teams)} Teams)")


def run_reports(gc, team_info=None):
    try:
        sh = gc.open_by_key(SPREADSHEET_ID)
        now = datetime.now(timezone(timedelta(hours=7)))
        if team_info is None:
            team_info = fetch_team_info()
        tmap = _load_team_map(sh, team_info)
        data = _read_log(sh, now.year, now.month)
        build_weekly_report(sh, data, tmap, now.year, now.month)
        build_summary_report(sh, data, tmap)
        return data
    except Exception as e:
        print(f"[ERROR] Report FAILED: {type(e).__name__}: {e}")
        return []


# ============================== MAIN ==============================
async def main():
    api_id_raw = os.environ.get("TELEGRAM_API_ID")
    api_hash = os.environ.get("TELEGRAM_API_HASH")
    session_str = os.environ.get("TELEGRAM_SESSION")

    if not api_id_raw or not api_hash or not session_str or not STATION_GROUP_ID:
        raise ValueError("Missing essential Telegram Credentials.")

    station_codes = fetch_station_codes()
    if not station_codes:
        return

    team_info = fetch_team_info()
    gc = get_sheets_client()
    ws = get_or_create_output_worksheet(gc)
    data = run_reports(gc, team_info)

    cambodia_tz = timezone(timedelta(hours=7))
    client = TelegramClient(StringSession(session_str), int(api_id_raw), api_hash)
    await client.start()

    try:
        for code in station_codes:
            timestamp = "'" + datetime.now(cambodia_tz).strftime("%Y-%m-%d %H:%M:%S")
            command_text = f"{COMMAND_PREFIX} {code}"

            reply_text = None
            try:
                async with client.conversation(STATION_GROUP_ID, timeout=RESPONSE_TIMEOUT_SEC) as conv:
                    await conv.send_message(command_text)
                    response = await conv.get_reply()
                    reply_text = response.raw_text
            except Exception as e:
                print(f"[WARN] No reply for '{code}': {e}")

            if not reply_text:
                row = [timestamp, code] + [""] * 14 + ["NO REPLY"]
                await append_row_with_retry(ws, row)
                await asyncio.sleep(DELAY_BETWEEN_CODES_SEC)
                continue

            parsed = parse_bot_reply(reply_text)
            row = [timestamp, code] + [parsed.get(col, "") for col in OUTPUT_HEADERS[2:-1]] + ["OK"]
            await append_row_with_retry(ws, row)

            try:
                fuel_level_val = float(parsed.get("Fuel Level(%)", ""))
            except (TypeError, ValueError):
                fuel_level_val = None

            if fuel_level_val is not None and fuel_level_val <= FUEL_ALERT_THRESHOLD:
                await send_fuel_alert(client, code, fuel_level_val, reply_text, team_info)

            data = run_reports(gc, team_info)
            await asyncio.sleep(DELAY_BETWEEN_CODES_SEC)

        # ផ្ញើរបាយការណ៍ Summary សរុបចុងក្រោយ និង Mention ឈ្មោះមេធីមចូល Telegram Main Group
        await send_summary_telegram_report(client, data, team_info)

    finally:
        await client.disconnect()

    print("[OK] Fuel monitor run completed.")


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1 and sys.argv[1] == "report":
        run_reports(get_sheets_client())
    else:
        asyncio.run(main())
