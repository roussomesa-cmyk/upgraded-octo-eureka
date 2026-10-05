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

# ផ្ញើ Alert ទៅច្រើនក្រុម ញែកដោយក្បៀស៖ -1001111,-1002222
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
TEAM_LIST_COL = 2                 # allteam column B = Team
LOW_FUEL_THRESHOLD = 15           # Remark ក្នុង Weekly Report
SUMMARY_THRESHOLD = 20            # Summary: ក្រោម 20%
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


# ====================== LIST SITE / TEAM INFO ======================
def _fetch_list_site_df():
    url = (f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq"
           f"?tqx=out:csv&sheet={STATION_LIST_SHEET.replace(' ', '%20')}")
    try:
        res = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=15)
    except requests.RequestException as e:
        print(f"⚠️ Failed to fetch station list: {e}")
        return None
    if res.status_code != 200:
        print(f"⚠️ Failed to fetch station list: HTTP {res.status_code}")
        return None
    df = pd.read_csv(io.StringIO(res.text))
    df.columns = df.columns.astype(str).str.strip()
    return df


def fetch_station_codes():
    df = _fetch_list_site_df()
    if df is None:
        return []
    if STATION_CODE_COLUMN not in df.columns:
        print(f"⚠️ Column '{STATION_CODE_COLUMN}' not found. Columns found: {list(df.columns)}")
        return []
    codes = df[STATION_CODE_COLUMN].dropna().astype(str).str.strip()
    return [c for c in codes if c and c.lower() not in ("nan", "none")]


def fetch_team_info():
    """Return {site_name: {"team": ..., "leader": ...}} ពី List Site."""
    df = _fetch_list_site_df()
    if df is None:
        return {}
    # header ក្នុង Sheet សរសេរ "Team Leatder" ដូច្នេះរកតាមបុព្វបទ
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
        raise ValueError("GCP_SA_KEY is missing — required to write to Google Sheet.")
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
        print(f"ℹ️ Created new worksheet '{OUTPUT_SHEET_NAME}' with headers.")
    return ws


async def append_row_with_retry(ws, row, max_attempts=3):
    """សរសេរជួរចូល Sheet ព្យាយាមម្តងទៀតបើមាន error បណ្តោះអាសន្ន."""
    for attempt in range(1, max_attempts + 1):
        try:
            # USER_ENTERED ដើម្បីឱ្យ Fuel/Battery ជាលេខ (Timestamp មាន ' នាំមុខ ដើម្បីរក្សាជាអត្ថបទ)
            ws.append_row(row, value_input_option="USER_ENTERED")
            return True
        except Exception as e:
            print(f"⚠️ Sheet write attempt {attempt} failed: {e}")
            if attempt < max_attempts:
                wait_sec = 10 * attempt
                print(f"   Retrying in {wait_sec}s...")
                await asyncio.sleep(wait_sec)
            else:
                print(f"❌ Sheet write failed after {max_attempts} attempts — skipping row.")
                return False


# ============================ FUEL ALERT ============================
async def send_fuel_alert(client, code, fuel_val, reply_text, team_info):
    info = team_info.get(code, {})
    leader = info.get("leader", "")
    team = info.get("team", "")
    msg = (
        f"🚨 Fuel ទាប! Site: {code}\n"
        f"⛽ Fuel Level = {fuel_val}%\n"
        + (f"👥 Team: {team}\n" if team else "")
        + (f"👤 Bong {leader}\n" if leader else "")
        + f"\n{reply_text}"
    )
    if not FUEL_ALERT_GROUP_IDS:
        print(f"🚨 Fuel {fuel_val}% for '{code}' but FUEL_ALERT_GROUP_ID not set — alert NOT sent.")
        return
    for gid in FUEL_ALERT_GROUP_IDS:
        try:
            await client.send_message(gid, msg)
            print(f"🚨 Fuel alert sent for '{code}' → {gid}")
        except Exception as e:
            print(f"⚠️ Failed to send alert for '{code}' to {gid}: {e}")


# ============================ REPORTS (Python គណនា រួចសរសេរតម្លៃចូល Sheet) ============================
def _load_team_map(sh, team_info):
    """Site -> Team ពី sheet allteam (A=Site, B=Team); បើអត់មានប្រើ Stock code ពី List Site."""
    tmap = {}
    try:
        for r in sh.worksheet(TEAM_LIST_SHEET).get_all_values()[1:]:
            if len(r) > 1 and r[0].strip() and r[1].strip():
                tmap.setdefault(r[0].strip(), r[1].strip())
    except Exception as e:
        print(f"⚠️ មិនអាចអាន sheet '{TEAM_LIST_SHEET}': {e}")
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
            row += cells + [""]          # column ទំនេរមួយខណ្ឌចែក
        grid.append(row)

    ws = _get_ws(sh, REPORT_SHEET_NAME, rows=max(height + 20, 200), cols=28)
    _write_grid(ws, grid)
    try:
        for w in range(4):
            c0 = w * 7
            hdr = f"{rowcol_to_a1(3, c0 + 1)}:{rowcol_to_a1(3, c0 + 6)}"
            ws.format(hdr, {"textFormat": {"bold": True},
                            "backgroundColor": {"red": 0.2, "green": 0.65, "blue": 0.55}})
            ws.format(rowcol_to_a1(1, c0 + 1), {"textFormat": {"bold": True, "fontSize": 12}})
    except Exception as e:
        print(f"ℹ️ format skipped: {e}")
    print(f"✅ Weekly Report បានបង្កើត ({len(data)} ជួរ, {year}-{month:02d})")


def build_summary_report(sh, data, tmap):
    teams = sorted(set(tmap.values()))
    for x in data:                       # ធានាថា Team ទាំងអស់ក្នុងទិន្នន័យមាន
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
    try:
        ws.format("A4:G4", {"textFormat": {"bold": True},
                            "backgroundColor": {"red": 0.2, "green": 0.65, "blue": 0.55}})
        ws.format(f"A{len(grid)}:G{len(grid)}", {"textFormat": {"bold": True}})
        ws.format("A1", {"textFormat": {"bold": True, "fontSize": 13}})
    except Exception as e:
        print(f"ℹ️ format skipped: {e}")
    print(f"✅ Summary Report បានបង្កើត ({len(teams)} Teams)")


def run_reports(gc, team_info=None):
    """Report សរសេរចូល Google Sheet ប៉ុណ្ណោះ — មិនផ្ញើ Telegram."""
    try:
        sh = gc.open_by_key(SPREADSHEET_ID)
        now = datetime.now(timezone(timedelta(hours=7)))
        if team_info is None:
            team_info = fetch_team_info()
        tmap = _load_team_map(sh, team_info)
        data = _read_log(sh, now.year, now.month)
        build_weekly_report(sh, data, tmap, now.year, now.month)
        build_summary_report(sh, data, tmap)
    except Exception as e:
        print(f"❌ Report FAILED: {type(e).__name__}: {e}")


# ============================== MAIN ==============================
async def main():
    api_id_raw = os.environ.get("TELEGRAM_API_ID")
    api_hash = os.environ.get("TELEGRAM_API_HASH")
    session_str = os.environ.get("TELEGRAM_SESSION")

    if not api_id_raw:
        raise ValueError("TELEGRAM_API_ID is missing or empty.")
    api_id = int(api_id_raw)
    if not api_hash:
        raise ValueError("TELEGRAM_API_HASH is missing or empty.")
    if not session_str:
        raise ValueError("TELEGRAM_SESSION is missing or empty.")
    if not STATION_GROUP_ID:
        raise ValueError("STATION_GROUP_ID is missing or empty — set this GitHub secret first.")

    if not FUEL_ALERT_GROUP_IDS:
        print("⚠️ FUEL_ALERT_GROUP_ID is not set — fuel alerts will be skipped (logged only).")

    station_codes = fetch_station_codes()
    if not station_codes:
        print("⚠️ No station codes found — nothing to do.")
        return
    print(f"ℹ️ Loaded {len(station_codes)} station codes.")
    team_info = fetch_team_info()

    gc = get_sheets_client()
    ws = get_or_create_output_worksheet(gc)
    run_reports(gc, team_info)   # បង្កើត Report មុន

    cambodia_tz = timezone(timedelta(hours=7))

    client = TelegramClient(StringSession(session_str), api_id, api_hash)
    await client.start()

    try:
        for code in station_codes:
            timestamp = "'" + datetime.now(cambodia_tz).strftime("%Y-%m-%d %H:%M:%S")
            command_text = f"{COMMAND_PREFIX} {code}"
            print(f"➡️ Sending '{command_text}' ...")

            reply_text = None
            try:
                async with client.conversation(STATION_GROUP_ID, timeout=RESPONSE_TIMEOUT_SEC) as conv:
                    await conv.send_message(command_text)
                    response = await conv.get_reply()
                    reply_text = response.raw_text
            except Exception as e:
                print(f"⚠️ No reply / error for '{code}': {e}")

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

            # Fuel Alert → ផ្ញើទៅក្រុម Telegram (មានតែនេះទេដែលផ្ញើ)
            if fuel_level_val is not None and fuel_level_val <= FUEL_ALERT_THRESHOLD:
                await send_fuel_alert(client, code, fuel_level_val, reply_text, team_info)

            # ធ្វើបច្ចុប្បន្នភាព Report ក្រោយរាល់ Site (សរសេរតែក្នុង Sheet)
            run_reports(gc, team_info)

            await asyncio.sleep(DELAY_BETWEEN_CODES_SEC)
    finally:
        await client.disconnect()

    # Report → សរសេរក្នុង Google Sheet ប៉ុណ្ណោះ (មិនផ្ញើ Telegram)
    run_reports(gc, team_info)

    print("✅ Fuel monitor run completed.")


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1 and sys.argv[1] == "report":
        # python fuel_monitor.py report  → បង្កើតតែ Report (មិនប្រើ Telegram)
        run_reports(get_sheets_client())
    else:
        asyncio.run(main())
