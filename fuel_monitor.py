import asyncio
import io
import json
import os
import re
from datetime import datetime, timezone, timedelta

import gspread
import pandas as pd
import requests
from google.oauth2.service_account import Credentials
from telethon import TelegramClient
from telethon.sessions import StringSession

SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID") or "1PmMSqfeBWhYJe5dMv3PrLOFKc2YmLYP8BdCvf9FyZX4"
STATION_GROUP_ID = int(os.environ.get("STATION_GROUP_ID") or "0")
FUEL_ALERT_GROUP_ID_RAW = os.environ.get("FUEL_ALERT_GROUP_ID")
FUEL_ALERT_GROUP_ID = int(FUEL_ALERT_GROUP_ID_RAW) if FUEL_ALERT_GROUP_ID_RAW else None

STATION_LIST_SHEET = "List Site"
STATION_CODE_COLUMN = "Site name"
OUTPUT_SHEET_NAME = os.environ.get("OUTPUT_SHEET_NAME") or "Fuel Monitor Log"

COMMAND_PREFIX = "/mn"
RESPONSE_TIMEOUT_SEC = 20
DELAY_BETWEEN_CODES_SEC = 180  # ⬅️ ៣ នាទី រវាងការសួរម្តងៗ

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


def fetch_station_codes():
    url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet={STATION_LIST_SHEET.replace(' ', '%20')}"
    headers = {"User-Agent": "Mozilla/5.0"}
    try:
        res = requests.get(url, headers=headers, timeout=15)
    except requests.RequestException as e:
        print(f"⚠️ Failed to fetch station list: {e}")
        return []

    if res.status_code != 200:
        print(f"⚠️ Failed to fetch station list: HTTP {res.status_code}")
        return []

    df = pd.read_csv(io.StringIO(res.text))
    df.columns = df.columns.astype(str).str.strip()

    if STATION_CODE_COLUMN not in df.columns:
        print(f"⚠️ Column '{STATION_CODE_COLUMN}' not found in '{STATION_LIST_SHEET}'. Columns found: {list(df.columns)}")
        return []

    codes = df[STATION_CODE_COLUMN].dropna().astype(str).str.strip()
    codes = [c for c in codes if c and c.lower() not in ("nan", "none")]
    return codes


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

    if FUEL_ALERT_GROUP_ID is None:
        print("⚠️ FUEL_ALERT_GROUP_ID is not set — fuel alerts will be skipped (logged only).")

    station_codes = fetch_station_codes()
    if not station_codes:
        print("⚠️ No station codes found — nothing to do.")
        return
    print(f"ℹ️ Loaded {len(station_codes)} station codes.")

    gc = get_sheets_client()
    ws = get_or_create_output_worksheet(gc)

    cambodia_tz = timezone(timedelta(hours=7))

    client = TelegramClient(StringSession(session_str), api_id, api_hash)
    await client.start()

    try:
        for code in station_codes:
            timestamp = datetime.now(cambodia_tz).strftime("%Y-%m-%d %H:%M:%S")
            command_text = f"{COMMAND_PREFIX} {code}"
            print(f"➡️ Sending '{command_text}' ...")

            reply_text = None
            try:
                async with client.conversation(STATION_GROUP_ID, timeout=RESPONSE_TIMEOUT_SEC) as conv:
                    await conv.send_message(command_text)
                    response = await conv.get_response()
                    reply_text = response.raw_text
            except Exception as e:
                print(f"⚠️ No reply / error for '{code}': {e}")

            if not reply_text:
                row = [timestamp, code, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "NO REPLY"]
                ws.append_row(row)
                await asyncio.sleep(DELAY_BETWEEN_CODES_SEC)
                continue

            parsed = parse_bot_reply(reply_text)
            row = [timestamp, code] + [parsed.get(col, "") for col in OUTPUT_HEADERS[2:-1]] + ["OK"]
            ws.append_row(row)

            fuel_level_raw = parsed.get("Fuel Level(%)", "")
            try:
                fuel_level_val = float(fuel_level_raw)
            except (TypeError, ValueError):
                fuel_level_val = None

            if fuel_level_val is not None and fuel_level_val == 0.0:
                alert_msg = f"⚠️ Fuel Level = 0.0% សម្រាប់ Station: {code}\n{reply_text}"
                if FUEL_ALERT_GROUP_ID:
                    await client.send_message(FUEL_ALERT_GROUP_ID, alert_msg)
                    print(f"🚨 Fuel alert sent for '{code}'.")
                else:
                    print(f"🚨 Fuel Level = 0 for '{code}' but FUEL_ALERT_GROUP_ID not set — alert NOT sent.")

            await asyncio.sleep(DELAY_BETWEEN_CODES_SEC)
    finally:
        await client.disconnect()

    print("✅ Fuel monitor run completed.")


if __name__ == "__main__":
    asyncio.run(main())
