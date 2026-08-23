from datetime import datetime
import os
import dataframe_image as dfi
import pandas as pd
from telethon.sessions import StringSession
from telethon.sync import TelegramClient

# ==========================================
# 1. Google Sheet & Task Mapping
# ==========================================
# Link CSV នៃ Google Sheet (Tab: Team chat IDs)
SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/1PmMSqfeBWhYJe5dMv3PrLOFKc2YmLYP8BdCvf9FyZX4/export?format=csv&gid=0"

# ឈ្មោះ Task តាមកូដក្នុង Column C (Sheet)
TASK_NAMES = {
    "A1": "Implement big plan maintenance and set parameter",
    "A2": "Maintenance generator sos & test ATS",
    "A6": "Maintenance air-conditioner",
    "A7": "Test Battery BTS",
    "B1": "Plan Maintenance B1",
    "B2": "Plan Maintenance B2",
    "B3": "Plan Maintenance B3",
    "B4": "Plan Maintenance B4",
    "B5": "Plan Maintenance B5",
    "B6": "Plan Maintenance B6",
    "B7": "Plan Maintenance B7",
    "B9": "Connect new power meter online IMES system",
    "B11": "Plan Maintenance B11",
    "B12": "Plan Maintenance B12",
    "B14": "Plan Maintenance B14",
    "B17": "Plan Maintenance B17",
    "C1": "Plan Maintenance C1",
    "C4": "Plan Maintenance C4",
}

# MAIN GROUP (CHA_Power Dept.) សម្រាប់ផ្ញើតារាង Overall Summary (ប្រភេទទី៣)
MAIN_GROUP_ID = -1001853372580

# ==========================================
# 2. STYLING FUNCTIONS WITH TITLE BANNERS
# ==========================================
COMMON_CAPTION_STYLE = {
    "selector": "caption",
    "props": [
        ("caption-side", "top"),
        ("font-size", "22px"),
        ("font-weight", "normal"),
        ("text-align", "center"),
        ("background-color", "#27AE60"),  # ពណ៌បៃតង
        ("color", "black"),
        ("padding", "10px"),
        ("border", "1px solid black"),
        ("font-family", "serif"),
    ],
}


def style_task_summary_image2(df, title):
  """Style ប្រភេទទី២ (Task Summary Table)"""
  styler = df.style.set_caption(title).set_table_styles([
      COMMON_CAPTION_STYLE,
      {
          "selector": "th",
          "props": [
              ("background-color", "#369388"),
              ("color", "black"),
              ("font-weight", "normal"),
              ("text-align", "center"),
              ("border", "1px solid black"),
              ("padding", "6px"),
          ],
      },
      {
          "selector": "td",
          "props": [
              ("text-align", "center"),
              ("border", "1px solid black"),
              ("padding", "5px"),
          ],
      },
  ])

  def apply_row_styles(row):
    if row.name == 0:  # Total Row ខាងលើ
      return [
          "color: red; font-style: italic; font-weight: bold;" for _ in row
      ]
    styles = [""] * len(row)
    styles[5] = (
        "background-color: #A2D9CE; font-weight: bold; font-style: italic;"
    )  # Column %
    return styles

  return styler.apply(apply_row_styles, axis=1)


def style_overall_image3(df, title):
  """Style ប្រភេទទី៣ (Overall Summary Table)"""
  styler = df.style.set_caption(title).set_table_styles([
      COMMON_CAPTION_STYLE,
      {
          "selector": "th",
          "props": [
              ("background-color", "#2EA44E"),
              ("color", "white"),
              ("font-weight", "bold"),
              ("text-align", "center"),
              ("border", "1px solid black"),
              ("padding", "6px"),
          ],
      },
      {
          "selector": "td",
          "props": [
              ("text-align", "center"),
              ("border", "1px solid black"),
              ("padding", "5px"),
          ],
      },
  ])

  def apply_total_style(row):
    if row.name == len(df) - 1:  # TOTAL Row ខាងក្រោម
      return ["font-weight: bold; background-color: #F2F2F2;"] * len(row)
    return [""] * len(row)

  return styler.apply(apply_total_style, axis=1)


# ==========================================
# 3. MAIN EXECUTION
# ==========================================
def main():
  # ទាញយកទិន្នន័យពី Sheet
  df_sheet_ids = pd.read_csv(SHEET_CSV_URL)

  # ទាញយកតែ Column C (Sheet) និង Column D (ChatID) ប៉ុណ្ណោះ
  task_chat_ids = {}
  df_clean = df_sheet_ids.dropna(subset=["Sheet", "ChatID"])

  for _, row in df_clean.iterrows():
    sheet_code = str(row["Sheet"]).strip()
    try:
      chat_id = int(float(str(row["ChatID"]).strip()))
      task_chat_ids[sheet_code] = chat_id
    except ValueError:
      continue

  now = datetime.now()
  shift_title = "Morning Shift" if now.hour < 12 else "Evening Shift"
  valid_teams = [f"CHA-T0{i}" for i in range(1, 8)]

  api_id = int(os.environ.get("TELEGRAM_API_ID"))
  api_hash = os.environ.get("TELEGRAM_API_HASH")
  session_str = os.environ.get("TELEGRAM_SESSION")

  with TelegramClient(StringSession(session_str), api_id, api_hash) as client:

    # -------------------------------------------------------------
    # ប្រភេទទី២ ៖ ផ្ញើទៅ Task Group នីមួយៗ (ទាញតាម Column C & Col D)
    # -------------------------------------------------------------
    for task_code, chat_id in task_chat_ids.items():
      task_title = TASK_NAMES.get(task_code, f"Task {task_code}")

      rows = []
      for idx, team in enumerate(valid_teams, start=1):
        rows.append({
            "No": idx,
            "Team": team,
            "Target Site": 5,
            "Approved": 4,
            "Not Approved": 0,
            "%": "80%",
            "Remain": 1,
            "Remark": "",
        })

      total_row = {
          "No": "",
          "Team": "",
          "Target Site": 35,
          "Approved": 28,
          "Not Approved": 0,
          "%": "80%",
          "Remain": 7,
          "Remark": "",
      }
      df_summary = pd.DataFrame([total_row] + rows)
      df_summary.columns = pd.MultiIndex.from_tuples([
          ("", "No"),
          ("", "Team"),
          ("", "Target Site"),
          ("Result", "Approved"),
          ("Result", "Not Approved"),
          ("", "%"),
          ("", "Remain"),
          ("", "Remark"),
      ])

      styled_2 = style_task_summary_image2(df_summary, task_title)
      img_path_2 = f"task_{task_code}.png"
      dfi.export(styled_2.hide(axis="index"), img_path_2)

      client.send_file(
          chat_id, img_path_2, caption=f"របាយការណ៍ Task {task_code} ({shift_title})"
      )

    # -------------------------------------------------------------
    # ប្រភេទទី៣ ៖ ផ្ញើទៅ MAIN GROUP "CHA_Power Dept." (-1001853372580)
    # -------------------------------------------------------------
    title_3 = f"Report Plan Power M{now.month}"

    overall_data = []
    for idx, team in enumerate(valid_teams, start=1):
      overall_data.append({
          "No": idx,
          "Branch": team,
          "Target Site": 50,
          "Approved": 40,
          "Not Approved": 1,
          "%": "80%",
          "Remain": 9,
      })

    overall_data.append({
        "No": "TOTAL",
        "Branch": "",
        "Target Site": 350,
        "Approved": 280,
        "Not Approved": 7,
        "%": "80%",
        "Remain": 63,
    })

    df_overall = pd.DataFrame(overall_data)
    styled_3 = style_overall_image3(df_overall, title_3)

    img_path_3 = "overall_report.png"
    dfi.export(styled_3.hide(axis="index"), img_path_3)

    client.send_file(
        MAIN_GROUP_ID,
        img_path_3,
        caption=f"របាយការណ៍សរុបរួម {title_3} - {shift_title}",
    )


if __name__ == "__main__":
  main()
