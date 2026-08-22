from datetime import datetime
import os
import dataframe_image as dfi
import pandas as pd
from telethon.sessions import StringSession
from telethon.sync import TelegramClient

# ==========================================
# 1. Google Sheet & Task Mapping
# ==========================================
SHEET_CSV_URL = (
    "YOUR_GOOGLE_SHEET_CSV_URL_HERE"  # ដាក់ Link Publish CSV Google Sheet របស់អ្នក
)

# ឈ្មោះ Task តាមកូដ A1, A2, B4, B5 ...
TASK_NAMES = {
    "A1": "Implement big plan maintenance and set parameter",
    "A2": "Maintenance generator sos & test ATS",
    "A6": "Maintenance air-conditioner",
    "A7": "Test Battery BTS",
    "B4": "Plan Maintenance B4",
    "B5": "Plan Maintenance B5",
    "B9": "Connect new power meter online IMES system",
}

# MAIN GROUP (CHA_Power Dept.) សម្រាប់ផ្ញើតារាង Overall Summary
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
        ("background-color", "#27AE60"),  # ពណ៌បៃតងដូចរូបទី២
        ("color", "black"),
        ("padding", "10px"),
        ("border", "1px solid black"),
        ("font-family", "serif"),
    ],
}


def style_detail_image1(df, title):
  """Style ប្រភេទទី១ (Detail Table)"""
  return df.style.set_caption(title).set_table_styles([
      COMMON_CAPTION_STYLE,
      {
          "selector": "th",
          "props": [
              ("background-color", "#1E8449"),
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
  # ទាញយក Mapping ពី Google Sheet
  df_raw = pd.read_csv(SHEET_CSV_URL)

  # ១. Map Chat ID របស់ Team នីមួយៗ (ពី Column A & B)
  team_chat_ids = {}
  for _, row in df_raw.dropna(subset=["Column A", "Column B"]).iterrows():
    team_name = (
        str(row["Column A"]).strip().replace("_TEAM0", "-T0").replace("_TEAM", "-T")
    )
    team_chat_ids[team_name] = int(row["Column B"])

  # ២. Map Chat ID របស់ Task/Sheet នីមួយៗ (ពី Column Sheet & ChatID)
  task_chat_ids = {}
  for _, row in df_raw.dropna(subset=["Sheet", "ChatID"]).iterrows():
    sheet_code = str(row["Sheet"]).strip()
    task_chat_ids[sheet_code] = int(row["ChatID"])

  now = datetime.now()
  shift_title = "Morning Shift" if now.hour < 12 else "Evening Shift"
  valid_teams = [f"CHA-T0{i}" for i in range(1, 8)]

  api_id = int(os.environ.get("TELEGRAM_API_ID"))
  api_hash = os.environ.get("TELEGRAM_API_HASH")
  session_str = os.environ.get("TELEGRAM_SESSION")

  with TelegramClient(StringSession(session_str), api_id, api_hash) as client:

    # -------------------------------------------------------------
    # ប្រភេទទី១ ៖ ផ្ញើទៅ Team Group នីមួយៗ (តាម Chat ID ក្នុង Column B)
    # -------------------------------------------------------------
    for team_code in valid_teams:
      df_team = (
          df_raw[df_raw["Team"] == team_code].copy()
          if "Team" in df_raw.columns
          else pd.DataFrame()
      )
      target_chat_id = team_chat_ids.get(team_code)

      if not df_team.empty and target_chat_id:
        first_task = (
            df_team["Group task"].iloc[0]
            if "Group task" in df_team.columns
            else "A1"
        )
        title_1 = TASK_NAMES.get(first_task, f"Task {first_task}")

        styled_1 = style_detail_image1(df_team, title_1)
        img_path = f"team_{team_code}.png"
        dfi.export(styled_1.hide(axis="index"), img_path)

        client.send_file(
            target_chat_id,
            img_path,
            caption=f"របាយការណ៍ {team_code} - {shift_title}",
        )

    # -------------------------------------------------------------
    # ប្រភេទទី២ ៖ ផ្ញើទៅ Task Group នីមួយៗ (តាម Chat ID ក្នុង Column D)
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

    # ផ្ញើទៅ Main Group CHA_Power Dept.
    client.send_file(
        MAIN_GROUP_ID,
        img_path_3,
        caption=f"របាយការណ៍សរុបរួម {title_3} - {shift_title}",
    )


if __name__ == "__main__":
  main()
