from datetime import datetime
import io
import os
import dataframe_image as dfi
import pandas as pd
import requests
from telethon.sessions import StringSession
from telethon.sync import TelegramClient

# Sheet ID
SPREADSHEET_ID = "1PmMSqfeBWhYJe5dMv3PrLOFKc2YmLYP8BdCvf9FyZX4"

# -------------------------------------------------------------
# កំណត់ចំណងជើងពេញលេញសម្រាប់ Task នីមួយៗ (Updated)
# -------------------------------------------------------------
TASK_NAMES = {
    "A1": "Implement big plan maintenance and set parameter",
    "A2": "Maintenance generator sos & test ATS",
    "A6": "Maintenance air-conditioner",
    "A7": "Test Battery BTS",
    "B1": "Updating and standardizing data on PMCD 2.0",
    "B2": "Solve Parameter wrong DC ZTE ZXDU68 V6.0",
    "B3": "Install ‎FAC 5G Ventilation Systems",
    "B4": "DC Connect new on IMES system",
    "B5": "Solve DC Cabinet Loss Data on IMES",
    "B6": "Install Generator new IMES system",
    "B7": (
        "Deployment of Replacement and Supplementary Works for Improvement of"
        " Electromechanical Power System Stability in 2026"
    ),
    "B9": "Connect new power meter online IMES system",
    "B11": "Swap new generator",
    "B12": (
        "The optimal deployment of power systems for enclosed BTS stations in"
        " 2021"
    ),
    "B13": "Check AC system of site has power consumption abnormal",
    "B14": "Swap Cabinet for battery and DC mini outdoor",
    "B17": "Connect battery online ",
    "C1": "Survey power system for upgrade cell and New site.C1",
    "C2": "Solve DAQ, battery and Generator offline",
    "C3": "Report.Branch check online DAQ &Cabinet ZTE on-air new site",
    "C4": "Check SRT have backup power less than 2h",
}

# Main Group (CHA_Power Dept.)
MAIN_GROUP_ID = -1001853372580
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
  """ទាញយក CSV តាមឈ្មោះ Sheet"""
  url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet={sheet_name_or_gid}"
  headers = {"User-Agent": "Mozilla/5.0"}
  res = requests.get(url, headers=headers)
  if res.status_code == 200:
    return pd.read_csv(io.StringIO(res.text))
  return None


def style_detail_table(df, title):
  """Style សម្រាប់តារាងលម្អិត Site (ផ្ញើទៅ Task Group នីមួយៗ)"""
  styler = df.style.set_caption(title).set_table_styles([
      COMMON_CAPTION_STYLE,
      {
          "selector": "th",
          "props": [
              ("background-color", "#369388"),
              ("color", "black"),
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
  return styler


def style_task_summary(df, title):
  """Style សម្រាប់ Task Summary Table (ផ្ញើទៅ Main Group)"""
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
    if row.name == 0:  # Row សរុប
      return [
          "color: red; font-style: italic; font-weight: bold;" for _ in row
      ]
    styles = [""] * len(row)
    styles[5] = (
        "background-color: #A2D9CE; font-weight: bold; font-style: italic;"
    )
    return styles

  return styler.apply(apply_row_styles, axis=1)


def style_overall_summary(df, title):
  """Style សម្រាប់ Overall Summary Table (ផ្ញើទៅ Main Group)"""
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
    if row.name == len(df) - 1:
      return ["font-weight: bold; background-color: #F2F2F2;"] * len(row)
    return [""] * len(row)

  return styler.apply(apply_total_style, axis=1)


def main():
  # ១. អាន Mapping ChatID ពី Sheet "Team chat IDs"
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

  now = datetime.now()
  shift_title = "Morning Shift" if now.hour < 12 else "Evening Shift"

  overall_stats = {
      team: {"Target": 0, "Approved": 0, "NotApproved": 0}
      for team in VALID_TEAMS
  }

  api_id = int(os.environ.get("TELEGRAM_API_ID"))
  api_hash = os.environ.get("TELEGRAM_API_HASH")
  session_str = os.environ.get("TELEGRAM_SESSION")

  with TelegramClient(StringSession(session_str), api_id, api_hash) as client:

    # -------------------------------------------------------------
    # ដំណើរការតាម Task Sheet នីមួយៗ
    # -------------------------------------------------------------
    for task_code, chat_id in task_chat_ids.items():
      task_title = TASK_NAMES.get(task_code, f"Task {task_code}")
      df_task = fetch_csv(task_code)

      if df_task is None or df_task.empty:
        continue

      # A. បង្កើត និង Filter តារាងលម្អិត Site (លុបជួរ nan និង #N/A)
      cols_to_show = [
          "No.",
          "Group task",
          "Branch",
          "Site name",
          "Q'ty task/Local task",
          "Result",
          "Remark",
          "Last date record",
          "History Task",
          "Team",
      ]
      available_cols = [c for c in cols_to_show if c in df_task.columns]

      if available_cols:
        df_detail = df_task[available_cols].copy()

        # ១. Filter លុបជួរដែលគ្មានទិន្នន័យ (លុបជួរដែល Group task ឬ Site name ឬ Team ទទេ/nan/#N/A)
        if "Group task" in df_detail.columns:
          df_detail = df_detail[
              df_detail["Group task"].notna()
              & (
                  ~df_detail["Group task"]
                  .astype(str)
                  .str.strip()
                  .str.lower()
                  .isin(["", "nan", "none", "#n/a", "n/a"])
              )
          ]

        if "Site name" in df_detail.columns:
          df_detail = df_detail[
              df_detail["Site name"].notna()
              & (
                  ~df_detail["Site name"]
                  .astype(str)
                  .str.strip()
                  .str.lower()
                  .isin(["", "nan", "none", "#n/a", "n/a"])
              )
          ]

        # ២. ជំនួស Cell ទទេ ឬ 'nan' ក្នុង Column ផ្សេងទៀតឱ្យទៅជា Blank ""
        df_detail = df_detail.fillna("")
        df_detail = df_detail.replace(
            to_replace=r"^(?i:nan|none|#n/a|n/a)$", value="", regex=True
        )

        # ប្រសិនបើនៅសល់ទិន្នន័យ ទើបបង្កើតរូបភាព និងផ្ញើទៅ Task Group
        if not df_detail.empty:
          styled_detail = style_detail_table(df_detail, task_title)
          img_detail_path = f"detail_{task_code}.png"

          dfi.export(
              styled_detail.hide(axis="index"), img_detail_path, max_rows=-1
          )

          client.send_file(
              chat_id,
              img_detail_path,
              caption=f"តារាងការងារលម្អិត {task_title} ({shift_title})",
          )

      # B. គណនា Summary សម្រាប់ Task នោះ រួចផ្ញើទៅ MAIN GROUP
      rows = []
      tot_target = tot_approved = tot_not_approved = tot_remain = 0

      for idx, team in enumerate(VALID_TEAMS, start=1):
        target_site = approved = not_approved = 0

        if "Team" in df_task.columns and "Result" in df_task.columns:
          df_team = df_task[
              df_task["Team"].astype(str).str.strip() == team
          ].copy()
          target_site = len(df_team)
          approved = len(
              df_team[
                  df_team["Result"].astype(str).str.strip().str.lower()
                  == "approved"
              ]
          )
          not_approved = len(
              df_team[
                  df_team["Result"].astype(str).str.strip().str.lower()
                  == "not approved"
              ]
          )

        remain = target_site - (approved + not_approved)
        pct_val = (
            f"{int(round((approved / target_site) * 100))}%"
            if target_site > 0
            else "0%"
        )

        overall_stats[team]["Target"] += target_site
        overall_stats[team]["Approved"] += approved
        overall_stats[team]["NotApproved"] += not_approved

        tot_target += target_site
        tot_approved += approved
        tot_not_approved += not_approved
        tot_remain += remain

        rows.append({
            "No": idx,
            "Team": team,
            "Target Site": target_site,
            "Approved": approved,
            "Not Approved": not_approved,
            "%": pct_val,
            "Remain": remain,
            "Remark": "",
        })

      tot_pct = (
          f"{int(round((tot_approved / tot_target) * 100))}%"
          if tot_target > 0
          else "0%"
      )
      total_row = {
          "No": "",
          "Team": "",
          "Target Site": tot_target,
          "Approved": tot_approved,
          "Not Approved": tot_not_approved,
          "%": tot_pct,
          "Remain": tot_remain,
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

      styled_summary = style_task_summary(df_summary, task_title)
      img_summary_path = f"summary_{task_code}.png"
      dfi.export(
          styled_summary.hide(axis="index"), img_summary_path, max_rows=-1
      )

      # ផ្ញើតារាង Task Summary ទៅ MAIN GROUP
      client.send_file(
          MAIN_GROUP_ID,
          img_summary_path,
          caption=f"របាយការណ៍សង្ខេប {task_title} ({shift_title})",
      )

    # -------------------------------------------------------------
    # ផ្ញើតារាង Overall Summary ទៅ MAIN GROUP
    # -------------------------------------------------------------
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

      overall_rows.append({
          "No": idx,
          "Branch": team,
          "Target Site": t,
          "Approved": a,
          "Not Approved": na,
          "%": pct,
          "Remain": r,
      })

    sum_pct = (
        f"{int(round((sum_approved / sum_target) * 100))}%"
        if sum_target > 0
        else "0%"
    )
    overall_rows.append({
        "No": "TOTAL",
        "Branch": "",
        "Target Site": sum_target,
        "Approved": sum_approved,
        "Not Approved": sum_not_approved,
        "%": sum_pct,
        "Remain": sum_remain,
    })

    df_overall = pd.DataFrame(overall_rows)
    styled_overall = style_overall_summary(df_overall, title_3)

    img_overall_path = "overall_report.png"
    dfi.export(styled_overall.hide(axis="index"), img_overall_path, max_rows=-1)

    # ផ្ញើតារាងសរុបរួមទៅ Main Group
    client.send_file(
        MAIN_GROUP_ID,
        img_overall_path,
        caption=f"របាយការណ៍សរុបរួម {title_3} - {shift_title}",
    )


if __name__ == "__main__":
  main()
