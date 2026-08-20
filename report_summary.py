from datetime import datetime
import os
import dataframe_image as dfi
import pandas as pd
from telethon.sessions import StringSession
from telethon.sync import TelegramClient

# ==========================================
# 1. TASK MAPPING DICTIONARY
# ==========================================
TASK_NAMES = {
    "A1": "Implement big plan maintenance and set parameter",
    "A2": "Maintenance generator sos & test ATS",
    "A5": "វាស់ម៉ាសដី",
    "A6": "Maintenance air-conditioner",
    "A7": "Test Battery BTS",
    "A8": "Maintenance solar",
    "B1": "Updating and standardizing data on PMCD 2.0",
    "B2": "Solve Parameter wrong DC ZTE ZXDU68 V6.0",
    "B3": "Install FAC 5G Ventilation Systems",
    "B4": "DC Connect new on IMES system",
    "B5": "Solve DC Cabinet Loss Data on IMES",
    "B6": "Install Generator new IMES system",
    (
        "B7"
    ): (
        "Deployment of Replacement and Supplementary Works for Improvement"
        " of Electromechanical Power System Stability in 2026"
    ),
    "B9": "Connect new power meter online IMES system",
    "B11": "Swap Generator",
    (
        "B12"
    ): (
        "The optimal deployment of power systems for enclosed BTS stations"
        " in 2021"
    ),
    "B13": "Check AC system of site has power consumption abnormal",
    "B14": "Swap Cabinet for battery and DC mini outdoor",
    "B16": "Swapbattery for site Mainnode",
    "B17": "Connect battery online",
    "B18": "Swap battery Shoto 100Ah",
    "C1": "Survey power system for upgrade cell and New site.",
    "C2": "Solve DAQ, battery and Generator offline",
    "C3": "Report.Branch check online DAQ &Cabinet ZTE on-air new site",
    "C4": "Check SRT have backup power less than 2h",
    (
        "C7"
    ): (
        "Report MFl all failed generators in all branches need to recall to"
        " stock"
    ),
    "C8": "Repair generator (at branch)",
    "E1": "Check operation of site Main Node",
    "E2": "Check status operation ATS (Test ATS)",
    "E3": "Check status operaion new Solar",
    "E4": "Replace and instyall ATS",
    "E5": "DC monitoring connection for a remote station via media converter",
}


# ==========================================
# 2. STYLING FUNCTIONS
# ==========================================
def style_detail_table(df, title):
  return df.style.set_caption(title).set_table_styles([
      {
          "selector": "caption",
          "props": [
              ("caption-side", "top"),
              ("font-size", "18px"),
              ("font-weight", "bold"),
              ("text-align", "center"),
              ("background-color", "#369388"),
              ("color", "black"),
              ("padding", "8px"),
              ("border", "1px solid black"),
          ],
      },
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


def style_summary_table(df, title):
  styler = df.style.set_caption(title).set_table_styles([
      {
          "selector": "caption",
          "props": [
              ("caption-side", "top"),
              ("font-size", "18px"),
              ("font-weight", "bold"),
              ("text-align", "center"),
              ("background-color", "#369388"),
              ("color", "black"),
              ("padding", "8px"),
              ("border", "1px solid black"),
          ],
      },
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
    styles = [""] * len(row)
    if row.name == 0:
      return [
          "color: red; font-style: italic; font-weight: bold;" for _ in row
      ]

    styles[0] = "background-color: #F2F2F2;"
    styles[1] = "background-color: #F2F2F2;"
    styles[5] = (
        "background-color: #A2D9CE; font-weight: bold; font-style: italic;"
    )
    return styles

  return styler.apply(apply_row_styles, axis=1)


# ==========================================
# 3. MAIN EXECUTION
# ==========================================
def main():
  # ទាញយកទិន្នន័យពី Google Sheet
  raw_data = []  # ជំនួសដោយ Data ពី Google Sheet របស់អ្នក
  df = pd.DataFrame(raw_data)

  df_cha = df[df["Branch"] == "CHA"].copy()

  now = datetime.now()
  is_morning = now.hour < 12

  if is_morning:
    df_detail = df_cha[df_cha["Result"] == "Not yet do"].copy()
    shift_title = "Morning Shift"
  else:
    df_detail = df_cha.copy()
    df_detail["sort_key"] = df_detail["Result"].apply(
        lambda x: 0 if x == "Approved" else 1
    )
    df_detail = (
        df_detail.sort_values(by="sort_key").drop(columns=["sort_key"]).copy()
    )
    shift_title = "Evening Shift"

  df_detail["No."] = range(1, len(df_detail) + 1)

  group_task_code = "A6"
  header_title_detail = TASK_NAMES.get(group_task_code, group_task_code)
  styled_detail = style_detail_table(df_detail, header_title_detail)

  detail_img_path = f"report_detail_{group_task_code}.png"
  dfi.export(styled_detail.hide(axis="index"), detail_img_path)

  summary_img_path = None
  if not is_morning:
    summary_header = f"Report Plan ELE M{now.month}"
    valid_teams = [f"CHA-T0{i}" for i in range(1, 8)]

    rows = []
    for idx, team in enumerate(valid_teams, start=1):
      team_data = df_cha[
          (df_cha["Team"] == team) & (df_cha["Group task"] == group_task_code)
      ]
      target = len(team_data)
      approved = len(team_data[team_data["Result"] == "Approved"])
      not_approved = target - approved
      remain = not_approved
      pct_val = (approved / target * 100) if target > 0 else 100.0

      rows.append({
          "No": idx,
          "Branch": team,
          "Target Site": target,
          "Approved": approved,
          "Not Approved": not_approved,
          "%": f"{int(pct_val)}%",
          "Remain": remain,
      })

    tot_target = sum(r["Target Site"] for r in rows)
    tot_app = sum(r["Approved"] for r in rows)
    tot_not_app = sum(r["Not Approved"] for r in rows)
    tot_remain = sum(r["Remain"] for r in rows)
    tot_pct = (tot_app / tot_target * 100) if tot_target > 0 else 100.0

    total_row = {
        "No": "",
        "Branch": "",
        "Target Site": tot_target,
        "Approved": tot_app,
        "Not Approved": tot_not_app,
        "%": f"{int(tot_pct)}%",
        "Remain": tot_remain,
    }

    df_summary = pd.DataFrame([total_row] + rows)
    df_summary.columns = pd.MultiIndex.from_tuples([
        ("", "No"),
        ("", "Branch"),
        ("", "Target Site"),
        ("Result", "Approved"),
        ("Result", "Not Approved"),
        ("", "%"),
        ("", "Remain"),
    ])

    styled_summary = style_summary_table(df_summary, summary_header)
    summary_img_path = "overall_summary_team.png"
    dfi.export(styled_summary.hide(axis="index"), summary_img_path)

  api_id = int(os.environ.get("TELEGRAM_API_ID"))
  api_hash = os.environ.get("TELEGRAM_API_HASH")
  session_str = os.environ.get("TELEGRAM_SESSION")
  group_id = int(os.environ.get("REPORT_NOTIFY_GROUP_ID"))

  caption_text = (
      f"របាយការណ៍លទ្ធផលការងារ - Branch: CHA ({shift_title})\nTask:"
      f" {header_title_detail}"
  )

  with TelegramClient(StringSession(session_str), api_id, api_hash) as client:
    client.send_file(group_id, detail_img_path, caption=caption_text)

    if summary_img_path and os.path.exists(summary_img_path):
      client.send_file(
          group_id,
          summary_img_path,
          caption="របាយការណ៍សរុប (Branch: CHA - Team T01 ដល់ T07)",
      )


if __name__ == "__main__":
  main()
