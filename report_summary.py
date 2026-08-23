import dataframe_image as dfi
import pandas as pd

# ១. រៀបចំទិន្នន័យ Task របស់អ្នក
tasks = [
    ("A1", "Implement big plan maintenance and set parameter"),
    ("A2", "Maintenance generator sos & test ATS"),
    ("A5", "វាស់ម៉ាសដី"),
    ("A6", "Maintenance air-conditioner"),
    ("A7", "Test Battery BTS"),
    ("A8", "Maintenance solar"),
    ("B1", "Updating and standardizing data on PMCD 2.0"),
    ("B2", "Solve Parameter wrong DC ZTE ZXDU68 V6.0"),
    ("B3", "Install FAC 5G Ventilation Systems"),
    ("B4", "DC Connect new on IMES system"),
    ("B5", "Solve DC Cabinet Loss Data on IMES"),
    ("B6", "Install Generator new IMES system"),
    (
        "B7",
        (
            "Deployment of Replacement and Supplementary Works for Improvement"
            " of Electromechanical Power System Stability in 2026"
        ),
    ),
    ("B9", "Connect new power meter online IMES system"),
    ("B11", "Swap Generator"),
    (
        "B12",
        (
            "The optimal deployment of power systems for enclosed BTS stations"
            " in 2021"
        ),
    ),
    ("B13", "Check AC system of site has power consumption abnormal"),
    ("B14", "Swap Cabinet for battery and DC mini outdoor"),
    ("B16", "Swapbattery for site Mainnode"),
    ("B17", "Connect battery online"),
    ("B18", "Swap battery Shoto 100Ah"),
    ("C1", "Survey power system for upgrade cell and New site."),
    ("C2", "Solve DAQ, battery and Generator offline"),
    (
        "C3",
        "Report.Branch check online DAQ &Cabinet ZTE on-air new site",
    ),
    ("C4", "Check SRT have backup power less than 2h"),
    (
        "C7",
        "Report MFl all failed generators in all branches need to recall to"
        " stock",
    ),
    ("C8", "Repair generator (at branch)"),
    ("E1", "Check operation of site Main Node"),
    ("E2", "Check status operation ATS (Test ATS)"),
    ("E3", "Check status operaion new Solar"),
    ("E4", "Replace and instyall ATS"),
    (
        "E5",
        "DC monitoring connection for a remote station via media converter",
    ),
]

# បង្កើត DataFrame
data = []
for idx, (code, name) in enumerate(tasks, 1):
  data.append({
      'No.': idx,
      'Group task': code,
      'Branch': 'CHA',
      'Site name': f'CHA03{idx:02d}',
      "Q'ty task/site": '',
      'Result': 'Approved' if idx <= 5 else 'Not yet do',
      'Remark': '',
      'Team': 'CHA-T05',
  })

df = pd.DataFrame(data)

# ២. កំណត់ Style តារាងឱ្យមានចំណងជើងក្បាលតារាង និងពណ៌បៃតង
styled_df = df.style.set_caption(
    'របាយការណ៍លទ្ធផលការងារ (Evening Progress)'
).set_table_styles([
    {
        'selector': 'caption',
        'props': [
            ('caption-side', 'top'),
            ('font-size', '16px'),
            ('font-weight', 'bold'),
            ('text-align', 'center'),
            ('background-color', '#1B4D3E'),  # ពណ៌បៃតងចាស់ក្បាលលើ
            ('color', 'white'),
            ('padding', '10px'),
            ('border', '1px solid #1B4D3E'),
        ],
    },
    {
        'selector': 'th',
        'props': [
            ('background-color', '#2E7D32'),  # ពណ៌ក្បាល Table Header
            ('color', 'white'),
            ('font-weight', 'bold'),
            ('text-align', 'center'),
            ('border', '1px solid #ccc'),
        ],
    },
    {
        'selector': 'td',
        'props': [('text-align', 'center'), ('border', '1px solid #ccc')],
    },
]).hide(axis='index')

# ៣. Export ជា រូបភាព
dfi.export(styled_df, 'report_table.png')

# ៤. សារអក្សរផ្ញើទៅ Telegram (គ្មាន Emoji 🟢 និង 📍)
caption_text = (
    'របាយការណ៍លទ្ធផលការងារ (Evening Progress)\nSheet: B1 | Team: CHA-T05'
    ' (20/Aug/2026)'
)

# 代码សម្រាប់រ៉ាន់ផ្ញើតាម Telegram:
# await client.send_file(group_id, 'report_table.png', caption=caption_text)
