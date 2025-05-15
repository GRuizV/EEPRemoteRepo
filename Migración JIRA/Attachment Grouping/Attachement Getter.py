import pandas as pd
import json
import requests
from pathlib import Path
from tqdm import tqdm  # For progress bar
from decouple import config


# === CONFIG ===
excel_file = r'D:\Users\proyectos1\OneDrive - NCS S.A.S\Escritorio\GR\Support\Migración de JIRA\2025-05-13 - JIRA Full Base Cleaned (With Attachments grouped).xlsx'
sheet_name = "Base"
issue_id_col = "B" 
attachments_col = "attachments_json"
output_dir = r'D:\Users\proyectos1\OneDrive - NCS S.A.S\Escritorio\GR\Support\Migración de JIRA\Tests\1st Test'


# === JIRA Auth ===
JIRA_EMAIL = config('JIRA_EMAIL')
JIRA_API_TOKEN = config('JIRA_API_TOKEN')
auth = (JIRA_EMAIL, JIRA_API_TOKEN)


# === Load Data ===
df = pd.read_excel(excel_file, sheet_name=sheet_name)


# Get actual column name from Excel's column B
issue_id_colname = df.columns[1]



# === Download Loop ===
# for _, row in tqdm(df.iterrows(), total=len(df)):     # Full loop
# for _, row in tqdm(df.iloc[:10].iterrows(), total=10):   # 1st test
for _, row in tqdm(df.iloc[2093:3879].iterrows(), total=1786):

    issue_id = str(row[issue_id_colname])
    attachments_raw = row[attachments_col]

    try:
        attachments = json.loads(attachments_raw)

    except json.JSONDecodeError:
        continue

    if attachments == ["No Attachment"]:
        continue


    # Create directory for issue
    issue_dir = Path(output_dir) / issue_id
    issue_dir.mkdir(parents=True, exist_ok=True)

    for attachment in attachments:

        filename = attachment.get("filename")
        url = attachment.get("url")

        if not filename or not url:
            continue

        # Full path to save
        filepath = issue_dir / filename

        try:
            response = requests.get(url, auth=auth)
            response.raise_for_status()

            with open(filepath, "wb") as f:
                f.write(response.content)


        except Exception as e:
            print(f"Failed to download {filename} for issue {issue_id}: {e}")





