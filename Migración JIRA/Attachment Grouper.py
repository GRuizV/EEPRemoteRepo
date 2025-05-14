import pandas as pd
import json


# === CONFIGURATION ===
file_path = r'D:\Users\proyectos1\OneDrive - NCS S.A.S\Escritorio\GR\Support\Migración de JIRA\2025-05-12 - JIRA Full Base Cleaned.xlsx'
output_path = r'D:\Users\proyectos1\OneDrive - NCS S.A.S\Escritorio\GR\Support\Migración de JIRA\2025-05-13 - JIRA Full Base Cleaned (With Attachments grouped).xlsx'
sheet_name = "Base"  # <-- Change to your actual sheet name


# Load the CSV
df = pd.read_excel(file_path, sheet_name=sheet_name)


# === COLUMN BOUNDS ===
col_start = 18  # Column S (0-indexed)
col_end = 47    # Column AV (inclusive)
attachment_cols = df.columns[col_start:col_end+1]  # pandas slices are exclusive




# === Attachment parser ===
def parse_attachment(field):
    try:
        parts = field.split(";")
        if len(parts) != 4:
            return None
        return {
            "date_added": parts[0].strip(),
            "content_id": parts[1].strip(),
            "filename": parts[2].strip(),
            "url": parts[3].strip()
        }
    except:
        return None
    



# === Main logic ===
def extract_attachments(row):

    attachments = []
    found_valid = False

    for idx, col in enumerate(attachment_cols):

        cell = row[col]

        if not isinstance(cell, str) or "rest/api/3/attachment/content" not in cell:

            # If first column fails, insert notice and break
            if idx == 0:
                return json.dumps(["No Attachment"], ensure_ascii=False)
            
            break  # Otherwise, just stop the loop

        parsed = parse_attachment(cell)

        if parsed:
            attachments.append(parsed)
            found_valid = True

        else:
            break

    return json.dumps(attachments, ensure_ascii=False) if found_valid else json.dumps(["No Attachment"], ensure_ascii=False)




# === Count attachments ===
def count_attachments(attachments_json_str):

    try:
        data = json.loads(attachments_json_str)

        if isinstance(data, list):
            if data == ["No Attachment"]:
                return 0
            
            return len(data)
        
        return 0
      
    except:
        return 0





# === Apply to all rows ===
df["attachments_json"] = df.apply(extract_attachments, axis=1)

# Optional: drop the attachment columns to clean the output
df.drop(columns=attachment_cols, inplace=True)

# Count up the attachments by incident
df["attachment_count"] = df["attachments_json"].apply(count_attachments)




# === Save to new Excel file ===
df.to_excel(output_path, index=False)



