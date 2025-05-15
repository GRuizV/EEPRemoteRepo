import pandas as pd
import re





# === Constants ===
excel_file = r'D:\Users\proyectos1\OneDrive - NCS S.A.S\Escritorio\GR\Support\Migración de JIRA\2025-05-12 - JIRA Full Base Cleaned.xlsx'
output_path = r'D:\Users\proyectos1\OneDrive - NCS S.A.S\Escritorio\GR\Support\Migración de JIRA\2025-05-12 - JIRA Full Base Cleaned (Conversations Wrapped).xlsx'
sheet_name = "Base"




# Load your Excel file
df = pd.read_excel(excel_file, sheet_name=sheet_name)


# Your user dictionary
user_dict = {
    "5fdcdede9edf280075d97b7f": "Tecficon SAS",
    "5fe119b934847e00696a3ae4": "RENE BRAUSSIN",
    "6038130dd416ea007016e85a": "Olga D Serrano",
    "603e7cd9460631006a62b7c8": "Juan Sebastian Gomez",
    "60539fe306cbba006a0d7072": "Pruebas MP",
    "60ba5f2593e3f50071a20ec0": "JUAN CAMILO ECHEVERRI DELGADO",
    "610869388c15ca006f650582": "jeniffer paz",
    "616f1bda892c420072f43ff4": "dina",
    "630fbc8455b0a9e29f1acbbf": "Antonio Jose Rodriguez",
    "63654b0ad66d8108a1264eee": "Carlos Augusto Duque Alvarez",
    "63a21bea7cde7bff9d76d39c": "Daniela Grajales",
    "642605077222b08f3e73f9d5": "Yenny Paola Quiñonez",
    "712020:634ed37b-36a7-49bf-9dce-dee09af76297": "Juan Jose Sanchez",
    "712020:a7c80d19-bc13-429f-94d5-a4e00a2bb4d0": "Nicolas Gonzalez Millan",
    "712020:e0cbe562-6fff-4be7-b6bc-bb99e5423cb5": "Emanuel Beltran",
    "712020:e364f553-0a23-4135-bdbb-b1a34a1582fa": "JOHN PENAGOS- Soporte PAMII"
}


# Define conversation columns
conversation_cols = df.columns[df.columns.get_loc("Comentario"):df.columns.get_loc("Comentario_141") + 1]




# === Main Logics Function ===

# Comments Parser
def parse_comment(raw_comment):

    try:
        timestamp, author_id, message = raw_comment.split(";", 2)

    except ValueError:
        return None  # Skip malformed comment


    # Replace the author
    author_name = user_dict.get(author_id.strip(), author_id.strip())
    author_str = f"@{author_name}"


    # Replace all mentions in the message
    def replace_mention(match):

        account_id = match.group(1)
        return f"@{user_dict.get(account_id.strip(), account_id.strip())}"


    message = re.sub(r"\[~accountid:(.*?)\]", replace_mention, message)


    return f"[{timestamp.strip()}] {author_str}: {message.strip()}"


# Conversation Grouper
def group_conversation(row):

    comments = []

    for col in conversation_cols:
        
        raw = row.get(col)

        if pd.notna(raw):
            parsed = parse_comment(raw)

            if parsed:
                comments.append(parsed)

    return "\n".join(comments)




# === Main Event Loop ===

# Create the new column with wrapped conversation
df["Conversation Wrapped"] = df.apply(group_conversation, axis=1)


# Save to new Excel file 
df.to_excel(output_path, index=False)