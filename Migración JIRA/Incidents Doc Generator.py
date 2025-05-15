import pandas as pd
import json
import os
from tqdm import tqdm



# === Config ===
EXCEL_FILE = r"D:\Users\proyectos1\OneDrive - NCS S.A.S\Escritorio\GR\Support\Migración de JIRA\2025-05-13 - JIRA Full Base Cleaned (WAG & CW).xlsx"
ROOT_DIR = r"D:\Users\proyectos1\OneDrive - NCS S.A.S\Escritorio\GR\Support\Migración de JIRA\Repositorio\Proyectos\Desarrollos en Angular"            # Directory where incident folders are

# === Audit Slice Range ===
START_INDEX = 3869     # Set your starting row (inclusive)
END_INDEX = 3880     # Set your ending row (exclusive)

# === Column Keys ===
KEY = "Clave de incidencia"
STATE = "Estado"
PROJECT = "Nombre del proyecto"
ASSIGNEE = "Persona asignada"
CREATED = "Creada"
SUMMARY = "Resumen"
DESCRIPTION = "Descripcion"
COMMENTS = "Conversation Wrapped"
ATTACHMENTS = "attachments_json"  # Make sure this matches your actual column name


# === Load Data ===
df = pd.read_excel(EXCEL_FILE)

# Slice the DataFrame for batch processing
df_slice = df.iloc[START_INDEX:END_INDEX]




# === Loop with Progress Bar ===
for _, row in tqdm(df_slice.iterrows(), total=len(df_slice), desc=f"Processing rows {START_INDEX} to {END_INDEX - 1}"):

    incident_key = str(row.get(KEY, "")).strip()
    folder_path = os.path.join(ROOT_DIR, incident_key)

    # if not os.path.isdir(folder_path):
    #     print(f"Folder not found for incident {incident_key}. Skipping...")
    #     continue

    # Ensure the directory exists
    os.makedirs(folder_path, exist_ok=True)

    estado = row.get(STATE, "")
    proyecto = row.get(PROJECT, "")
    asignado = row.get(ASSIGNEE, "")
    creada = row.get(CREATED, "")
    resumen = row.get(SUMMARY, "")
    descripcion = row.get(DESCRIPTION, "")
    comentarios = row.get(COMMENTS, "")

    # Parse attachments
    adjuntos_raw = row.get(ATTACHMENTS, "")
    attachment_names = []

    if pd.notna(adjuntos_raw):

        try:

            attachments_list = json.loads(adjuntos_raw)

            for item in attachments_list:
                filename = item.get("filename")

                if filename:
                    attachment_names.append(filename)

        except Exception as e:
            attachment_names.append(f"(error leyendo adjuntos: {e})")

    adjuntos_text = "\n".join(attachment_names) if attachment_names else "Sin adjuntos"


    # Final output
    output_text = f"""Incidencia: {incident_key}

Estado: {estado}

Nombre del proyecto: {proyecto}

Persona Asignada: {asignado}

Fecha de Creacion: {creada}


Resumen
{resumen}


Descripcion
{descripcion}


Comentarios
{comentarios}


Adjuntos
{adjuntos_text}
"""

    # Write .txt
    output_file_path = os.path.join(folder_path, f"{incident_key}.txt")

    with open(output_file_path, "w", encoding="utf-8") as f:
        f.write(output_text)