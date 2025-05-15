import os

# Set the path to your main attachments directory
base_path = r"D:\Users\proyectos1\OneDrive - NCS S.A.S\Escritorio\GR\Support\Migración de JIRA\Repositorio\Proyectos\PAMII v1.0"  # Replace this with your actual path

report = []

for folder_name in os.listdir(base_path):

    folder_path = os.path.join(base_path, folder_name)
    
    if os.path.isdir(folder_path):

        file_count = len(
            [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]
        )
        report.append((folder_name, file_count))


# ==== Print the report ===

# 
print(f"{'Folder':<40} | {'File Count'}")
print("-" * 55)

for folder, count in sorted(report):
    print(f"{folder:<40} | {count}")