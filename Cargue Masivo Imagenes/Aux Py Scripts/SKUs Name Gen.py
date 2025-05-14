import pandas as pd
import random
import string

# Load Excel file
file_path = r'D:\Users\proyectos1\OneDrive - NCS S.A.S\Escritorio\GR\Support\CUT OVER ABRIL\Caso - Cargue Masivo Imagenes\SKUs Simulation.xlsx'
sheet_name = 'Test SKUs'

# Read the Excel sheet
df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')

# Extract the column of SKUs (assuming the header is 'SKUs')
sku_list = df['SKUs'].dropna().astype(str).tolist()


# Define function to generate a random suffix
def random_suffix(length=6):
    return ''.join(random.choices(string.ascii_letters + string.digits, k=length))

# Build dataset
data = []
for sku in sku_list:
    for i in range(5):  # 5 picture names per SKU
        suffix = random_suffix()
        image_name = f"{sku}_{suffix}"
        data.append({"SKU": sku, "ImageName": image_name})

# Convert to DataFrame
df = pd.DataFrame(data)

# Save to Excel
df.to_excel("sku_image_mapping.xlsx", index=False)



