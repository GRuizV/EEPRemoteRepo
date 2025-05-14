import pandas as pd

# Load Excel file
file_path = r'D:\Users\proyectos1\OneDrive - NCS S.A.S\Escritorio\GR\Support\CUT OVER ABRIL\Caso - Cargue Masivo Imagenes\Cargue Real\ImagenesFTP - Base 2.xlsx'
output_path = r'D:\Users\proyectos1\OneDrive - NCS S.A.S\Escritorio\GR\Support\CUT OVER ABRIL\Caso - Cargue Masivo Imagenes\Cargue Real\Python - SKU Grouping.xlsx'
sheet_name = 'Cleaned SKUs'
URL_BASE = 'https://ftp.quest.com.co/imagenes/QUEST/ECOMMERCE/'

# Read the Excel sheet
df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')

# Make sure the column with the image names is correctly named
df['SKU_REF'] = df['ImageName'].str.rsplit('-', n=1).str[0]

# Group by SKU_REF and join the image names with a semi conlon
result = df.groupby('SKU_REF')['ImageName'].apply(
    lambda names: ';'.join(f"{URL_BASE}{name}" for name in names)
).reset_index()

# Save the result to a new xlsx (optional)
result.to_excel(output_path, index=False)

# Display the result
print(result.head())


