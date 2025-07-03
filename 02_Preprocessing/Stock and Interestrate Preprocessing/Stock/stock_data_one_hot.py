import pandas as pd
import os
import warnings

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# ============================================
# CONFIGURATION - EDIT HERE
# ============================================
# Input folder and files
input_folder = "Raw"
file_one = "stock_data.xlsx"

# Columns to keep (only these columns will be included)
columns_to_keep = ['Date', 'Open', 'Close', ] 
#Colums:  ['Date', 'Open', 'High', 'Low', 'Close', 'Volume']

# Output folder and file
output_folder = "Preprocessed"
output_file = "stock_data_combined_onehot.xlsx"
# ============================================

# File paths
os.chdir(os.path.dirname(os.path.abspath(__file__)))

# Load and combine all sheets
sheet_names = ['DAX_EUR', 'MDAX_EUR', 'SDAX_EUR']
combined_data = []

for sheet in sheet_names:
    # Load each sheet
    df_sheet = pd.read_excel(f"{input_folder}/{file_one}", sheet_name=sheet)
    
    # Keep only specified columns
    df_sheet = df_sheet[columns_to_keep]
    
    # Add index column to identify the source (remove "_EUR" from sheet name)
    df_sheet['Index'] = sheet.replace('_EUR', '')
    
    # Append to combined data
    combined_data.append(df_sheet)

# Combine all sheets into one dataframe
df = pd.concat(combined_data, ignore_index=True)

# One-Hot Encoding for Index column
df_onehot = pd.get_dummies(df, columns=['Index'], prefix='Index', dtype=float)

# Save to Excel
os.makedirs(output_folder, exist_ok=True)
df_onehot.to_excel(f"{output_folder}/{output_file}", index=False)

# Display results
print(f"Dimensions: {df_onehot.shape[0]} rows × {df_onehot.shape[1]} columns")
print(f"Columns kept: {', '.join(columns_to_keep)}")
print(f"One-Hot Encoded Columns: {[col for col in df_onehot.columns if 'Index_' in col]}")
print(f"Sheets combined: {len(sheet_names)} ({', '.join([name.replace('_EUR', '') for name in sheet_names])})")
print(df_onehot.head(-10))

print("\nColumn data types:")
for col in df_onehot.columns:
    print(f"   • {col}: {df_onehot[col].dtype}")

print(f"\n✅ File saved: {output_folder}/{output_file}")
