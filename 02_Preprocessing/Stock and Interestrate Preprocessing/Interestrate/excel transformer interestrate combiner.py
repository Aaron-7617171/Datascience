import pandas as pd
import os
from datetime import datetime
import warnings

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# ============================================
# CONFIGURATION - EDIT HERE
# ============================================
# Input folder and files
input_folder = "Raw"
change_file = "2022_2025_change.xlsx"
rate_file = "2022_2025_rate.xlsx"

# Output folder and file
output_folder = "Combined"
output_file = "2022_2025_combined.xlsx"
# ============================================

# File paths
os.chdir(os.path.dirname(os.path.abspath(__file__)))

# Load and clean files
def clean_file(file_path):
    df = pd.read_excel(file_path)
    df = df.iloc[13:].reset_index(drop=True)
    df.columns = df.iloc[0]
    df = df.drop(df.index[0]).reset_index(drop=True)
    
    # Rename columns
    df = df.rename(columns={
        "Deposit facility - date of changes (raw data) - Level (FM.D.U2.EUR.4F.KR.DFR.LEV)": "Interest Rate",
        "Deposit facility - date of changes (raw data) - Level (FM.D.U2.EUR.4F.KR.DFR.LEV) - Modified value (Period-to-period change)": "Interest Rate_Change"
    })
    
    # Filter from June 2022
    df['DATE'] = pd.to_datetime(df['DATE'])
    df = df[df['DATE'] > datetime(2022, 5, 31)].reset_index(drop=True)
    
    return df

# Load files
df_change = clean_file(f"{input_folder}/{change_file}")
df_rate = clean_file(f"{input_folder}/{rate_file}")

# Merge data
df = pd.merge(df_rate, df_change[['DATE', 'Interest Rate_Change']], on='DATE', how='left')

# Clean final data
df = df.drop('TIME PERIOD', axis=1)
df['Interest Rate'] = pd.to_numeric(df['Interest Rate'], errors='coerce')
df['Interest Rate_Change'] = pd.to_numeric(df['Interest Rate_Change'], errors='coerce')

# Convert DATE to string format dd.mm.yyyy before saving
df['DATE'] = df['DATE'].dt.strftime('%d.%m.%Y')

# Save to Excel
os.makedirs(output_folder, exist_ok=True)
df.to_excel(f"{output_folder}/{output_file}", index=False)

# Display results
print(f"Dimensions: {df.shape[0]} rows × {df.shape[1]} columns")
print(f"Non-zero changes: {len(df[df['Interest Rate_Change'] != 0.0])} rows")
print(df.head(-10))

print("\nColumn data types:")
for col in df.columns:
    print(f"   • {col}: {df[col].dtype}")

print(f"\n✅ File saved: {output_folder}/{output_file}")
