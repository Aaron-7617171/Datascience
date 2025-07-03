import pandas as pd
import os
import warnings

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# ============================================
# CONFIGURATION - EDIT HERE
# ============================================
# Input folder and files
input_folder = "Combined"
interest_file = "2022_2025_combined.xlsx"

# Date file (same folder as .py)
date_file = "EZB Press Release Days.xlsx"

# Output folder and file
output_folder = "Preprocessed"
output_file = "interest_rate_2022_2025.xlsx"
# ============================================

# File paths
os.chdir(os.path.dirname(os.path.abspath(__file__)))

# Read the interest file
interest_path = os.path.join(input_folder, interest_file)
df_interest = pd.read_excel(interest_path)

# Read the date file (same folder as .py)
df_date = pd.read_excel(date_file)

# Drop the 'folder_name' column from df_date
if 'folder_name' in df_date.columns:
    df_date = df_date.drop(columns=['folder_name'])

# Rename DATE column in df_interest to 'date' for consistency
if 'DATE' in df_interest.columns:
    df_interest = df_interest.rename(columns={'DATE': 'date'})

# Convert date columns to datetime for proper comparison
df_interest['date'] = pd.to_datetime(df_interest['date'], format='%d.%m.%Y')
df_date['date'] = pd.to_datetime(df_date['date'], format='%d.%m.%Y')

# Create df_neu by merging on date - inner join to keep only matching dates
df_neu = pd.merge(df_interest, df_date, on='date', how='inner')

# Sort by date from oldest to newest (ascending)
df_neu = df_neu.sort_values('date', ascending=True).reset_index(drop=True)

# Calculate the correct Interest Rate_Change as difference between consecutive Interest Rate values
df_neu['Interest Rate_Change'] = df_neu['Interest Rate'].diff().fillna(0)

# Rename 'Interest Rate' to 'Interest Rate_Old'
df_neu = df_neu.rename(columns={'Interest Rate': 'Interest Rate_Old'})

# Rename 'date' to 'Date'
df_neu = df_neu.rename(columns={'date': 'Date'})

# Convert date back to string format dd.mm.yyyy for display
df_neu['Date'] = df_neu['Date'].dt.strftime('%d.%m.%Y')

# Save to Excel
os.makedirs(output_folder, exist_ok=True)
df_neu.to_excel(f"{output_folder}/{output_file}", index=False)

# Display entire dataframe
print("Complete combined dataframe df_neu (oldest to newest) with Date and Interest Rate_Old:")
print(df_neu)

print(f"\nCombined dataframe dimensions: {df_neu.shape[0]} rows × {df_neu.shape[1]} columns")

print("\nCombined dataframe columns:")
for col in df_neu.columns:
    print(f"   • {col}: {df_neu[col].dtype}")

print(f"\n✅ File saved: {output_folder}/{output_file}")
