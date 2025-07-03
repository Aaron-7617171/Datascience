import pandas as pd
import os
import warnings
import numpy as np

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# CONFIGURATION
input_folder = "Data"
interest_file = "interest_rate_2022_2025.xlsx"
stock_file = "stock_data_combined_onehot.xlsx"
sentiment_file = "ecb_sentiment_analysis.xlsx"
output_folder = "Dataset"
output_file = "DS_14_t_3days_complete.xlsx"

# File paths
os.chdir(os.path.dirname(os.path.abspath(__file__)))

# Load data
interest_path = os.path.join(input_folder, interest_file)
df_interest = pd.read_excel(interest_path, index_col=0, parse_dates=True)

stock_path = os.path.join(input_folder, stock_file)
df_stock = pd.read_excel(stock_path, index_col=0, parse_dates=True)

# Convert data - BOTH indices to same format
onehot_cols = [col for col in df_stock.columns if col.startswith('Index_')]
df_stock[onehot_cols] = df_stock[onehot_cols].astype(float)
df_stock.index = pd.to_datetime(df_stock.index, format='%d.%m.%Y')
df_interest.index = pd.to_datetime(df_interest.index, format='%d.%m.%Y')

# Create df_neu
df_neu = df_stock.copy()
common_dates = df_neu.index.intersection(df_interest.index)
df_neu = df_neu.loc[common_dates]

# Initialize columns for Close_t-14 to Close_t-1 (correct order)
for i in range(1, 15):
    df_neu.insert(i-1, f'Close_t-{15-i}', np.nan)

# Initialize columns for Close_t+1 to Close_t+3
for i in range(1, 4):
    df_neu.insert(15 + i, f'Close_t+{i}', np.nan)

# Fill Close_t-14 to Close_t+3 values
index_columns = ['Index_DAX', 'Index_MDAX', 'Index_SDAX']
for i in range(len(df_neu)):
    row = df_neu.iloc[i]
    date = df_neu.index[i]
    column_with_one = [col for col in index_columns if row[col] == 1.0][0]
    
    search_date = pd.to_datetime(date, dayfirst=True)
    mask = (df_stock.index == search_date) & (df_stock[column_with_one] == 1.0)
    
    if mask.any():
        iloc_position = np.where(mask)[0][0]
        try:
            # Historical prices (backwards)
            for j in range(14, 0, -1):
                df_neu.iloc[i, df_neu.columns.get_loc(f'Close_t-{j}')] = df_stock.iloc[iloc_position - j]['Close']
            
            # Future prices (forwards)
            for j in range(1, 4):
                df_neu.iloc[i, df_neu.columns.get_loc(f'Close_t+{j}')] = df_stock.iloc[iloc_position + j]['Close']
        except IndexError:
            pass

# Add Interest Rate columns
df_neu['Interest Rate_Old'] = df_interest.loc[common_dates, 'Interest Rate_Old']
df_neu['Interest Rate_Change'] = df_interest.loc[common_dates, 'Interest Rate_Change']
#bishier klappt!!


# Load sentiment data and add to df_neu
sentiment_path = os.path.join(input_folder, sentiment_file)
df_sentiment = pd.read_excel(sentiment_path)

# Remove last row from sentiment data (likely summary row)
df_sentiment = df_sentiment.iloc[:-1]

# Convert sentiment Date column to datetime format
df_sentiment['Date'] = pd.to_datetime(df_sentiment['Date'], format='%d_%B_%Y')

# Set Date as index for sentiment data
df_sentiment.set_index('Date', inplace=True)

# Add all sentiment columns to df_neu using for loop
sentiment_columns = ['FinBERT_Sentences', 'FinBERT_Chunks', 'RoBERTa_Sentences', 'RoBERTa_Chunks']
for col in sentiment_columns:
    df_neu[col] = df_neu.index.map(df_sentiment[col])


# Save and display
os.makedirs(output_folder, exist_ok=True)
df_neu.to_excel(f"{output_folder}/{output_file}", index=True)

print("df_neu (final dataset):")
print(df_neu)
print(f"\ndf_neu dimensions: {df_neu.shape[0]} rows × {df_neu.shape[1]} columns")
print(f"✅ File saved: {output_folder}/{output_file}")
