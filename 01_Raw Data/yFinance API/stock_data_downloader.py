import yfinance as yf
import pandas as pd
from datetime import datetime
import os

class StockDataDownloader:
    """
    A class to download historical stock and ETF data from Yahoo Finance.
    
    This class provides functionality to:
    - Download historical data for multiple stocks/ETFs
    - Specify custom date ranges
    - Save data in Excel formats
    - Handle different time intervals (daily, weekly, monthly)
    """
    
    def __init__(self, output_dir="Stock Data Raw"):
        """
        Initialize the StockDataDownloader.
        
        Args:
            output_dir (str): Directory where downloaded files will be saved
        """
        # Get the directory where the script is located
        script_dir = os.path.dirname(os.path.abspath(__file__))
        self.output_dir = os.path.join(script_dir, output_dir)
        
        # Symbol name mapping
        self.symbol_names = {
            '^GDAXI': 'DAX',
            '^MDAXI': 'MDAX',
            '^SDAXI': 'SDAX'
        }
    
    def get_display_name(self, symbol):
        """Get display name for symbol, fallback to original symbol if not found"""
        return self.symbol_names.get(symbol, symbol)
    
    def download_data(self, symbols, start_date, end_date=None, interval='1d'):
        """
        Download historical data for the given symbols.
        
        Args:
            symbols (list): List of stock/ETF symbols (e.g., ['AAPL', '^GDAXI', '^IXIC'])
            start_date (str): Start date in 'YYYY-MM-DD' format
            end_date (str, optional): End date in 'YYYY-MM-DD' format. Defaults to today
            interval (str): Data interval ('1d' for daily, '1wk' for weekly, '1mo' for monthly)
            
        Returns:
            dict: Dictionary containing DataFrames for each symbol and their currency
        """
        if end_date is None:
            end_date = datetime.now().strftime('%Y-%m-%d')
            
        data = {}
        for symbol in symbols:
            try:
                ticker = yf.Ticker(symbol)
                df = ticker.history(start=start_date, end=end_date, interval=interval)
                if not df.empty:
                    # Remove timezone information from the index
                    df.index = df.index.tz_localize(None)
                    # Get currency information for sheet naming only
                    info = ticker.info
                    currency = info.get('currency', 'EUR')
                    # Round price columns to 2 decimal places (Eurocent-genau)
                    price_columns = ['Open', 'High', 'Low', 'Close']
                    for col in price_columns:
                        if col in df.columns:
                            df[col] = pd.to_numeric(df[col], errors='coerce').round(2)
                    data[symbol] = {'df': df, 'currency': currency}
                    display_name = self.get_display_name(symbol)
                    print(f"Successfully downloaded data for {display_name} ({currency})")
                else:
                    print(f"No data found for {symbol}")
            except Exception as e:
                print(f"Error downloading {symbol}: {str(e)}")
        
        return data
    
    def save_to_excel(self, data):
        """
        Save the downloaded data to an Excel file with multiple sheets.
        
        Args:
            data (dict): Dictionary containing DataFrames for each symbol
        """
        if not data:
            print("No data to save")
            return
            
        # Speichere alle Daten in einer Excel-Datei mit mehreren Sheets
        filename = os.path.join(self.output_dir, "stock_data.xlsx")
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            for symbol, item in data.items():
                df = item['df'].copy()
                currency = item['currency']
                display_name = self.get_display_name(symbol)
                
                # Format date index with strftime
                df.index = df.index.strftime('%d.%m.%Y')
                
                # Sheetname mit Display-Name und Währung
                sheet_name = f"{display_name}_{currency}"
                
                # Write to Excel
                df.to_excel(writer, sheet_name=sheet_name)
                
                # Get the worksheet
                worksheet = writer.sheets[sheet_name]
                
                # Set number format for price columns
                price_columns = ['Open', 'High', 'Low', 'Close']
                for idx, col in enumerate(df.columns):
                    col_letter = chr(65 + idx + 1)  # B, C, D, etc.
                    if col in price_columns:
                        # Nur Zahlen ohne Währungssymbol, 2 Dezimalstellen
                        worksheet.column_dimensions[col_letter].number_format = '#,##0.00'
                    elif col == 'Volume':
                        # Volume als ganze Zahlen
                        worksheet.column_dimensions[col_letter].number_format = '#,##0'
                    # Set column width
                    try:
                        max_length = max(
                            df[col].astype(str).apply(len).max(),
                            len(str(col))
                        )
                        worksheet.column_dimensions[col_letter].width = max_length + 2
                    except:
                        worksheet.column_dimensions[col_letter].width = 12  # Default width
        
        print("Created stock_data.xlsx")
        
        # Zusätzlich: Jede Abfrage als eigene Excel-Datei speichern
        for symbol, item in data.items():
            df = item['df'].copy()
            currency = item['currency']
            display_name = self.get_display_name(symbol)
            
            # Format date index with strftime
            df.index = df.index.strftime('%d.%m.%Y')
            
            # Einzelne Excel-Datei für jedes Symbol mit Display-Name
            single_filename = os.path.join(self.output_dir, f"{display_name}.xlsx")
            with pd.ExcelWriter(single_filename, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name=f"{display_name}_{currency}")
                
                # Get the worksheet
                worksheet = writer.sheets[f"{display_name}_{currency}"]
                
                # Set number format for price columns
                price_columns = ['Open', 'High', 'Low', 'Close']
                for idx, col in enumerate(df.columns):
                    col_letter = chr(65 + idx + 1)  # B, C, D, etc.
                    if col in price_columns:
                        # Nur Zahlen ohne Währungssymbol, 2 Dezimalstellen
                        worksheet.column_dimensions[col_letter].number_format = '#,##0.00'
                    elif col == 'Volume':
                        # Volume als ganze Zahlen
                        worksheet.column_dimensions[col_letter].number_format = '#,##0'
                    # Set column width
                    try:
                        max_length = max(
                            df[col].astype(str).apply(len).max(),
                            len(str(col))
                        )
                        worksheet.column_dimensions[col_letter].width = max_length + 2
                    except:
                        worksheet.column_dimensions[col_letter].width = 12  # Default width
            
            print(f"Created {display_name}.xlsx")

# Example usage
if __name__ == "__main__":
    # Create an instance of the downloader
    downloader = StockDataDownloader()
    
    # Example symbols (you can modify these)
    symbols = [
        '^GDAXI',    # DAX Index
        '^MDAXI',    # MDAX Index
        '^SDAXI',    # SDAX Index
        #'^IXIC',     # NASDAQ Composite
        #'AAPL',      # Apple Inc.
        #'MSFT',      # Microsoft
        #'SPY'        # S&P 500 ETF
    ]
    
    # Example date range (you can modify these)
    start_date = '2022-06-01'
    end_date = '2025-06-15'
    
    # Download the data
    data = downloader.download_data(symbols, start_date, end_date, interval='1d')
    
    # Save to Excel only
    downloader.save_to_excel(data)
