import os
import pandas as pd
from datetime import datetime

# ============================================
# FOLDER INPUT
# ============================================
folder_path_ezb = "TEXT/EZB"  # Change this to your EZB folder name containing date folders
folder_path_fed = "TEXT/FED"  # Change this to your FED folder name containing date folders
# ============================================

def get_folder_choice():
    """
    Ask user to choose between EZB and FED folders
    """
    print("\n📂 Choose folder to analyze:")
    print("1 - EZB (European Central Bank)")
    print("2 - FED (Federal Reserve)")
    
    while True:
        choice = input("\nEnter your choice (1 or 2): ").strip()
        
        if choice == "1":
            print("✅ Selected: EZB")
            return folder_path_ezb, "EZB"
        elif choice == "2":
            print("✅ Selected: FED")
            return folder_path_fed, "FED"
        else:
            print("❌ Invalid choice. Please enter 1 or 2.")

def list_and_process_folders():
    # Get user's folder choice
    folder_path, folder_name = get_folder_choice()
    
    # Get directory where the script file is located
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Build relative path to target folder from script location
    target_path = os.path.join(script_dir, folder_path)
    
    try:
        # Check if target directory exists
        if not os.path.exists(target_path):
            print(f"Error: Directory {target_path} does not exist")
            return None
        
        # Get all items in target directory
        items = os.listdir(target_path)
        
        # Filter only directories (not files)
        folders = [item for item in items if os.path.isdir(os.path.join(target_path, item))]
        
        # Sort folders alphabetically
        folders.sort()
        
        # Print all found folders
        print(f"Found {len(folders)} folders in {folder_name} directory:")
        print("-" * 40)
        for folder in folders:
            print(folder)
        
        # Process folders to extract valid dates
        valid_dates = []
        invalid_folders = []
        
        for folder in folders:
            try:
                # Parse folder name in format "day_Month_year"
                parts = folder.split('_')
                if len(parts) == 3:
                    day, month_name, year = parts
                    
                    # Convert month name to number
                    date_str = f"{day} {month_name} {year}"
                    parsed_date = datetime.strptime(date_str, "%d %B %Y")
                    
                    # Format as dd.mm.yyyy
                    formatted_date = parsed_date.strftime("%d.%m.%Y")
                    valid_dates.append({
                        'folder_name': folder,
                        'date': formatted_date,
                        'datetime': parsed_date
                    })
                else:
                    invalid_folders.append(folder)
                    
            except ValueError:
                # If parsing fails, add to invalid folders
                invalid_folders.append(folder)
        
        # Create DataFrame from valid dates
        if valid_dates:
            df = pd.DataFrame(valid_dates)
            # Sort by datetime
            df = df.sort_values('datetime').reset_index(drop=True)
            
            print(f"\nValid date folders ({len(valid_dates)}):")
            print("-" * 40)
            for _, row in df.iterrows():
                print(f"{row['folder_name']} -> {row['date']}")
        else:
            df = pd.DataFrame(columns=['folder_name', 'date', 'datetime'])
            print("\nNo valid date folders found")
        
        # Print invalid folders
        if invalid_folders:
            print(f"\nInvalid folders ({len(invalid_folders)}):")
            print("-" * 40)
            for folder in invalid_folders:
                print(folder)
        
        # Save DataFrame as XLSX in script directory with folder-specific name
        if not df.empty:
            xlsx_filename = f"{folder_name} Press Release Days.xlsx"
            xlsx_path = os.path.join(script_dir, xlsx_filename)
            # Save only folder_name and date columns to XLSX
            df[['folder_name', 'date']].to_excel(xlsx_path, index=False)
            print(f"\nDataFrame saved as XLSX: {xlsx_path}")
        else:
            print("\nNo data to save - DataFrame is empty")
        
        return df
        
    except PermissionError:
        print(f"Error: No permission to access {target_path}")
        return None
    except Exception as e:
        print(f"Error occurred: {str(e)}")
        return None

# Execute the function
if __name__ == "__main__":
    df_dates = list_and_process_folders()
    
    if df_dates is not None and not df_dates.empty:
        print(f"\nDataFrame created with {len(df_dates)} entries:")
        print("-" * 50)
        print(df_dates[['folder_name', 'date']])
    else:
        print("\nNo DataFrame created or DataFrame is empty")
