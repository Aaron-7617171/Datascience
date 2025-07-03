from pypdf import PdfReader
import os
import re
import glob

# ============================================
# FOLDER INPUT - BATCH PROCESSING
# ============================================
pdf_folder_ezb = "PDF\EZB"  # Change this to your folder name containing PDFs
pdf_folder_fed = "PDF\FED"  # Change this to your folder name containing PDFs
# ============================================

# .py program must be in the same folder as the PDF folder and "TEXT" folders

# Change to the folder of the .py file
script_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(script_dir)
print(f"📁 Working directory changed to: {script_dir}")

def get_central_bank_choice():
    """
    Ask user to choose between EZB and FED
    """
    print("\n🏦 Choose Central Bank:")
    print("1 - EZB (European Central Bank)")
    print("2 - FED (Federal Reserve)")
    
    while True:
        choice = input("\nEnter your choice (1 or 2): ").strip()
        
        if choice == "1":
            print("✅ Selected: EZB")
            return "EZB"
        elif choice == "2":
            print("✅ Selected: FED")
            return "FED"
        else:
            print("❌ Invalid choice. Please enter 1 or 2.")

def get_date_format_choice():
    """
    Ask user to choose date format
    """
    print("\n📅 Choose Date Format:")
    print("1 - Month as text (e.g., 17_April_2025)")
    print("2 - Month as number (e.g., 17_04_2025)")
    
    while True:
        choice = input("\nEnter your choice (1 or 2): ").strip()
        
        if choice == "1":
            print("✅ Selected: Month as text")
            return "text"
        elif choice == "2":
            print("✅ Selected: Month as number")
            return "number"
        else:
            print("❌ Invalid choice. Please enter 1 or 2.")

def convert_month_to_number(month_name):
    """
    Convert month name to number with leading zero
    """
    month_mapping = {
        'january': '01', 'jan': '01',
        'february': '02', 'feb': '02',
        'march': '03', 'mar': '03',
        'april': '04', 'apr': '04',
        'may': '05',
        'june': '06', 'jun': '06',
        'july': '07', 'jul': '07',
        'august': '08', 'aug': '08',
        'september': '09', 'sep': '09',
        'october': '10', 'oct': '10',
        'november': '11', 'nov': '11',
        'december': '12', 'dec': '12'
    }
    
    return month_mapping.get(month_name.lower(), month_name)

def extract_date_from_filename(pdf_name, date_format):
    """
    Extract date from PDF filename for PRESS CONFERENCE files
    Expected format: "PRESS CONFERENCE_6_March_2025.pdf"
    """
    # Remove .pdf extension and extract the part after "PRESS CONFERENCE_"
    basename = os.path.splitext(os.path.basename(pdf_name))[0]
    
    if basename.startswith("PRESS CONFERENCE_"):
        # Extract date part after "PRESS CONFERENCE_"
        date_part = basename[len("PRESS CONFERENCE_"):]
        
        # Try to parse date in format "6_March_2025"
        parts = date_part.split('_')
        if len(parts) == 3:
            day, month_name, year = parts
            day = day.zfill(2)  # Add leading zero if needed
            
            if date_format == "number":
                # Convert month name to number and format as DD_MM_YYYY
                month_number = convert_month_to_number(month_name)
                formatted_date = f"{day}_{month_number}_{year}"
            else:
                # Keep month as text and format as DD_Month_YYYY
                formatted_date = f"{day}_{month_name}_{year}"
            
            print(f"📅 Extracted date from filename: {formatted_date}")
            return formatted_date
    
    return None

def extract_date_from_text(text, date_format, pdf_name):
    """
    Extracts the date that appears after "Combined monetary policy decisions and statement"
    If no date found in text and PDF is PRESS CONFERENCE, extract from filename
    """
    # Search for the specific text and the date after it
    pattern = r'Combined monetary policy decisions and\s*statement\s*(\d{1,2}\s+\w+\s+\d{4})'
    
    match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
    if match:
        date_str = match.group(1).strip()
        print(f"📅 Found date after 'Combined monetary policy decisions and statement': {date_str}")
        
        # Parse the date: "17 April 2025"
        date_parts = date_str.split()
        if len(date_parts) == 3:
            day = date_parts[0].zfill(2)  # Add leading zero if needed
            month_name = date_parts[1]
            year = date_parts[2]
            
            if date_format == "number":
                # Convert month name to number and format as DD_MM_YYYY
                month_number = convert_month_to_number(month_name)
                formatted_date = f"{day}_{month_number}_{year}"
            else:
                # Keep month as text and format as DD_Month_YYYY
                formatted_date = f"{day}_{month_name}_{year}"
            
            print(f"📅 Formatted date: {formatted_date}")
            return formatted_date
        else:
            # Fallback: use original format if parsing fails
            formatted_date = date_str.replace(" ", "_")
            return formatted_date
    
    # No date found in text - check if it's a PRESS CONFERENCE file
    basename = os.path.basename(pdf_name)
    if basename.startswith("PRESS CONFERENCE"):
        print("📅 No date found in text, but PDF is PRESS CONFERENCE - trying to extract from filename")
        filename_date = extract_date_from_filename(pdf_name, date_format)
        if filename_date:
            return filename_date
    
    # Fallback if no date found anywhere - include PDF name
    pdf_basename = os.path.splitext(os.path.basename(pdf_name))[0]  # Remove .pdf extension
    fallback_name = f"xxxx_no date found__{pdf_basename}"
    print(f"⚠️ No date found in text or filename, using: {fallback_name}")
    return fallback_name

def get_unique_folder_name(base_path, folder_name):
    """
    Creates a unique folder name by adding 'new' if folder already exists
    """
    original_path = os.path.join(base_path, folder_name)
    
    # If folder doesn't exist, use original name
    if not os.path.exists(original_path):
        return folder_name
    
    # If folder exists, add 'new' until we find a free name
    current_name = folder_name
    while True:
        new_folder_name = f"{current_name}new"
        new_path = os.path.join(base_path, new_folder_name)
        if not os.path.exists(new_path):
            print(f"📁 Folder '{folder_name}' already exists, using '{new_folder_name}' instead")
            return new_folder_name
        current_name = new_folder_name

def extract_sections_precise(text, headers):
    """
    Extracts sections only when headers appear alone on a line
    Special handling for Conclusion section which ends at "We are now ready to take your questions."
    """
    lines = text.splitlines()
    sections = {}
    current_section = None
    current_text = []

    for line in lines:
        stripped_line = line.strip()
        
        # Special handling for Conclusion section - stop at "We are now ready to take your questions."
        if current_section and current_section.lower() == "conclusion":
            if "we are now ready to take your questions" in stripped_line.lower():
                # Add this line and stop the conclusion section
                current_text.append(line)
                sections[current_section] = '\n'.join(current_text).strip()
                current_section = None
                current_text = []
                continue
        
        # Check if the line is EXACTLY a header (alone on the line)
        if any(stripped_line.lower() == h.lower() for h in headers):
            # Save previous section
            if current_section:
                sections[current_section] = '\n'.join(current_text).strip()
            
            # Start new section
            current_section = stripped_line
            current_text = [stripped_line]
        else:
            # Add text to current section
            if current_section:
                current_text.append(line)

    # Save last section
    if current_section:
        sections[current_section] = '\n'.join(current_text).strip()

    return sections

def process_single_pdf(pdf_file, central_bank, date_format):
    """
    Process a single PDF file
    """
    try:
        print(f"\n{'='*60}")
        print(f"📄 Processing PDF: {os.path.basename(pdf_file)}")
        print(f"{'='*60}")
        
        # Open PDF and extract text
        reader = PdfReader(pdf_file)
        text = ""

        for page in reader.pages:
            text += page.extract_text()

        # Extract date from text (with filename fallback for PRESS CONFERENCE files)
        extracted_date = extract_date_from_text(text, date_format, pdf_file)
        
        # Get unique folder name to avoid overwriting
        base_path = os.path.join("TEXT", central_bank)
        unique_folder_name = get_unique_folder_name(base_path, extracted_date)
        
        # Create subfolder with central bank structure: TEXT\EZB\unique_folder_name or TEXT\FED\unique_folder_name
        date_folder = os.path.join(base_path, unique_folder_name)
        os.makedirs(date_folder, exist_ok=True)
        print(f"📁 Subfolder created: {date_folder}")
        
        # Save complete text as 0_FULL.txt
        txt_file = os.path.join(date_folder, "0_FULL.txt")
        with open(txt_file, "w", encoding="utf-8") as f:
            f.write(text)

        # Define headers
        headers = [
            "Financial and monetary conditions",
            "Inflation", 
            "Economic activity",
            "Risk assessment",
            "Press conference",
            "Conclusion"
        ]
        
        # Extract sections
        sections = extract_sections_precise(text, headers)
        
        # Mapping for filenames with numbering
        section_mapping = {
            "Conclusion": "1_CONCLUSION",
            "Inflation": "2_INFLATION",
            "Economic activity": "3_ECONOMIC_ACTIVITY",
            "Risk assessment": "4_RISK_ASSESSMENT",
            "Press conference": "5_PRESS_CONFERENCE",
            "Financial and monetary conditions": "6_FINANCIAL_MONETARY_CONDITIONS"
        }
        
        extracted_sections = 0
        
        # Save each found section
        for section_title, section_content in sections.items():
            if section_title in section_mapping:
                section_key = section_mapping[section_title]
                # Numbered filename
                section_filename = f"{section_key}.txt"
                section_file = os.path.join(date_folder, section_filename)
                
                with open(section_file, "w", encoding="utf-8") as f:
                    f.write(section_content)
                
                print(f"✅ {section_key} extracted and saved")
                extracted_sections += 1

        print(f"\n✅ PDF successfully processed!")
        print(f"📄 Number of pages: {len(reader.pages)}")
        print(f"📝 Text length: {len(text)} characters")
        print(f"📑 Extracted sections: {extracted_sections}/6")
        print(f"💾 Saved to: {date_folder}")
        
        return True
        
    except Exception as e:
        print(f"❌ Error processing {os.path.basename(pdf_file)}: {e}")
        return False

# Main batch processing
try:
    # Get user choices
    central_bank = get_central_bank_choice()
    date_format = get_date_format_choice()
    
    # Select appropriate PDF folder based on central bank choice
    if central_bank == "EZB":
        pdf_folder = pdf_folder_ezb
    else:
        pdf_folder = pdf_folder_fed
    
    # Find all PDF files in the specified folder
    pdf_pattern = os.path.join(pdf_folder, "*.pdf")
    pdf_files = glob.glob(pdf_pattern)
    
    if not pdf_files:
        print(f"❌ No PDF files found in folder: {pdf_folder}")
        print(f"📁 Make sure the folder exists and contains PDF files")
        exit()
    
    print(f"\n🚀 Starting batch processing...")
    print(f"🏦 Central Bank: {central_bank}")
    print(f"📅 Date Format: {'Month as text' if date_format == 'text' else 'Month as number'}")
    print(f"📂 Folder: {pdf_folder}")
    print(f"📄 Found {len(pdf_files)} PDF files:")
    
    for i, pdf_file in enumerate(pdf_files, 1):
        print(f"   {i}. {os.path.basename(pdf_file)}")
    
    # Process each PDF
    successful = 0
    failed = 0
    
    for pdf_file in pdf_files:
        if process_single_pdf(pdf_file, central_bank, date_format):
            successful += 1
        else:
            failed += 1
    
    # Final summary
    print(f"\n{'='*60}")
    print(f"🎉 BATCH PROCESSING COMPLETE!")
    print(f"{'='*60}")
    print(f"🏦 Central Bank: {central_bank}")
    print(f"📅 Date Format: {'Month as text' if date_format == 'text' else 'Month as number'}")
    print(f"✅ Successfully processed: {successful} PDFs")
    print(f"❌ Failed: {failed} PDFs")
    print(f"📊 Total: {len(pdf_files)} PDFs")
    print(f"💾 All files saved to: TEXT/{central_bank}/")
    
except Exception as e:
    print(f"❌ Critical error during batch processing: {e}")
