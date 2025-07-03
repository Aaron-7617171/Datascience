from pypdf import PdfReader
import os
import re

# ============================================
# MANUAL PDF NAME INPUT
# ============================================
pdf_name = "PRESS CONFERENCE_6_March.pdf"  # Change this to your PDF filename
# ============================================

# .py program must be in the same folder as "PDF" and "TEXT" folders

# Change to the folder of the .py file
script_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(script_dir)
print(f"📁 Working directory changed to: {script_dir}")

def extract_date_from_text(text):
    """
    Extracts the date that appears after "Combined monetary policy decisions and statement"
    """
    # Search for the specific text and the date after it
    pattern = r'Combined monetary policy decisions and\s*statement\s*(\d{1,2}\s+\w+\s+\d{4})'
    
    match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
    if match:
        date_str = match.group(1).strip()
        print(f"📅 Found date after 'Combined monetary policy decisions and statement': {date_str}")
        
        # Format date for filenames
        # "17 April 2025" -> "17_April_2025"
        formatted_date = date_str.replace(" ", "_")
        return formatted_date
    
    # Fallback if no date found
    print("⚠️ No date found after 'Combined monetary policy decisions and statement', using default: no date found")
    return "no date found"

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

try:
    # PDF filename using the variable
    pdf_file = os.path.join("PDF", pdf_name)
    print(f"📄 Processing PDF: {pdf_name}")
    
    # Open PDF and extract text
    reader = PdfReader(pdf_file)
    text = ""

    for page in reader.pages:
        text += page.extract_text()

    # Extract date from text (only after the specific text)
    extracted_date = extract_date_from_text(text)
    
    # Get unique folder name to avoid overwriting
    base_path = os.path.join("TEXT", "EZB")
    unique_folder_name = get_unique_folder_name(base_path, extracted_date)
    
    # Create subfolder with EZB structure: TEXT\EZB\unique_folder_name
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
            
            print(f"✅ {section_key} extracted and saved in: {date_folder}/{section_filename}")
            extracted_sections += 1

    print(f"\n✅ Text successfully extracted and saved in: {date_folder}/0_FULL.txt")
    print(f"📄 Number of pages: {len(reader.pages)}")
    print(f"📝 Text length: {len(text)} characters")
    print(f"📑 Extracted sections: {extracted_sections}/6")
    
except Exception as e:
    print(f"❌ Error processing PDF: {e}")
