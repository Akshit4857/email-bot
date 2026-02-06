import streamlit as st
import pandas as pd
import zipfile
import os
import re
import shutil
import docx
from pypdf import PdfReader
from io import BytesIO
import extract_msg
from email import policy
from email.parser import BytesParser
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Change Log Automator", layout="wide")
st.title("🛡️ IT Change Request Automator")
st.markdown("Upload **Excel Tracker** + **Mail/Doc Zip**. Script will fill multiple columns based on rules.")

# ==========================================
# ⚙️ CONFIGURATION: MAPPING RULES (Yahan Dhyan Dein)
# ==========================================
# Left side: Excel ka Column Name (Exact spelling)
# Right side: Email/Doc mein dhoondne wala Regex Pattern
# Aap is list ko badha sakte hain

EXTRACTION_RULES = {
    "Change ID": r"(?:Change\s*ID|CR\s*No|Ticket)\s*[:\-]?\s*([A-Za-z0-9\-]+)",
    "Application": r"Application\s*[:\-]?\s*(.*)",
    "Change description": r"Description\s*[:\-]?\s*(.*)",
    "Change Type": r"Type\s*[:\-]?\s*(Normal|Emergency|Standard)",
    "Requested by": r"Requested\s*by\s*[:\-]?\s*([A-Za-z\s]+)",
    "Date of Approval": r"Approval\s*Date\s*[:\-]?\s*(\d{2}[-/\.]\d{2}[-/\.]\d{4})",
    "Release ID": r"Release\s*ID\s*[:\-]?\s*(.*)",
    "Developer": r"Developer\s*[:\-]?\s*([A-Za-z\s]+)",
    # Aur columns yahan add karein...
}

# Excel mein kaunsa column Unique Key hai? (Folder mein file dhoondne ke liye)
MATCH_ID_COLUMN = "Change ID" 

# ==========================================
# READERS (DO NOT TOUCH)
# ==========================================

def get_pdf_text(path):
    try:
        reader = PdfReader(path)
        text = ""
        for page in reader.pages:
            text += page.extract_text() or ""
        return text
    except: return ""

def get_word_text(path):
    try:
        doc = docx.Document(path)
        return "\n".join([p.text for p in doc.paragraphs])
    except: return ""

def get_msg_text(path):
    try:
        msg = extract_msg.Message(path)
        # Subject + Body dono combine kar rahe hain taaki data miss na ho
        return f"{msg.subject}\n{msg.body}"
    except: return ""

def get_eml_text(path):
    try:
        with open(path, 'rb') as f:
            msg = BytesParser(policy=policy.default).parse(f)
            return f"{msg['subject']}\n{msg.get_body(preferencelist=('plain')).get_content()}"
    except: return ""

def extract_text_smart(file_path):
    ext = file_path.lower()
    if ext.endswith('.pdf'): return get_pdf_text(file_path)
    elif ext.endswith('.docx'): return get_word_text(file_path)
    elif ext.endswith('.msg'): return get_msg_text(file_path)
    elif ext.endswith('.eml'): return get_eml_text(file_path)
    return None

# ==========================================
# MAIN LOGIC
# ==========================================

def run_automation(excel_file, zip_file):
    # 1. Temp Setup
    temp_dir = "temp_audit_docs"
    if os.path.exists(temp_dir): shutil.rmtree(temp_dir)
    os.makedirs(temp_dir)

    with zipfile.ZipFile(zip_file, 'r') as z:
        z.extractall(temp_dir)

    all_files = []
    for root, _, files in os.walk(temp_dir):
        for f in files:
            all_files.append(os.path.join(root, f))

    # 2. Excel Setup
    df = pd.read_excel(excel_file)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    
    wb = load_workbook(output)
    ws = wb.active
    
    # Styles
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    no_fill = PatternFill(fill_type=None)

    # Headers Map (Column Name -> Index)
    headers = [str(c.value).strip() for c in ws[1]]
    
    # Validate ID Column
    if MATCH_ID_COLUMN not in headers:
        st.error(f"❌ '{MATCH_ID_COLUMN}' column Excel mein nahi mila. Config check karein.")
        return None

    id_col_idx = headers.index(MATCH_ID_COLUMN)

    # 3. Processing Rows
    progress_bar = st.progress(0)
    status_text = st.empty()

    for i, row in df.iterrows():
        progress_bar.progress((i + 1) / len(df))
        
        # Row ID fetch karein
        row_id = row.get(MATCH_ID_COLUMN) # Use config column
        search_key = str(row_id).strip().lower()

        if not search_key or search_key == "nan":
            continue
        
        status_text.text(f"Processing ID: {search_key}")

        # Document dhundna (Match ID in Filename)
        doc_text = None
        for f_path in all_files:
            if search_key in os.path.basename(f_path).lower():
                doc_text = extract_text_smart(f_path)
                break 
        
        # --- MULTI-COLUMN FILLING MAGIC ---
        if doc_text:
            # Document mil gaya, ab rules check karo
            for col_name, pattern in EXTRACTION_RULES.items():
                if col_name in headers:
                    target_idx = headers.index(col_name) + 1
                    cell = ws.cell(row=i+2, column=target_idx)
                    
                    # Agar cell pehle se khali hai tabhi bharein (Overwrite protection)
                    # Remove 'if not cell.value:' below if you want to overwrite always
                    if not cell.value: 
                        match = re.search(pattern, doc_text, re.IGNORECASE)
                        if match:
                            cell.value = match.group(1).strip()
                            cell.fill = no_fill
                        else:
                            # Rule hai par data nahi mila -> Yellow
                            cell.fill = yellow_fill
        else:
            # Document hi nahi mila -> Saare target columns Yellow
            for col_name in EXTRACTION_RULES.keys():
                if col_name in headers:
                    target_idx = headers.index(col_name) + 1
                    cell = ws.cell(row=i+2, column=target_idx)
                    cell.fill = yellow_fill

    # Cleanup
    shutil.rmtree(temp_dir)
    final_out = BytesIO()
    wb.save(final_out)
    final_out.seek(0)
    return final_out

# ==========================================
# UI
# ==========================================
col1, col2 = st.columns(2)
uploaded_excel = col1.file_uploader("1. Excel File", type=["xlsx"])
uploaded_zip = col2.file_uploader("2. Docs/Emails Zip", type=["zip"])

if st.button("🚀 Start Extraction") and uploaded_excel and uploaded_zip:
    with st.spinner("Analyzing email trails..."):
        result = run_automation(uploaded_excel, uploaded_zip)
        if result:
            st.success("Extraction Complete!")
            st.download_button("📥 Download Final Excel", result, "Updated_Tracker.xlsx")
