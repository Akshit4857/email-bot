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
st.set_page_config(page_title="Smart Doc Processor", layout="wide")
st.title("🤖 Universal Document Data Extractor")
st.markdown("""
**Supported Formats:** `.msg`, `.eml`, `.pdf`, `.docx`  
Upload karo Excel + Zip (docs ka), aur magic dekho! ✨
""")

# --- SIDEBAR SETTINGS ---
with st.sidebar:
    st.header("⚙️ Settings")
    id_col = st.text_input("Match ID Column (Excel)", value="Ticket_ID")
    target_col = st.text_input("Target Column to Fill", value="PO_Number")
    regex_input = st.text_input("Regex Pattern", value=r"PO\s*[:\-]?\s*(\d{10})")
    st.info("💡 **Regex Tip:** Agar pattern strict chahiye to hi change karein.")

# --- HELPER FUNCTIONS (READERS) ---

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
        return msg.body
    except: return ""

def get_eml_text(path):
    try:
        with open(path, 'rb') as f:
            msg = BytesParser(policy=policy.default).parse(f)
            return msg.get_body(preferencelist=('plain')).get_content()
    except: return ""

def extract_text_smart(file_path):
    """File extension check karke sahi tool use karta hai"""
    ext = file_path.lower()
    if ext.endswith('.pdf'): return get_pdf_text(file_path)
    elif ext.endswith('.docx'): return get_word_text(file_path)
    elif ext.endswith('.msg'): return get_msg_text(file_path)
    elif ext.endswith('.eml'): return get_eml_text(file_path)
    return None

# --- MAIN PROCESSING LOGIC ---

def run_automation(excel_file, zip_file):
    # 1. Temp Folder Setup
    temp_dir = "temp_docs"
    if os.path.exists(temp_dir): shutil.rmtree(temp_dir)
    os.makedirs(temp_dir)

    # 2. Unzip Files
    with zipfile.ZipFile(zip_file, 'r') as z:
        z.extractall(temp_dir)

    # List all files (recursively, agar folder ke andar folder ho)
    all_files = []
    for root, _, files in os.walk(temp_dir):
        for f in files:
            all_files.append(os.path.join(root, f))

    # 3. Load Excel Logic
    df = pd.read_excel(excel_file)
    
    # Create Output Buffer
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    
    wb = load_workbook(output)
    ws = wb.active
    yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    no_fill = PatternFill(fill_type=None)

    # Find Columns
    headers = [c.value for c in ws[1]]
    try:
        target_idx = headers.index(target_col) + 1
        # ID column map karne k liye pandas use krre h
    except ValueError:
        st.error(f"❌ Target Column '{target_col}' Excel mein nahi mila!")
        return None

    # Progress Bar
    bar = st.progress(0)
    status = st.empty()

    # 4. Loop Logic
    for i, row in df.iterrows():
        bar.progress((i + 1) / len(df))
        
        search_id = str(row.get(id_col, "")).strip().lower()
        if not search_id or search_id == "nan": continue

        status.text(f"Searching for: {search_id}...")
        
        # File Match Logic (Filename contains ID)
        matched_text = None
        for f_path in all_files:
            if search_id in os.path.basename(f_path).lower():
                matched_text = extract_text_smart(f_path)
                break # Pehli match milte hi ruk jao (Fast)
        
        # Regex Extraction
        final_val = None
        if matched_text:
            match = re.search(regex_input, matched_text, re.IGNORECASE)
            if match:
                final_val = match.group(1).strip()
        
        # Update Excel Cell
        cell = ws.cell(row=i+2, column=target_idx)
        if final_val:
            cell.value = final_val
            cell.fill = no_fill
        else:
            cell.value = None
            cell.fill = yellow # STRICTLY YELLOW

    # Cleanup
    shutil.rmtree(temp_dir)
    
    # Save Final File
    final_out = BytesIO()
    wb.save(final_out)
    final_out.seek(0)
    return final_out

# --- UI EXECUTION ---
col1, col2 = st.columns(2)
upl_excel = col1.file_uploader("1. Excel File Upload karein", type=["xlsx"])
upl_zip = col2.file_uploader("2. Sare Docs ka ZIP upload karein", type=["zip"])

if st.button("🚀 Start Processing") and upl_excel and upl_zip:
    with st.spinner("Processing... Dhyan rahe ye 'Strict Mode' hai."):
        result = run_automation(upl_excel, upl_zip)
        if result:
            st.success("Ho gaya! Niche se file download karein. 🎉")
            st.download_button("📥 Download Processed Excel", result, "Final_Result.xlsx")