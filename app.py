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

# --- PAGE SETUP ---
st.set_page_config(page_title="Auto-Detect Bot", layout="wide")
st.title("⚡ Dynamic Excel Auto-Filler")
st.markdown("Bas Excel aur Files daalo. Script columns padh kar khud data dhoond legi.")

# --- SMART PATTERN GENERATOR ---
def generate_smart_regex(header_name):
    """
    Excel header se Regex banata hai.
    Example: "Invoice Date" -> r"Invoice[\s_\-]*Date\s*[:\-\=]\s*(.*)"
    Ye spaces, underscore, aur colon/hyphen sab handle karega.
    """
    # Special characters hatake safe string banao
    clean_header = re.escape(header_name)
    
    # Space ko flexible banao (Space ya Underscore chalega)
    flexible_header = clean_header.replace(r"\ ", r"[\s_\-]*")
    
    # Final Pattern: Header + Separator (: or - or =) + Value
    # Group 1 mein value capture hogi
    return rf"{flexible_header}\s*[:\-\=]\s*(.*)"

# --- READERS (Format Handlers) ---
def get_pdf_text(path):
    try:
        reader = PdfReader(path)
        return "".join([p.extract_text() or "" for p in reader.pages])
    except: return ""

def get_word_text(path):
    try:
        return "\n".join([p.text for p in docx.Document(path).paragraphs])
    except: return ""

def get_msg_text(path):
    try:
        msg = extract_msg.Message(path)
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

# --- MAIN LOGIC ---
def run_dynamic_automation(excel_file, zip_file, id_col, target_cols):
    # 1. Temp Folder
    temp_dir = "temp_dynamic_docs"
    if os.path.exists(temp_dir): shutil.rmtree(temp_dir)
    os.makedirs(temp_dir)

    with zipfile.ZipFile(zip_file, 'r') as z:
        z.extractall(temp_dir)

    all_files = []
    for root, _, files in os.walk(temp_dir):
        for f in files:
            all_files.append(os.path.join(root, f))

    # 2. Excel Load
    df = pd.read_excel(excel_file)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    
    wb = load_workbook(output)
    ws = wb.active
    yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    no_fill = PatternFill(fill_type=None)

    # Headers Map
    headers = [str(c.value).strip() for c in ws[1]]
    
    # ID Column Index
    try:
        id_idx = headers.index(id_col)
    except ValueError:
        st.error(f"ID Column '{id_col}' Excel mein nahi mila!")
        return None

    # 3. Processing
    bar = st.progress(0)
    status = st.empty()

    for i, row in df.iterrows():
        bar.progress((i + 1) / len(df))
        
        # Row ID
        search_key = str(row.get(id_col, "")).strip().lower()
        if not search_key or search_key == "nan": continue

        status.text(f"Scanning for: {search_key}...")

        # Find File
        file_text = None
        for f_path in all_files:
            if search_key in os.path.basename(f_path).lower():
                file_text = extract_text_smart(f_path)
                break
        
        # Fill Target Columns
        if file_text:
            for col_name in target_cols:
                if col_name in headers:
                    col_idx = headers.index(col_name) + 1
                    cell = ws.cell(row=i+2, column=col_idx)
                    
                    # --- DYNAMIC REGEX MAGIC HERE ---
                    # Column name se pattern banao
                    dynamic_pattern = generate_smart_regex(col_name)
                    
                    # Search
                    # re.IGNORECASE se "po number", "PO NUMBER" sab match hoga
                    match = re.search(dynamic_pattern, file_text, re.IGNORECASE)
                    
                    if match:
                        # Value mil gayi -> Clean & Fill
                        # .split('\n')[0] isliye taaki sirf pehli line uthaye (agar multiline ho to)
                        value = match.group(1).strip().split('\n')[0]
                        cell.value = value
                        cell.fill = no_fill
                    else:
                        # File mili par ye wala column ka data nahi mila -> Yellow
                        if not cell.value: # Agar pehle se bhara hai to mat chedo
                            cell.fill = yellow
        else:
            # File hi nahi mili -> Selected columns yellow kardo
            for col_name in target_cols:
                if col_name in headers:
                    col_idx = headers.index(col_name) + 1
                    ws.cell(row=i+2, column=col_idx).fill = yellow

    # Cleanup
    shutil.rmtree(temp_dir)
    final_out = BytesIO()
    wb.save(final_out)
    final_out.seek(0)
    return final_out

# --- UI INTERFACE ---
col1, col2 = st.columns(2)
f_excel = col1.file_uploader("1. Excel Upload", type=["xlsx"])
f_zip = col2.file_uploader("2. Zip Upload", type=["zip"])

if f_excel and f_zip:
    # --- STEP 1: READ HEADERS ---
    df_preview = pd.read_excel(f_excel)
    all_columns = df_preview.columns.tolist()
    
    st.write("---")
    st.subheader("🛠️ Auto-Configuration")
    
    c1, c2 = st.columns(2)
    # User se pucho ID column kaunsa hai
    selected_id = c1.selectbox("Wo column chuno jisse file match hogi (Unique ID):", all_columns, index=0)
    
    # User se pucho kaunse columns AUTOMATICALLY bharne hain
    # ID column ko by default hata diya list se
    default_targets = [c for c in all_columns if c != selected_id]
    selected_targets = c2.multiselect("Kaunse Columns automatic bharne hain?", all_columns, default=default_targets)

    if st.button("🚀 Start Magic"):
        if not selected_targets:
            st.error("Kam se kam ek column to select karo bharne ke liye!")
        else:
            with st.spinner("Bot Excel ke columns padh raha hai aur data dhoond raha hai..."):
                # File pointer reset for processing
                f_excel.seek(0)
                f_zip.seek(0)
                
                result = run_dynamic_automation(f_excel, f_zip, selected_id, selected_targets)
                
                if result:
                    st.success("Done! ✅")
                    st.download_button("📥 Download Result", result, "Auto_Filled.xlsx")
