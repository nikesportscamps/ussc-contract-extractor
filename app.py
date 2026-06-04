import streamlit as st
import os
import re
import io
import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

# ─────────────────────────────────────────────
# Page config
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="USSC Contract Extractor",
    page_icon="📋",
    layout="centered"
)

# ─────────────────────────────────────────────
# Styling
# ─────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono&display=swap');

    html, body, [class*="css"] {
        font-family: 'DM Sans', sans-serif;
    }

    .main {
        background-color: #f8f9fb;
    }

    .block-container {
        padding-top: 2.5rem;
        max-width: 780px;
    }

    h1 {
        font-size: 1.8rem !important;
        font-weight: 600 !important;
        color: #1a2744 !important;
        margin-bottom: 0.25rem !important;
    }

    .subtitle {
        color: #6b7280;
        font-size: 0.95rem;
        margin-bottom: 2rem;
    }

    .info-box {
        background: #eef2ff;
        border-left: 4px solid #1a2744;
        padding: 0.85rem 1.1rem;
        border-radius: 0 8px 8px 0;
        font-size: 0.88rem;
        color: #374151;
        margin-bottom: 1.5rem;
    }

    .result-box {
        background: #f0fdf4;
        border-left: 4px solid #16a34a;
        padding: 0.85rem 1.1rem;
        border-radius: 0 8px 8px 0;
        font-size: 0.88rem;
        color: #166534;
        margin: 1rem 0;
    }

    .stButton > button {
        background-color: #1a2744;
        color: white;
        border: none;
        border-radius: 8px;
        padding: 0.6rem 1.8rem;
        font-family: 'DM Sans', sans-serif;
        font-weight: 500;
        font-size: 0.95rem;
        cursor: pointer;
        transition: background 0.2s;
        width: 100%;
    }

    .stButton > button:hover {
        background-color: #243660;
    }

    .stDownloadButton > button {
        background-color: #16a34a;
        color: white;
        border: none;
        border-radius: 8px;
        padding: 0.6rem 1.8rem;
        font-family: 'DM Sans', sans-serif;
        font-weight: 500;
        font-size: 0.95rem;
        width: 100%;
    }

    .stDownloadButton > button:hover {
        background-color: #15803d;
    }

    .file-count {
        font-size: 0.85rem;
        color: #6b7280;
        margin-top: 0.4rem;
    }

    footer {visibility: hidden;}
    #MainMenu {visibility: hidden;}
    header {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# Extraction logic
# ─────────────────────────────────────────────

USSC_SIGNERS = [
    "Nora Osei", "Olivia Bowman", "Sarah Hebberd", "Tim Phelan",
    "Terrence Trammell", "Luke Gromer", "William Phelan"
]

def clean(text):
    if not text:
        return ""
    text = re.sub(r'_+', '', text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

def collapse_spaced(text):
    if not text:
        return text
    text = re.sub(r'_', '', text)
    def collapse(m):
        return m.group(0).replace(' ', '')
    text = re.sub(r'(?<!\w)([A-Za-z0-9] )+[A-Za-z0-9](?!\w)', collapse, text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

def extract_all_dates(text):
    collapsed = collapse_spaced(text)
    return re.findall(r'\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}', collapsed)

def extract_header_name(pages):
    if not pages:
        return "NOT FOUND"
    lines = [l.strip() for l in pages[0].split('\n') if l.strip()]
    year_pattern = re.compile(r'^20\d{2}$')
    for i, line in enumerate(lines):
        if year_pattern.match(line):
            if i > 0:
                candidate = lines[i-1]
                if 'DocuSign' not in candidate and 'Envelope' not in candidate:
                    return clean(candidate)
    return clean(lines[0]) if lines else "NOT FOUND"

def extract_expiration_date(full_text):
    m = re.search(r'shall continue until\s+(December 31,\s*20\d{2})', full_text, re.IGNORECASE)
    if m:
        return clean(m.group(1))
    return "NOT FOUND"

def find_signature_page(pages):
    for text in pages:
        if text and 'EXECUTED by the parties' in text:
            return text
    for text in pages:
        if text and 'U.S. Sports Camps' in text and 'Date' in text:
            return text
    return ""

def extract_signature_fields(pages):
    operator_date = "NOT FOUND"
    ussc_name = "NOT FOUND"
    ussc_date = "NOT FOUND"

    sig_text = find_signature_page(pages)
    if not sig_text:
        return operator_date, ussc_name, ussc_date

    collapsed_full = collapse_spaced(sig_text)
    lines = sig_text.split('\n')
    collapsed_lines = [collapse_spaced(l) for l in lines]

    # USSC signer: check known signers first
    for signer in USSC_SIGNERS:
        if signer in collapsed_full:
            ussc_name = signer
            break

    # Fallback: last Name: occurrence
    if ussc_name == "NOT FOUND":
        parts = re.split(r'Name:', collapsed_full)
        if len(parts) >= 2:
            last_part = parts[-1].strip()
            words = last_part.split()
            name_words = []
            for w in words:
                if re.match(r'^[A-Za-zÀ-ÿ]+$', w):
                    name_words.append(w)
                else:
                    break
                if len(name_words) == 2:
                    break
            if name_words:
                ussc_name = ' '.join(name_words)

    # Dates
    for collapsed in collapsed_lines:
        if re.search(r'[Dd]ate:', collapsed):
            dates = extract_all_dates(collapsed)
            if len(dates) >= 2:
                if operator_date == "NOT FOUND":
                    operator_date = dates[0]
                if ussc_date == "NOT FOUND":
                    ussc_date = dates[1]
            elif len(dates) == 1:
                if operator_date == "NOT FOUND":
                    operator_date = dates[0]
                elif ussc_date == "NOT FOUND":
                    ussc_date = dates[0]
        elif operator_date == "NOT FOUND" or ussc_date == "NOT FOUND":
            dates = extract_all_dates(collapsed)
            if len(dates) >= 2:
                if operator_date == "NOT FOUND":
                    operator_date = dates[0]
                if ussc_date == "NOT FOUND":
                    ussc_date = dates[1]
            elif len(dates) == 1:
                if operator_date == "NOT FOUND":
                    operator_date = dates[0]
                elif ussc_date == "NOT FOUND":
                    ussc_date = dates[0]

    return operator_date, ussc_name, ussc_date

def process_pdf_bytes(file_bytes, filename):
    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            pages_text = [p.extract_text() or "" for p in pdf.pages]
            full_text = '\n'.join(pages_text)
            op_date, ussc_name, ussc_date = extract_signature_fields(pages_text)
            return {
                "File": filename,
                "Director / Entity Name": extract_header_name(pages_text),
                "Date Operator Signed": op_date,
                "USSC Signer Name": ussc_name,
                "Date USSC Signed": ussc_date,
                "Contract Expiration Date": extract_expiration_date(full_text),
            }
    except Exception as e:
        return {
            "File": filename,
            "Director / Entity Name": f"ERROR: {e}",
            "Date Operator Signed": "",
            "USSC Signer Name": "",
            "Date USSC Signed": "",
            "Contract Expiration Date": "",
        }

def build_excel(data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Contract Data"
    headers = [
        "File", "Director / Entity Name", "Date Operator Signed",
        "USSC Signer Name", "Date USSC Signed", "Contract Expiration Date"
    ]
    header_fill = PatternFill("solid", fgColor="1F3864")
    header_font = Font(color="FFFFFF", bold=True)
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for row_num, record in enumerate(data, 2):
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=row_num, column=col_num, value=record.get(header, ""))
            cell.alignment = Alignment(vertical="center")
            if row_num % 2 == 0:
                cell.fill = PatternFill("solid", fgColor="EEF2F7")
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            try:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = min(max_len + 4, 50)
    ws.freeze_panes = "A2"
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ─────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────

st.markdown("# 📋 USSC Contract Extractor")
st.markdown('<p class="subtitle">Upload your Nike Sports Camp operator agreements and download a clean Excel summary in seconds.</p>', unsafe_allow_html=True)

st.markdown("""
<div class="info-box">
    <strong>What this extracts:</strong> Director / Entity Name &nbsp;·&nbsp; Date Operator Signed &nbsp;·&nbsp; 
    USSC Signer Name &nbsp;·&nbsp; Date USSC Signed &nbsp;·&nbsp; Contract Expiration Date
    <br><br>
    <strong>Supports:</strong> All contract years (2019–2026), Operator Agreements and Independent Contractor Agreements
</div>
""", unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    "Upload contract PDFs",
    type=["pdf"],
    accept_multiple_files=True,
    label_visibility="collapsed"
)

if uploaded_files:
    st.markdown(f'<p class="file-count">✅ {len(uploaded_files)} file{"s" if len(uploaded_files) != 1 else ""} selected</p>', unsafe_allow_html=True)

    if st.button("Extract Contract Data"):
        results = []
        progress = st.progress(0)
        status = st.empty()

        for i, uploaded_file in enumerate(uploaded_files):
            status.markdown(f'<p class="file-count">Processing {i+1} of {len(uploaded_files)}: {uploaded_file.name}</p>', unsafe_allow_html=True)
            file_bytes = uploaded_file.read()
            result = process_pdf_bytes(file_bytes, uploaded_file.name)
            results.append(result)
            progress.progress((i + 1) / len(uploaded_files))

        status.empty()
        progress.empty()

        excel_bytes = build_excel(results)

        st.markdown(f"""
        <div class="result-box">
            🎉 Done! <strong>{len(results)} contract{"s" if len(results) != 1 else ""}</strong> processed successfully.
        </div>
        """, unsafe_allow_html=True)

        st.download_button(
            label="⬇️ Download Contract_Data.xlsx",
            data=excel_bytes,
            file_name="Contract_Data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.markdown("""
    <div style="border: 2px dashed #d1d5db; border-radius: 12px; padding: 2.5rem; text-align: center; color: #9ca3af; font-size: 0.9rem; margin-top: 0.5rem;">
        Drag and drop your PDF contracts here<br>or click <strong>Browse files</strong> above
    </div>
    """, unsafe_allow_html=True)

st.markdown("---")
st.markdown('<p style="text-align:center; color:#9ca3af; font-size:0.8rem;">US Sports Camps · Nike Sports Camps · Contract Data Tool</p>', unsafe_allow_html=True)
