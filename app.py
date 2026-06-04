import streamlit as st
import re
import io
import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

st.set_page_config(
    page_title="USSC Operator Agreement Extractor",
    page_icon="📋",
    layout="centered"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono&display=swap');
    html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
    .main { background-color: #f8f9fb; }
    .block-container { padding-top: 2.5rem; max-width: 780px; }
    h1 { font-size: 1.8rem !important; font-weight: 600 !important; color: #1a2744 !important; margin-bottom: 0.25rem !important; }
    .subtitle { color: #6b7280; font-size: 0.95rem; margin-bottom: 2rem; }
    .info-box { background: #eef2ff; border-left: 4px solid #1a2744; padding: 0.85rem 1.1rem; border-radius: 0 8px 8px 0; font-size: 0.88rem; color: #374151; margin-bottom: 1.5rem; }
    .result-box { background: #f0fdf4; border-left: 4px solid #16a34a; padding: 0.85rem 1.1rem; border-radius: 0 8px 8px 0; font-size: 0.88rem; color: #166534; margin: 1rem 0; }
    .stButton > button { background-color: #1a2744; color: white; border: none; border-radius: 8px; padding: 0.6rem 1.8rem; font-family: 'DM Sans', sans-serif; font-weight: 500; font-size: 0.95rem; width: 100%; }
    .stButton > button:hover { background-color: #243660; }
    .stDownloadButton > button { background-color: #16a34a; color: white; border: none; border-radius: 8px; padding: 0.6rem 1.8rem; font-family: 'DM Sans', sans-serif; font-weight: 500; font-size: 0.95rem; width: 100%; }
    .stDownloadButton > button:hover { background-color: #15803d; }
    .file-count { font-size: 0.85rem; color: #6b7280; margin-top: 0.4rem; }
    footer {visibility: hidden;} #MainMenu {visibility: hidden;} header {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# Known USSC signers — matched by email too
# ─────────────────────────────────────────────
USSC_SIGNERS = [
    ("Nora Osei", ["nosei", "nora.osei", "noraosei"]),
    ("Olivia Bowman", ["obowman", "olivia.bowman", "oliviabowman"]),
    ("Sarah Hebberd", ["shebberd", "sarah.hebberd"]),
    ("Tim Phelan", ["tphelan", "tim.phelan"]),
    ("Terrence Trammell", ["ttrammell", "terrence.trammell"]),
    ("Luke Gromer", ["lgromer", "luke.gromer"]),
    ("William Phelan", ["wphelan", "william.phelan"]),
    ("Andrew Harrington", ["aharrington", "andrew.harrington"]),
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

def collapse_newline_spaced(text):
    """
    Fix names broken across newlines like:
    'O\n_\nl\n_\ni\n_\nv\n_\ni\n_\na' -> 'Olivia'
    """
    # Remove underscores
    text = re.sub(r'_', '', text)
    # Collapse single chars separated by newlines: "O\nl\ni\nv\ni\na" -> "Olivia"
    text = re.sub(r'(?<!\w)([A-Za-z]\n)+[A-Za-z](?!\w)', lambda m: m.group(0).replace('\n', ''), text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

def extract_all_dates(text):
    collapsed = collapse_spaced(text)
    candidates = re.findall(r'\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}', collapsed)
    valid = []
    for c in candidates:
        parts = re.split(r'[\/\-]', c)
        if len(parts) == 3:
            try:
                month = int(parts[0])
                day = int(parts[1])
                year_raw = int(parts[2])
                year = year_raw if year_raw > 100 else 2000 + year_raw
                if 1 <= month <= 12 and 1 <= day <= 31 and 2015 <= year <= 2035:
                    valid.append(c)
            except:
                pass
    return valid

def detect_contract_type(pages_text, full_text):
    insurance_signals = [
        "Commercial General Liability",
        "additional insureds",
        "employer's liability insurance",
        "business automobile liability"
    ]
    insurance_count = sum(1 for s in insurance_signals if s.lower() in full_text.lower())
    if insurance_count >= 2:
        return "Entity"

    entity_signals = [
        '"Entity"', '("Entity")', 'collectively, "Operator"', 'collectively "Operator"',
    ]
    has_entity_keyword = any(s.lower() in full_text[:3000].lower() for s in entity_signals)

    sig_text = find_signature_page(pages_text)
    ussc_pos = sig_text.find('U.S. Sports Camps') if sig_text else -1
    operator_block = sig_text[:ussc_pos] if ussc_pos > 0 else sig_text
    has_llc_in_sig = bool(re.search(r'\bLLC\b|\bInc\b|\bCorp\b', operator_block, re.IGNORECASE))

    if has_entity_keyword or has_llc_in_sig:
        return "Individual with Entity Payment"

    return "Individual"

def extract_header_name(pages):
    if not pages:
        return "NOT FOUND"
    lines = [l.strip() for l in pages[0].split('\n') if l.strip()]
    year_pattern = re.compile(r'^20\d{2}$')
    for i, line in enumerate(lines):
        if year_pattern.match(line):
            if i > 0:
                candidate = lines[i-1]
                if not any(x in candidate for x in ['DocuSign', 'Envelope', 'Docusign']):
                    return clean(candidate)
    for line in lines:
        if not any(x in line for x in ['DocuSign', 'Envelope', 'Docusign']):
            return clean(line)
    return "NOT FOUND"

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

def find_ussc_signer(sig_text):
    """
    Try multiple strategies to find the USSC signer name:
    1. Match known email prefixes (most reliable)
    2. Match known names in collapsed text
    3. Match known names in newline-collapsed text
    4. Fallback: last Name: occurrence
    """
    # Strategy 1: email prefix match (most reliable — email is never broken up)
    for name, email_hints in USSC_SIGNERS:
        for hint in email_hints:
            if hint.lower() in sig_text.lower():
                return name

    # Strategy 2: known name in collapsed text
    collapsed = collapse_spaced(sig_text)
    for name, _ in USSC_SIGNERS:
        if name in collapsed:
            return name

    # Strategy 3: known name after fixing newline-broken chars
    newline_fixed = collapse_newline_spaced(sig_text)
    for name, _ in USSC_SIGNERS:
        if name in newline_fixed:
            return name

    # Strategy 4: last Name: occurrence in collapsed text
    parts = re.split(r'Name:', collapsed)
    if len(parts) >= 2:
        last_part = parts[-1].strip()
        words = last_part.split()
        name_words = []
        for w in words:
            if re.match(r'^[A-Za-zÀ-ÿ\-]+$', w):
                name_words.append(w)
            else:
                break
            if len(name_words) == 2:
                break
        if name_words:
            candidate = ' '.join(name_words)
            skip = ['Director', 'Senior', 'Growth', 'Running', 'Sports',
                    'Camps', 'Owner', 'Title', 'Address', 'Email']
            if candidate not in skip and not any(s == candidate for s in skip):
                return candidate

    return "NOT FOUND"

def extract_signature_fields(pages):
    operator_date = "NOT FOUND"
    ussc_date = "NOT FOUND"

    sig_text = find_signature_page(pages)
    if not sig_text:
        return "NOT FOUND", "NOT FOUND", "NOT FOUND"

    ussc_name = find_ussc_signer(sig_text)

    collapsed_lines = [collapse_spaced(l) for l in sig_text.split('\n')]

    for collapsed in collapsed_lines:
        if re.search(r'\bDate:', collapsed):
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
            contract_type = detect_contract_type(pages_text, full_text)
            return {
                "File": filename,
                "Director / Entity Name": extract_header_name(pages_text),
                "Contract Type": contract_type,
                "Date Operator Signed": op_date,
                "USSC Signer Name": ussc_name,
                "Date USSC Signed": ussc_date,
                "Contract Expiration Date": extract_expiration_date(full_text),
            }
    except Exception as e:
        return {
            "File": filename,
            "Director / Entity Name": f"ERROR: {e}",
            "Contract Type": "",
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
        "File", "Director / Entity Name", "Contract Type",
        "Date Operator Signed", "USSC Signer Name",
        "Date USSC Signed", "Contract Expiration Date"
    ]
    header_fill = PatternFill("solid", fgColor="1F3864")
    header_font = Font(color="FFFFFF", bold=True)
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    type_colors = {
        "Entity": "FFF3CD",
        "Individual with Entity Payment": "D1ECF1",
        "Individual": "D4EDDA",
    }

    for row_num, record in enumerate(data, 2):
        contract_type = record.get("Contract Type", "")
        row_color = type_colors.get(contract_type, "FFFFFF")
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=row_num, column=col_num, value=record.get(header, ""))
            cell.alignment = Alignment(vertical="center")
            cell.fill = PatternFill("solid", fgColor=row_color)

    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            try:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            except:
                pass
        ws.column_dimensions[col_letter].width = min(max_len + 4, 55)

    ws.freeze_panes = "A2"
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ─────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────
st.markdown("# 📋 USSC Operator Agreement Extractor")
st.markdown('<p class="subtitle">Upload up to 100 US Sports Camps Operator Agreements to pull the pertinent information from each contract in a simple, downloadable Excel sheet. It\'ll be done before you can say "supercalifragilisticexpialidocious" 20 times ;)</p>', unsafe_allow_html=True)

st.markdown("""
<div class="info-box">
    <strong>What this extracts:</strong> Director / Entity Name &nbsp;·&nbsp; Contract Type &nbsp;·&nbsp;
    Date Operator Signed &nbsp;·&nbsp; USSC Signer Name &nbsp;·&nbsp; Date USSC Signed &nbsp;·&nbsp; Contract Expiration Date
    <br><br>
    <strong>Supports:</strong> All contract years (2019–2026) &nbsp;·&nbsp; Individual &nbsp;·&nbsp; Entity &nbsp;·&nbsp; Individual with Entity Payment
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
            <br><small>Rows are color-coded by contract type: 🟡 Entity &nbsp;·&nbsp; 🔵 Individual with Entity Payment &nbsp;·&nbsp; 🟢 Individual</small>
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
st.markdown('<p style="text-align:center; color:#9ca3af; font-size:0.8rem;">Created By: Dr. Nora Osei &nbsp;·&nbsp; 2026 &nbsp;·&nbsp; US Sports Camps &nbsp;·&nbsp; Nike Sports Camps</p>', unsafe_allow_html=True)
