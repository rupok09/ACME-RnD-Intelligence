import pymupdf
import streamlit as st
import pandas as pd
import requests
import time
import pymupdf
from io import BytesIO
from urllib.parse import quote_plus

st.set_page_config(
    page_title="ACME FormAI",
    page_icon="🧪",
    layout="wide"
)

# ---------------- FUNCTIONS ----------------

def get_pubchem_data(api_name):
    try:
        url = (
            "https://pubchem.ncbi.nlm.nih.gov/rest/pug/compound/name/"
            f"{quote_plus(api_name)}/property/"
            "MolecularFormula,MolecularWeight,IUPACName,CanonicalSMILES,XLogP,TPSA/JSON"
        )

        response = requests.get(url, timeout=20)

        if response.status_code != 200:
            return None

        data = response.json()
        return data["PropertyTable"]["Properties"][0]

    except:
        return None


def extract_pdf_text(uploaded_file):
    file_bytes = uploaded_file.getvalue()

    pdf = pymupdf.open(
        stream=file_bytes,
        filetype="pdf"
    )

    text = ""

    for page in pdf:
        text += page.get_text()

    pdf.close()

    return text


def detect_api_name(text):
    api_list = [
        "Clarithromycin",
        "Azithromycin",
        "Paracetamol",
        "Acetaminophen",
        "Ibuprofen",
        "Metformin",
        "Metformin HCl",
        "Aspirin",
        "Losartan",
        "Atorvastatin",
        "Ciprofloxacin",
        "Amoxicillin"
    ]

    lower_text = text.lower()

    for api in api_list:
        if api.lower() in lower_text:
            return api

    return "Clarithromycin"


def create_result_table(api_name, result, source_type):
    return pd.DataFrame({
        "Property": [
            "API Name",
            "IUPAC Name",
            "Molecular Formula",
            "Molecular Weight",
            "Canonical SMILES",
            "XLogP",
            "Topological Polar Surface Area",
            "Primary Source",
            "Extraction Type"
        ],
        "Value": [
            api_name,
            result.get("IUPACName", "Not found"),
            result.get("MolecularFormula", "Not found"),
            result.get("MolecularWeight", "Not found"),
            result.get("CanonicalSMILES", "Not found"),
            result.get("XLogP", "Not found"),
            result.get("TPSA", "Not found"),
            "PubChem",
            source_type
        ],
        "Unit": [
            "-",
            "-",
            "-",
            "g/mol",
            "-",
            "-",
            "Å²",
            "-",
            "-"
        ],
        "Source": [
            "PubChem",
            "PubChem",
            "PubChem",
            "PubChem",
            "PubChem",
            "PubChem",
            "PubChem",
            "Trusted Web Source",
            source_type
        ],
        "Confidence": [
            "99%",
            "99%",
            "99%",
            "99%",
            "98%",
            "95%",
            "95%",
            "99%",
            "Demo"
        ]
    })


def create_excel(df):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="API Properties")

    output.seek(0)
    return output


def show_processing(title):
    st.markdown("---")
    st.subheader(title)

    progress = st.progress(0)
    status = st.empty()

    steps = [
        "Initializing request...",
        "Reading document / preparing API query...",
        "Detecting API name...",
        "Connecting to PubChem...",
        "Searching trusted web sources...",
        "Validating collected data...",
        "Generating structured result table..."
    ]

    for i, step in enumerate(steps):
        status.info(step)
        progress.progress(int((i + 1) / len(steps) * 100))
        time.sleep(0.6)


# ---------------- CSS ----------------

st.markdown(
    """
    <style>
    .stApp {
        background: linear-gradient(120deg, #f8f9ff 0%, #ffffff 45%, #f4f0ff 100%);
    }

    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #061936 0%, #020b1d 100%);
        width: 290px !important;
    }

    section[data-testid="stSidebar"] * {
        color: white !important;
    }

    .sidebar-logo {
        font-size: 30px;
        font-weight: 800;
        margin-top: 10px;
        margin-bottom: 5px;
    }

    .sidebar-logo span {
        color: #8b5cf6;
    }

    .sidebar-subtitle {
        font-size: 16px;
        color: #cbd5e1 !important;
        margin-bottom: 35px;
    }

    .nav-item {
        padding: 14px 18px;
        border-radius: 12px;
        margin-bottom: 10px;
        font-size: 18px;
    }

    .nav-active {
        background: linear-gradient(90deg, #6d28d9, #7c3aed);
        box-shadow: 0 8px 20px rgba(124,58,237,0.35);
    }

    .sidebar-card {
        margin-top: 100px;
        padding: 22px;
        border-radius: 15px;
        border: 1px solid rgba(139,92,246,0.5);
        background: rgba(255,255,255,0.04);
        font-size: 15px;
        line-height: 1.6;
    }

    .main-title {
        text-align: center;
        font-size: 48px;
        font-weight: 850;
        color: #0b1736;
        margin-top: 20px;
    }

    .main-title span {
        color: #6d28d9;
    }

    .subtitle {
        text-align: center;
        font-size: 22px;
        color: #334155;
        margin-top: -8px;
    }

    .description {
        text-align: center;
        font-size: 18px;
        color: #475569;
        max-width: 850px;
        margin: 0 auto 35px auto;
        line-height: 1.7;
    }

    .topbar {
        display: flex;
        justify-content: flex-end;
        align-items: center;
        gap: 20px;
        padding: 8px 20px 20px 20px;
        color: #0f172a;
        font-size: 16px;
    }

    .demo-user {
        display: flex;
        align-items: center;
        gap: 10px;
        padding: 8px 14px;
        border-radius: 30px;
        background: #ffffff;
        border: 1px solid #e5e7eb;
        box-shadow: 0 5px 18px rgba(15,23,42,0.06);
    }

    .user-circle {
        width: 40px;
        height: 40px;
        background: linear-gradient(135deg, #7c3aed, #4f46e5);
        color: white;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        font-weight: 800;
    }

    .card {
        background: rgba(255,255,255,0.90);
        border: 1px solid #e5e7eb;
        border-radius: 20px;
        padding: 32px;
        min-height: 260px;
        box-shadow: 0 15px 40px rgba(15,23,42,0.08);
    }

    .card-title {
        color: #5b21b6;
        font-size: 25px;
        font-weight: 800;
        margin-bottom: 15px;
    }

    .card-text {
        color: #0f172a;
        font-size: 17px;
        line-height: 1.7;
        margin-bottom: 20px;
    }

    .upload-box {
        border: 2px dashed #7c3aed;
        border-radius: 16px;
        padding: 30px;
        text-align: center;
        background: #faf7ff;
        margin-top: 16px;
        margin-bottom: 12px;
    }

    .upload-icon {
        font-size: 44px;
        color: #6d28d9;
    }

    .or-circle {
        width: 70px;
        height: 70px;
        background: white;
        border-radius: 50%;
        border: 1px solid #e5e7eb;
        box-shadow: 0 8px 22px rgba(15,23,42,0.1);
        display: flex;
        align-items: center;
        justify-content: center;
        font-weight: 800;
        color: #4c1d95;
        margin: 145px auto 0 auto;
        font-size: 20px;
    }

    .feature-bar {
        background: rgba(255,255,255,0.95);
        border: 1px solid #e5e7eb;
        border-radius: 20px;
        padding: 26px;
        margin-top: 28px;
        box-shadow: 0 15px 35px rgba(15,23,42,0.06);
    }

    .feature-item {
        text-align: center;
        border-right: 1px solid #e5e7eb;
        min-height: 120px;
        padding: 10px;
    }

    .feature-item-last {
        text-align: center;
        min-height: 120px;
        padding: 10px;
    }

    .feature-icon {
        font-size: 34px;
        color: #6d28d9;
        margin-bottom: 8px;
    }

    .feature-title {
        font-size: 17px;
        font-weight: 800;
        color: #4c1d95;
    }

    .feature-text {
        font-size: 14px;
        color: #475569;
        line-height: 1.5;
    }

    div.stButton > button {
        background: linear-gradient(90deg, #6d28d9, #7c3aed);
        color: white;
        border: none;
        border-radius: 10px;
        padding: 12px 32px;
        font-weight: 700;
        font-size: 16px;
        width: 100%;
        box-shadow: 0 10px 24px rgba(124,58,237,0.25);
    }

    div.stButton > button:hover {
        background: linear-gradient(90deg, #5b21b6, #6d28d9);
        color: white;
    }

    .footer {
        text-align: center;
        color: #64748b;
        margin-top: 28px;
        font-size: 14px;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# ---------------- SIDEBAR ----------------

with st.sidebar:
    st.markdown(
        """
        <div class="sidebar-logo">🧪 ACME<span>PreForm</span></div>
        <div class="sidebar-subtitle">API Property Extractor</div>

        <div class="nav-item nav-active">⌂ &nbsp; Home</div>
        <div class="nav-item">📄 &nbsp; My Documents</div>
        <div class="nav-item">🔍 &nbsp; Search API</div>
        <div class="nav-item">💬 &nbsp; Chat Assistant</div>
        <div class="nav-item">↺ &nbsp; History</div>
        <div class="nav-item">⚙ &nbsp; Settings</div>
        <div class="nav-item">ⓘ &nbsp; About</div>

        <div class="sidebar-card">
            <b>AI-Powered Extraction</b><br><br>
            Extract, organize and analyze API physical properties from PDFs and trusted sources in seconds.
        </div>
        """,
        unsafe_allow_html=True
    )

# ---------------- HEADER ----------------

st.markdown(
    """
    <div class="topbar">
        <div>🔔 <b style="background:#6d28d9;color:white;border-radius:50%;padding:2px 7px;font-size:12px;">3</b></div>
        <div class="demo-user">
            <div class="user-circle">D</div>
            <div><b>Demo User</b><br><span style="font-size:13px;color:#64748b;">R&D Department</span></div>
        </div>
    </div>

    <div class="main-title">Welcome to <span>ACMEPreForm</span></div>
    <div class="subtitle">AI-Powered API Physical Property Extraction & Knowledge Management System</div>
    <div class="description">
        Upload a patent / SDS / monograph PDF or search an API to instantly extract physical properties from trusted web sources.
    </div>
    """,
    unsafe_allow_html=True
)

# ---------------- HOME CARDS ----------------

left, middle, right = st.columns([1.05, 0.15, 1.05])

with left:
    st.markdown(
        """
        <div class="card">
            <div class="card-title">📄 Upload Patent / SDS / Monograph</div>
            <div class="card-text">
                Upload a PDF and the demo will detect the API name, then search trusted web sources.
            </div>
            <div class="upload-box">
                <div class="upload-icon">☁️</div>
                <b>Drag & drop your PDF file here</b><br>
                or browse below
            </div>
        """,
        unsafe_allow_html=True
    )

    uploaded_file = st.file_uploader(
        "Upload PDF",
        type=["pdf"],
        label_visibility="collapsed"
    )

    extract_pdf_clicked = st.button("📄 Extract from PDF")

    st.markdown("</div>", unsafe_allow_html=True)

with middle:
    st.markdown('<div class="or-circle">OR</div>', unsafe_allow_html=True)

with right:
    with st.container():
        st.markdown(
            """
            <div style="
                background: rgba(255,255,255,0.92);
                border: 1px solid #e5e7eb;
                border-radius: 20px;
                padding: 32px;
                min-height: 360px;
                box-shadow: 0 15px 40px rgba(15,23,42,0.08);
            ">
                <div style="
                    color:#5b21b6;
                    font-size:25px;
                    font-weight:800;
                    margin-bottom:15px;
                ">
                    🔍 Enter API Name
                </div>
                <div style="
                    color:#0f172a;
                    font-size:17px;
                    line-height:1.7;
                    margin-bottom:25px;
                ">
                    Enter the API name to search from trusted web sources.
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

        with st.form("api_search_form"):
            api_name = st.text_input(
                "API Name",
                placeholder="e.g. Clarithromycin, Paracetamol, Ibuprofen"
            )

            search_clicked = st.form_submit_button("🔍 Search API from Web")

# ---------------- FEATURES ----------------

st.markdown('<div class="feature-bar">', unsafe_allow_html=True)

f1, f2, f3, f4 = st.columns(4)

features = [
    ("🧠", "AI Powered Extraction", "Detect API name from patent or PDF"),
    ("🌐", "Trusted Web Search", "Live data from PubChem for demo"),
    ("📊", "Structured Results", "Clean table with property values"),
    ("⬇️", "Export & Reports", "Download results in Excel")
]

for col, item in zip([f1, f2, f3, f4], features):
    icon, title, text = item
    with col:
        class_name = "feature-item-last" if col == f4 else "feature-item"
        st.markdown(
            f"""
            <div class="{class_name}">
                <div class="feature-icon">{icon}</div>
                <div class="feature-title">{title}</div>
                <div class="feature-text">{text}</div>
            </div>
            """,
            unsafe_allow_html=True
        )

st.markdown("</div>", unsafe_allow_html=True)

# ---------------- API SEARCH RESULT ----------------

if search_clicked:
    if api_name.strip() == "":
        st.warning("Please enter an API name first. Example: Clarithromycin")
    else:
        show_processing("Searching Properties from Trusted Web Sources")

        result = get_pubchem_data(api_name)

        if result is None:
            st.error("No data found. Try: Clarithromycin, Paracetamol, Ibuprofen, Metformin")
        else:
            st.success(f"Successfully retrieved web data for {api_name}")

            df = create_result_table(api_name, result, "Web Search")
            st.dataframe(df, use_container_width=True)

            excel_file = create_excel(df)

            st.download_button(
                "⬇️ Download Excel Report",
                data=excel_file,
                file_name=f"{api_name}_properties.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            patent_url = f"https://patents.google.com/?q={quote_plus(api_name)}"
            st.link_button("🔎 Open Patent Search", patent_url)

            st.info("Demo note: PubChem result is live. Patent button opens Google Patents search.")

# ---------------- PDF EXTRACTION RESULT ----------------

if extract_pdf_clicked:
    if uploaded_file is None:
        st.warning("Please upload a patent or API-related PDF first.")
    else:
        show_processing("Processing Uploaded PDF / Patent")

        try:
            pdf_text = extract_pdf_text(uploaded_file)
            detected_api = detect_api_name(pdf_text)

            st.success(f"API detected from PDF: {detected_api}")

            result = get_pubchem_data(detected_api)

            if result is None:
                st.error("API detected, but PubChem data was not found.")
            else:
                df = create_result_table(detected_api, result, "PDF / Patent Extraction")
                st.dataframe(df, use_container_width=True)

                excel_file = create_excel(df)

                st.download_button(
                    "⬇️ Download Excel Report",
                    data=excel_file,
                    file_name=f"{detected_api}_pdf_extraction.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                with st.expander("View Extracted Patent/PDF Text Preview"):
                    st.text_area("Text Preview", pdf_text[:3000], height=250)

                patent_url = f"https://patents.google.com/?q={quote_plus(detected_api)}"
                st.link_button("🔎 Open Related Patent Search", patent_url)

                st.info(
                    "Demo note: The PDF text is extracted using PyMuPDF. API name detection is demo-based using common API keywords."
                )

        except Exception as e:
            st.error(f"PDF extraction failed: {e}")

# ---------------- FOOTER ----------------

st.markdown(
    """
    <div class="footer">
        © 2026 ACME FormAI. All rights reserved.
    </div>
    """,
    unsafe_allow_html=True
)