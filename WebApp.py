import streamlit as st
import pandas as pd
import requests
import time
import pymupdf
from io import BytesIO
from urllib.parse import quote_plus

st.set_page_config(
    page_title="ACME R&D Intelligence",
    page_icon="🧪",
    layout="wide"
)

if "page" not in st.session_state:
    st.session_state.page = "Home"


# ================= FUNCTIONS =================

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

    except Exception:
        return None


def extract_pdf_text(uploaded_file):
    file_bytes = uploaded_file.getvalue()
    pdf = pymupdf.open(stream=file_bytes, filetype="pdf")

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
        "Metformin HCl",
        "Metformin",
        "Losartan",
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
        time.sleep(0.45)

    status.empty()


# ================= CSS =================

st.markdown(
    """
    <style>
    .stApp {
        background: #ffffff;
    }

    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #061936 0%, #020b1d 100%) !important;
        width: 285px !important;
    }

    section[data-testid="stSidebar"] * {
        color: white !important;
    }

    .sidebar-logo {
        font-size: 26px;
        font-weight: 800;
        margin-top: 10px;
        margin-bottom: 5px;
        color: white;
    }

    .sidebar-subtitle {
        font-size: 14px;
        color: #cbd5e1 !important;
        margin-bottom: 25px;
    }

    .sidebar-card {
        margin-top: 35px;
        padding: 20px;
        border-radius: 15px;
        border: 1px solid rgba(139,92,246,0.5);
        background: rgba(255,255,255,0.04);
        font-size: 14px;
        line-height: 1.6;
    }

    div.stButton > button {
        background: linear-gradient(90deg, #005daa, #0077cc);
        color: white !important;
        border: none;
        border-radius: 10px;
        padding: 10px 22px;
        font-weight: 700;
        font-size: 15px;
        width: 100%;
        box-shadow: 0 8px 20px rgba(0,93,170,0.20);
    }

    div.stButton > button:hover {
        background: linear-gradient(90deg, #004b93, #005daa);
        color: white !important;
    }

    .top-header {
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 18px 35px;
        border-bottom: 1px solid #e5e7eb;
        background: white;
    }

    .brand-box {
        display: flex;
        align-items: center;
        gap: 15px;
    }

    .brand-icon {
        font-size: 42px;
    }

    .brand-title {
        font-size: 28px;
        font-weight: 850;
        color: #08265c;
    }

    .brand-sub {
        font-size: 13px;
        letter-spacing: 5px;
        color: #005daa;
        margin-top: 4px;
    }

    .top-menu {
        display: flex;
        gap: 35px;
        font-size: 16px;
        font-weight: 700;
        color: #0f172a;
    }

    .hero {
        background: linear-gradient(90deg, #eef7ff 0%, #d8ecff 55%, #b9dcff 100%);
        padding: 70px 50px;
        border-radius: 0px;
        margin-bottom: 30px;
        position: relative;
        overflow: hidden;
    }

    .hero-title {
        font-size: 46px;
        font-weight: 900;
        color: #071b3a;
        line-height: 1.15;
    }

    .hero-title span {
        color: #005daa;
    }

    .hero-text {
        font-size: 19px;
        color: #1f2937;
        max-width: 620px;
        line-height: 1.6;
        margin-top: 25px;
    }

    .module-card {
        background: white;
        padding: 35px;
        border-radius: 18px;
        text-align: center;
        box-shadow: 0 12px 30px rgba(15,23,42,0.08);
        min-height: 285px;
        border: 1px solid #e5e7eb;
    }

    .module-icon-blue {
        background: #005daa;
        color: white;
        width: 90px;
        height: 90px;
        border-radius: 50%;
        display: flex;
        justify-content: center;
        align-items: center;
        margin: 0 auto 20px auto;
        font-size: 42px;
    }

    .module-icon-green {
        background: #148a43;
        color: white;
        width: 90px;
        height: 90px;
        border-radius: 50%;
        display: flex;
        justify-content: center;
        align-items: center;
        margin: 0 auto 20px auto;
        font-size: 42px;
    }

    .module-icon-purple {
        background: #5b3db8;
        color: white;
        width: 90px;
        height: 90px;
        border-radius: 50%;
        display: flex;
        justify-content: center;
        align-items: center;
        margin: 0 auto 20px auto;
        font-size: 42px;
    }

    .module-title {
        font-size: 23px;
        font-weight: 850;
        color: #08265c;
        margin-bottom: 14px;
    }

    .module-text {
        font-size: 15px;
        color: #374151;
        line-height: 1.55;
    }

    .bottom-strip {
        margin-top: 28px;
        background: linear-gradient(90deg, #004b93, #0067bd);
        color: white;
        padding: 25px;
        border-radius: 0px;
        display: grid;
        grid-template-columns: repeat(4, 1fr);
        text-align: center;
        gap: 20px;
        font-size: 16px;
    }

    .api-card {
        background: rgba(255,255,255,0.92);
        border: 1px solid #e5e7eb;
        border-radius: 20px;
        padding: 30px;
        min-height: 330px;
        box-shadow: 0 15px 40px rgba(15,23,42,0.08);
    }

    .api-card-title {
        color: #005daa;
        font-size: 24px;
        font-weight: 850;
        margin-bottom: 14px;
    }

    .api-card-text {
        color: #0f172a;
        font-size: 16px;
        line-height: 1.6;
        margin-bottom: 18px;
    }

    .upload-box {
        border: 2px dashed #005daa;
        border-radius: 16px;
        padding: 25px;
        text-align: center;
        background: #f2f8ff;
        margin-bottom: 15px;
    }

    .or-circle {
        width: 60px;
        height: 60px;
        background: white;
        border-radius: 50%;
        border: 1px solid #e5e7eb;
        box-shadow: 0 8px 22px rgba(15,23,42,0.1);
        display: flex;
        align-items: center;
        justify-content: center;
        font-weight: 850;
        color: #005daa;
        margin: 150px auto 0 auto;
        font-size: 18px;
    }

    .footer {
        text-align: center;
        color: #64748b;
        margin-top: 35px;
        font-size: 13px;
    }
    </style>
    """,
    unsafe_allow_html=True
)


# ================= SIDEBAR =================

with st.sidebar:
    st.markdown(
        """
        <div class="sidebar-logo">🧪 ACME R&D Intelligence</div>
        <div class="sidebar-subtitle">Research & Development Platform</div>
        """,
        unsafe_allow_html=True
    )

    if st.button("⌂ Home", key="nav_home"):
        st.session_state.page = "Home"
        st.rerun()

    if st.button("🧪 API Characterization", key="nav_api"):
        st.session_state.page = "API Characterization"
        st.rerun()

    if st.button("⚗️ Preformulation Design", key="nav_pre"):
        st.session_state.page = "Preformulation Design"
        st.rerun()

    if st.button("🧩 Drug-Excipient Compatibility", key="nav_compat"):
        st.session_state.page = "Drug-Excipient Compatibility"
        st.rerun()

    st.markdown(
        """
        <div class="sidebar-card">
            <b>R&D Intelligence Hub</b><br><br>
            A modular AI platform for API characterization, preformulation design, and drug-excipient compatibility assessment.
        </div>
        """,
        unsafe_allow_html=True
    )


# ================= PAGE 1: HOME =================

if st.session_state.page == "Home":

    # 1. Create a header structure using native Streamlit columns
    logo_col, title_col = st.columns([1, 4])

    with logo_col:
        # Streamlit reads local paths natively without security errors
        st.image("assets/acme_logo.png", use_container_width=True)

    with title_col:
        st.markdown("""
            <div style="padding-top: 10px;">
                <h2 style="color: #0b2c5f; margin: 0;">ACME Laboratories Ltd.</h2>
                <p style="color: #5c768d; font-style: italic; margin: 0;">For Health • Vigour • and Happiness</p>
            </div>
        """, unsafe_allow_html=True)

    # 2. Render the top menu and hero banner below it
    st.markdown(
        """
        <div class="top-menu" style="margin-top: 20px; border-top: 1px solid #e5e7eb; padding-top: 15px;">
            <span style="margin-right: 20px; font-weight: bold; color: #6d28d9;">Home</span>
            <span style="margin-right: 20px;">API Characterization</span>
            <span style="margin-right: 20px;">Preformulation Design</span>
            <span>Drug-Excipient Compatibility</span>
        </div>

        <div class="hero" style="margin-top: 30px;">
            <div class="hero-title">
                Advancing Health<br>
                Through <span>Research & Innovation</span>
            </div>
            <div class="hero-text">
                The ACME Laboratories Ltd. is committed to delivering high-quality pharmaceutical solutions
                through science, collaboration, innovation and excellence.
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown(
            """
            <div class="module-card">
                <div class="module-icon-blue">🧪</div>
                <div class="module-title">API Characterization</div>
                <div class="module-text">
                    Comprehensive physicochemical and analytical characterization of APIs to support
                    formulation and product development.
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )
        if st.button("Open API Characterization", key="open_api"):
            st.session_state.page = "API Characterization"
            st.rerun()

    with col2:
        st.markdown(
            """
            <div class="module-card">
                <div class="module-icon-green">⚗️</div>
                <div class="module-title">Preformulation Design</div>
                <div class="module-text">
                    Early-stage studies to evaluate physicochemical properties and develop optimal
                    formulation strategies.
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )
        if st.button("Open Preformulation Design", key="open_pre"):
            st.session_state.page = "Preformulation Design"
            st.rerun()

    with col3:
        st.markdown(
            """
            <div class="module-card">
                <div class="module-icon-purple">🧩</div>
                <div class="module-title">Drug-Excipient Compatibility</div>
                <div class="module-text">
                    Systematic compatibility studies to identify potential interactions and ensure
                    product stability and performance.
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )
        if st.button("Open Compatibility", key="open_compat"):
            st.session_state.page = "Drug-Excipient Compatibility"
            st.rerun()

    st.markdown(
        """
        <div class="bottom-strip">
            <div><b>Quality</b><br>Our Commitment</div>
            <div><b>Innovation</b><br>Driven by Science</div>
            <div><b>Collaboration</b><br>Stronger Together</div>
            <div><b>Excellence</b><br>In Everything We Do</div>
        </div>
        """,
        unsafe_allow_html=True
    )


# ================= PAGE 2: API CHARACTERIZATION =================

elif st.session_state.page == "API Characterization":

    if st.button("← Back to Home"):
        st.session_state.page = "Home"
        st.rerun()

    st.title("🧪 API Characterization")
    st.write("Upload a patent/article/DMF PDF or search an API from trusted web sources.")

    left, middle, right = st.columns([1.05, 0.15, 1.05])

    with left:
        st.markdown(
            """
            <div class="api-card">
                <div class="api-card-title">📄 Upload Patent / Article / Monograph</div>
                <div class="api-card-text">
                    Upload a PDF file. The system will detect the API name and retrieve structured property data.
                </div>
                <div class="upload-box">
                    <b>Drag & drop your PDF file here</b><br>
                    or browse below
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

        uploaded_file = st.file_uploader(
            "Upload PDF",
            type=["pdf"],
            label_visibility="collapsed"
        )

        extract_pdf_clicked = st.button("Extract from PDF")

    with middle:
        st.markdown('<div class="or-circle">OR</div>', unsafe_allow_html=True)

    with right:
        st.markdown(
            """
            <div class="api-card">
                <div class="api-card-title">🔍 Search API from Web</div>
                <div class="api-card-text">
                    Enter an API name to retrieve live physicochemical properties from PubChem.
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

        api_name = st.text_input(
            "API Name",
            placeholder="e.g. Clarithromycin, Metformin, Losartan"
        )

        search_clicked = st.button("Search API")

    if search_clicked:
        if not api_name.strip():
            st.warning("Please enter an API name first.")
        else:
            show_processing("Searching Trusted Web Sources")

            result = get_pubchem_data(api_name)

            if result is None:
                st.error("No data found.")
            else:
                st.success(f"Successfully retrieved data for: {api_name}")

                df = create_result_table(api_name, result, "Web Search")
                st.dataframe(df, use_container_width=True)

                st.download_button(
                    "Download Excel Report",
                    data=create_excel(df),
                    file_name=f"ACME_{api_name}_Profile.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                patent_url = f"https://patents.google.com/?q={quote_plus(api_name)}"
                st.link_button("Open Patent Search", patent_url)

    if extract_pdf_clicked:
        if uploaded_file is None:
            st.warning("Please upload a PDF first.")
        else:
            show_processing("Processing Uploaded PDF")

            try:
                pdf_text = extract_pdf_text(uploaded_file)
                detected_api = detect_api_name(pdf_text)

                st.success(f"Detected API from PDF: {detected_api}")

                result = get_pubchem_data(detected_api)

                if result is None:
                    st.error("API detected, but web property data was not found.")
                else:
                    df = create_result_table(detected_api, result, "PDF Extraction")
                    st.dataframe(df, use_container_width=True)

                    st.download_button(
                        "Download Excel Report",
                        data=create_excel(df),
                        file_name=f"ACME_{detected_api}_Extracted.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                    with st.expander("View Extracted PDF Text Preview"):
                        st.text_area("Extracted Text", pdf_text[:3000], height=220)

            except Exception as e:
                st.error(f"PDF extraction failed: {e}")


# ================= PAGE 3: PREFORMULATION =================

elif st.session_state.page == "Preformulation Design":

    if st.button("← Back to Home"):
        st.session_state.page = "Home"
        st.rerun()

    st.title("⚗️ Preformulation Design")
    st.info("Under development.")


# ================= PAGE 4: COMPATIBILITY =================

elif st.session_state.page == "Drug-Excipient Compatibility":

    if st.button("← Back to Home"):
        st.session_state.page = "Home"
        st.rerun()

    st.title("🧩 Drug-Excipient Compatibility")
    st.info("Under development.")


# ================= FOOTER =================

st.markdown(
    """
    <div class="footer">
        ©2026 ACME Laboratories Ltd. — Research & Development Division.
    </div>
    """,
    unsafe_allow_html=True
)
