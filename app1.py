import streamlit as st
import os
import tempfile
import fitz
import pandas as pd
from docx import Document
from dotenv import load_dotenv
import google.generativeai as genai
from fpdf import FPDF
import textwrap
import re
from datetime import date

# ─── Configuration ────────────────────────────────────────────────────────────

load_dotenv()

try:
    GEMINI_API_KEY = os.getenv("GEMINI_API_KEY") or st.secrets.get("GEMINI_API_KEY")
except Exception:
    GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

if GEMINI_API_KEY:
    GEMINI_API_KEY = GEMINI_API_KEY.strip()
    genai.configure(api_key=GEMINI_API_KEY)

st.set_page_config(
    page_title="Practicum Review Report Generator",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── Custom CSS ───────────────────────────────────────────────────────────────

st.markdown("""
<style>
    /* Import Inter font */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;600&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }

    /* Main background */
    .stApp {
        background-color: #F7F8FC;
    }

    /* Sidebar */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0F2044 0%, #1A3A6B 100%);
        border-right: none;
    }
    [data-testid="stSidebar"] * {
        color: #C9D6EE !important;
    }
    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3 {
        color: #FFFFFF !important;
    }
    [data-testid="stSidebar"] .stMarkdown p {
        font-size: 0.85rem;
        line-height: 1.6;
    }

    /* Page header */
    .page-header {
        background: linear-gradient(135deg, #0F2044 0%, #1A3A6B 60%, #1E4D8C 100%);
        border-radius: 12px;
        padding: 32px 36px;
        margin-bottom: 28px;
        color: white;
    }
    .page-header h1 {
        font-size: 1.75rem;
        font-weight: 700;
        margin: 0 0 8px 0;
        letter-spacing: -0.02em;
        color: white;
    }
    .page-header p {
        font-size: 0.95rem;
        opacity: 0.8;
        margin: 0;
        color: #C9D6EE;
    }

    /* Step cards */
    .step-card {
        background: white;
        border-radius: 10px;
        padding: 24px 28px;
        margin-bottom: 18px;
        border: 1px solid #E2E8F0;
        box-shadow: 0 1px 4px rgba(15,32,68,0.06);
    }
    .step-label {
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.7rem;
        font-weight: 600;
        color: #1A3A6B;
        letter-spacing: 0.12em;
        text-transform: uppercase;
        margin-bottom: 6px;
    }
    .step-title {
        font-size: 1.05rem;
        font-weight: 600;
        color: #0F2044;
        margin-bottom: 4px;
    }
    .step-desc {
        font-size: 0.85rem;
        color: #64748B;
        margin-bottom: 14px;
    }

    /* Report container */
    .report-container {
        background: white;
        border-radius: 12px;
        padding: 32px 36px;
        border: 1px solid #E2E8F0;
        box-shadow: 0 2px 12px rgba(15,32,68,0.08);
        margin-top: 8px;
    }

    /* Status badges */
    .badge {
        display: inline-block;
        padding: 4px 12px;
        border-radius: 20px;
        font-size: 0.78rem;
        font-weight: 600;
        font-family: 'JetBrains Mono', monospace;
    }
    .badge-success { background: #DCFCE7; color: #15803D; }
    .badge-warning { background: #FEF9C3; color: #854D0E; }
    .badge-error   { background: #FEE2E2; color: #991B1B; }

    /* Generate button override */
    div[data-testid="stButton"] > button {
        background: linear-gradient(135deg, #1A3A6B, #1E4D8C);
        color: white;
        border: none;
        border-radius: 8px;
        padding: 12px 32px;
        font-family: 'Inter', sans-serif;
        font-weight: 600;
        font-size: 0.95rem;
        letter-spacing: 0.01em;
        transition: opacity 0.2s;
        width: 100%;
    }
    div[data-testid="stButton"] > button:hover {
        opacity: 0.88;
        color: white;
    }

    /* Download button */
    div[data-testid="stDownloadButton"] > button {
        background: white;
        color: #1A3A6B;
        border: 2px solid #1A3A6B;
        border-radius: 8px;
        padding: 10px 28px;
        font-family: 'Inter', sans-serif;
        font-weight: 600;
        font-size: 0.92rem;
        transition: all 0.2s;
        width: 100%;
        margin-top: 8px;
    }
    div[data-testid="stDownloadButton"] > button:hover {
        background: #0F2044;
        color: white;
    }

    /* Divider */
    hr {
        border: none;
        border-top: 1px solid #E2E8F0;
        margin: 24px 0;
    }

    /* File uploader area */
    [data-testid="stFileUploader"] {
        border-radius: 8px;
    }

    /* Spinner text */
    .stSpinner > div {
        color: #1A3A6B !important;
    }
</style>
""", unsafe_allow_html=True)


# ─── Session state init ────────────────────────────────────────────────────────

if "report" not in st.session_state:
    st.session_state.report = None
if "pdf_bytes" not in st.session_state:
    st.session_state.pdf_bytes = None


# ─── Sidebar ──────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("## 📋 How It Works")
    st.markdown("""
**Step 1 — Upload Practicum**
Upload the student's practicum or assignment document (PDF or DOCX).

**Step 2 — Upload SOP / Rubric**
Upload the evaluation rubric or SOP in Excel format (.xlsx).

**Step 3 — Generate Report**
Click **Generate AI Review** to produce a structured, criteria-based report.

**Step 4 — Download PDF**
Download the formatted report as a PDF for record-keeping.
""")
    st.markdown("---")
    st.markdown("### ⚙️ Model Info")
    st.markdown("""
- **Model:** Gemini 2.5 Flash Lite
- **Engine:** Google Generative AI
- **Output:** Structured review report
""")
    st.markdown("---")
    st.markdown(
        "<p style='font-size:0.75rem; opacity:0.6;'>Automated Practicum Review System · "
        + str(date.today().strftime("%Y")) + "</p>",
        unsafe_allow_html=True
    )


# ─── Page Header ──────────────────────────────────────────────────────────────

st.markdown("""
<div class="page-header">
    <h1>📋 Practicum Review Report Generator</h1>
    <p>Upload a practicum document and SOP/rubric file — the system will evaluate the submission
    against your criteria and generate a structured, professional review report.</p>
</div>
""", unsafe_allow_html=True)

# API key warning
if not GEMINI_API_KEY:
    st.error("⚠️  **Gemini API key not found.** Add `GEMINI_API_KEY` to your `.env` file or Streamlit secrets to continue.")
    st.stop()


# ─── File extraction helpers ──────────────────────────────────────────────────

def read_pdf(file_path: str) -> str:
    text = ""
    with fitz.open(file_path) as pdf:
        for page in pdf:
            text += page.get_text() + "\n"
    return text.strip()


def read_docx(file_path: str) -> str:
    document = Document(file_path)
    return "\n".join(p.text for p in document.paragraphs).strip()


def extract_practicum_text(file_path: str) -> str:
    if file_path.endswith(".pdf"):
        return read_pdf(file_path)
    elif file_path.endswith(".docx"):
        return read_docx(file_path)
    return ""


def read_sop_excel(file_path: str) -> str:
    excel_file = pd.ExcelFile(file_path)
    sop_parts = []
    for sheet_name in excel_file.sheet_names:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        sop_parts.append(f"Sheet: {sheet_name}\n{df.to_string(index=False)}")
    return "\n\n".join(sop_parts)


# ─── AI Review ────────────────────────────────────────────────────────────────

def generate_ai_review(practicum_text: str, sop_text: str) -> str:
    practicum_text = practicum_text[:18000]
    sop_text = sop_text[:12000]
    today = date.today().strftime("%d %B %Y")

    model = genai.GenerativeModel("models/gemini-2.5-flash-lite")

    prompt = f"""
You are an Automated Practicum Review System.

Review the uploaded practicum or activity document strictly against the SOP and rubric provided.

Guidelines:
- Use formal UK English.
- Maintain a professional, balanced, diagnostic, and evaluator-friendly tone.
- Do not assume missing information. Do not produce generic feedback.
- Highlight strengths first, then areas for improvement.
- Keep the report concise but meaningful.
- Use a markdown table only for the Scorecard section.
- Do not use tables in the Detailed Feedback section.
- Today's date is {today}.

SOP AND RUBRIC:
{sop_text}

PRACTICUM / ACTIVITY DOCUMENT:
{practicum_text}

Generate the report using this exact structure:

# PRACTICUM REVIEW REPORT

Document Title:
[Extract from document or write: Not specified]

Assignment Code:
[Extract from document or write: Not specified]

Course:
[Extract from document or write: Not specified]

Review Date:
{today}

Reviewer:
Automated Practicum Review System

Product Type:
[Assignment / Discussion / Field Work / Practicum]

---

## Executive Summary
Write 5–6 concise sentences covering overall quality, key strengths, and key areas for improvement.

---

## Overall Score

Overall Quality Score: X / 5
Overall Rating: [Poor / Needs Improvement / Satisfactory / Good / Excellent]

---



## Scorecard

Create the scorecard in this markdown table format:

| # | Criteria | Score / 5 | Rating | Short Justification |
|---|----------|-----------|--------|---------------------|
| 1 | Adherence to Framework |  |  |  |
| 2 | Activity Description |  |  |  |
| 3 | Required Submission |  |  |  |
| 4 | Procedure / Prompts |  |  |  |
| 5 | Learning Outcome Alignment |  |  |  |
| 6 | Rubric Quality |  |  |  |
| 7 | Learner-facing Clarity |  |  |  |
| 8 | Grammar and Consistency |  |  |  |
| 9 | Organisation and Flow |  |  |  |

---

## Detailed Feedback

For each criterion, use this exact format:

### Criterion 1: Adherence to Framework

**Strengths:**
- Write 1-2 bullet points explaining what is done well.

**Areas for Improvement:**
- Write 1-2 bullet points explaining what needs improvement.

**Suggested Next Step:**
- Write one clear actionable next step.

Repeat the same format for all 9 criteria.

Do not combine Strengths, Areas for Improvement, and Suggested Next Step in one paragraph.
Keep each part clearly separated.

---


---

## Final Recommendation

[Choose one: Approved / Approved with Minor Improvements / Needs Revision / Major Revision Required]

Brief justification (2–3 sentences).

---

## Top 3 Priority Improvements

- [Most critical improvement]
- [Second improvement]
- [Third improvement]
"""

    response = model.generate_content(prompt)
    return response.text


# ─── PDF Builder ──────────────────────────────────────────────────────────────

def sanitise(text: str) -> str:
    """Replace Unicode characters that FPDF's core fonts cannot encode."""
    replacements = {
        "\u2013": "-", "\u2014": "-",
        "\u2018": "'", "\u2019": "'",
        "\u201c": '"', "\u201d": '"',
        "\u2022": "-", "\u2019": "'",
        "\u00a0": " ", "\t": "    ",
        "`": "'",
    }
    for orig, repl in replacements.items():
        text = text.replace(orig, repl)
    # Remove any remaining non-latin-1 characters
    text = text.encode("latin-1", errors="replace").decode("latin-1")
    return text


def clean_markdown(text: str) -> str:
    """Strip markdown syntax for plain-text PDF rendering."""
    text = re.sub(r"\*{1,3}(.*?)\*{1,3}", r"\1", text)   # bold/italic
    text = re.sub(r"_{1,2}(.*?)_{1,2}", r"\1", text)       # underline
    text = re.sub(r"#+\s*", "", text)                       # headings
    text = re.sub(r"[-\s|]{10,}", "", text)                 # table separators
    text = re.sub(r" {2,}", " ", text)                      # multiple spaces
    return text


def create_pdf_report(report_text: str) -> bytes:
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=20)
    pdf.set_margins(18, 18, 18)
    pdf.add_page()

    page_width = pdf.w - 36  # usable width

    # ── Cover block ──
    pdf.set_fill_color(15, 32, 68)   # navy
    pdf.rect(0, 0, pdf.w, 42, "F")
    pdf.set_font("Helvetica", "B", 17)
    pdf.set_text_color(255, 255, 255)
    pdf.set_xy(18, 12)
    pdf.cell(page_width, 10, "PRACTICUM REVIEW REPORT", ln=True)
    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(180, 200, 230)
    pdf.set_x(18)
    pdf.cell(page_width, 7, "Automated Practicum Review System", ln=True)
    pdf.ln(14)

    # ── Body ──
    pdf.set_text_color(30, 30, 30)
    clean = clean_markdown(report_text)

    for raw_line in clean.split("\n"):
        line = sanitise(raw_line.strip())

        if not line:
            pdf.ln(3)
            continue

        # Section divider (---)
        if re.match(r"^-{3,}$", line):
            pdf.set_draw_color(200, 210, 230)
            pdf.set_line_width(0.4)
            pdf.line(18, pdf.get_y(), pdf.w - 18, pdf.get_y())
            pdf.ln(4)
            continue

        # Detect heading levels
        is_h1 = line.isupper() and len(line) < 70
        is_h2 = line.endswith(":") and len(line) < 55 and line[0].isupper()
        is_bullet = line.startswith("- ") or line.startswith("* ")

        if is_h1:
            pdf.ln(3)
            pdf.set_font("Helvetica", "B", 12)
            pdf.set_text_color(15, 32, 68)
        elif is_h2:
            pdf.ln(2)
            pdf.set_font("Helvetica", "B", 10)
            pdf.set_text_color(26, 58, 107)
        elif is_bullet:
            pdf.set_font("Helvetica", "", 9.5)
            pdf.set_text_color(40, 40, 40)
            line = "  " + line          # indent bullet
        else:
            pdf.set_font("Helvetica", "", 9.5)
            pdf.set_text_color(40, 40, 40)

        wrapped = textwrap.wrap(line, width=95, break_long_words=True, break_on_hyphens=True)
        for wline in wrapped:
            pdf.set_x(18)
            pdf.multi_cell(page_width, 5.5, wline)

        pdf.ln(0.8)

    # ── Footer ──
    pdf.set_y(-16)
    pdf.set_font("Helvetica", "", 7.5)
    pdf.set_text_color(150, 150, 160)
    pdf.set_x(18)
    pdf.cell(
        page_width, 8,
        f"Automated Practicum Review System  ·  Generated {date.today().strftime('%d %B %Y')}",
        align="C"
    )

    return bytes(pdf.output(dest="S"))


# ─── Upload UI ────────────────────────────────────────────────────────────────

col1, col2 = st.columns(2, gap="large")

with col1:
    st.markdown("""
<div class="step-card">
    <div class="step-label">Step 01</div>
    <div class="step-title">Upload Practicum Document</div>
    <div class="step-desc">Accepted formats: PDF, DOCX</div>
</div>
""", unsafe_allow_html=True)
    practicum_file = st.file_uploader(
        "Practicum or assignment file",
        type=["pdf", "docx"],
        label_visibility="collapsed",
        key="practicum_upload"
    )
    if practicum_file:
        st.markdown(f'<span class="badge badge-success">✓ {practicum_file.name}</span>', unsafe_allow_html=True)

with col2:
    st.markdown("""
<div class="step-card">
    <div class="step-label">Step 02</div>
    <div class="step-title">Upload SOP / Rubric</div>
    <div class="step-desc">Accepted format: Excel (.xlsx)</div>
</div>
""", unsafe_allow_html=True)
    sop_file = st.file_uploader(
        "SOP or rubric Excel file",
        type=["xlsx"],
        label_visibility="collapsed",
        key="sop_upload"
    )
    if sop_file:
        st.markdown(f'<span class="badge badge-success">✓ {sop_file.name}</span>', unsafe_allow_html=True)


# ─── Generate ────────────────────────────────────────────────────────────────

st.markdown("<br>", unsafe_allow_html=True)
st.markdown("""
<div class="step-card">
    <div class="step-label">Step 03</div>
    <div class="step-title">Generate AI Review</div>
    <div class="step-desc">The AI will evaluate the submission against your rubric and produce a structured report.</div>
</div>
""", unsafe_allow_html=True)

if st.button("⚡  Generate AI Review", use_container_width=True):
    if not practicum_file:
        st.error("Please upload a practicum or assignment document before generating.")
    elif not sop_file:
        st.error("Please upload the SOP / rubric Excel file before generating.")
    else:
        # Use temp files — no leftover disk state
        with tempfile.NamedTemporaryFile(
            delete=False, suffix=os.path.splitext(practicum_file.name)[1]
        ) as tmp_prac:
            tmp_prac.write(practicum_file.getbuffer())
            practicum_path = tmp_prac.name

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_sop:
            tmp_sop.write(sop_file.getbuffer())
            sop_path = tmp_sop.name

        try:
            with st.spinner("Extracting text from uploaded files…"):
                practicum_text = extract_practicum_text(practicum_path)
                sop_text = read_sop_excel(sop_path)

            if not practicum_text:
                st.error("Could not extract text from the practicum file. Ensure the PDF/DOCX contains readable text (not scanned images).")
                st.stop()

            with st.spinner("Analysing submission and generating review…"):
                report = generate_ai_review(practicum_text, sop_text)

            with st.spinner("Building PDF report…"):
                pdf_bytes = create_pdf_report(report)

            st.session_state.report = report
            st.session_state.pdf_bytes = pdf_bytes
            st.success("✅  Review generated successfully.")

        except Exception as err:
            st.error(f"An error occurred during report generation: `{err}`")

        finally:
            # Clean up temp files
            for path in [practicum_path, sop_path]:
                try:
                    os.unlink(path)
                except OSError:
                    pass


# ─── Report Display & Download ────────────────────────────────────────────────

if st.session_state.report:
    st.markdown("<hr>", unsafe_allow_html=True)

    col_title, col_dl = st.columns([3, 1])
    with col_title:
        st.markdown("### 📄  Generated Review Report")
    with col_dl:
        if st.session_state.pdf_bytes:
            st.download_button(
                label="⬇  Download PDF",
                data=st.session_state.pdf_bytes,
                file_name=f"practicum_review_{date.today().strftime('%Y%m%d')}.pdf",
                mime="application/pdf",
                key="dl_top",
                use_container_width=True,
            )

    st.markdown('<div class="report-container">', unsafe_allow_html=True)
    st.markdown(st.session_state.report)
    st.markdown("</div>", unsafe_allow_html=True)

    # Second download button at the bottom
    st.markdown("<br>", unsafe_allow_html=True)
    if st.session_state.pdf_bytes:
        st.download_button(
            label="⬇  Download PDF Report",
            data=st.session_state.pdf_bytes,
            file_name=f"practicum_review_{date.today().strftime('%Y%m%d')}.pdf",
            mime="application/pdf",
            key="dl_bottom",
            use_container_width=True,
        )