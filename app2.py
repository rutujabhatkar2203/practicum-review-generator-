import streamlit as st
import os
import tempfile
import fitz
import pandas as pd
from docx import Document
from dotenv import load_dotenv
import google.generativeai as genai
import markdown as md_lib
from io import BytesIO
import re
from datetime import date

# ─── Try importing PDF backends (prefer weasyprint, fallback to reportlab) ───

PDF_BACKEND = None

try:
    from weasyprint import HTML as WeasyprintHTML
    PDF_BACKEND = "weasyprint"
except Exception:
    pass

if PDF_BACKEND is None:
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib import colors
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import mm
        from reportlab.platypus import (
            SimpleDocTemplate, Paragraph, Spacer, HRFlowable,
            Table, TableStyle, KeepTogether
        )
        from reportlab.lib.enums import TA_LEFT, TA_CENTER
        PDF_BACKEND = "reportlab"
    except Exception:
        PDF_BACKEND = None

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
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;600&display=swap');

    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
    .stApp { background-color: #F7F8FC; }

    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0F2044 0%, #1A3A6B 100%);
        border-right: none;
    }
    [data-testid="stSidebar"] * { color: #C9D6EE !important; }
    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3 { color: #FFFFFF !important; }
    [data-testid="stSidebar"] .stMarkdown p { font-size: 0.85rem; line-height: 1.6; }

    .page-header {
        background: linear-gradient(135deg, #0F2044 0%, #1A3A6B 60%, #1E4D8C 100%);
        border-radius: 12px; padding: 32px 36px; margin-bottom: 28px;
    }
    .page-header h1 {
        font-size: 1.75rem; font-weight: 700; margin: 0 0 8px 0;
        letter-spacing: -0.02em; color: white;
    }
    .page-header p { font-size: 0.95rem; margin: 0; color: #C9D6EE; }

    .step-card {
        background: white; border-radius: 10px; padding: 22px 26px;
        margin-bottom: 16px; border: 1px solid #E2E8F0;
        box-shadow: 0 1px 4px rgba(15,32,68,0.06);
    }
    .step-label {
        font-family: 'JetBrains Mono', monospace; font-size: 0.7rem; font-weight: 600;
        color: #1A3A6B; letter-spacing: 0.12em; text-transform: uppercase; margin-bottom: 6px;
    }
    .step-title { font-size: 1.05rem; font-weight: 600; color: #0F2044; margin-bottom: 4px; }
    .step-desc  { font-size: 0.85rem; color: #64748B; margin-bottom: 14px; }

    /* ── Report card ── */
    .report-container {
        background: white; border-radius: 12px; padding: 36px 40px;
        border: 1px solid #E2E8F0;
        box-shadow: 0 2px 12px rgba(15,32,68,0.08); margin-top: 8px;
    }
    .report-container h1 {
        color: #0F2044; font-size: 1.5rem; font-weight: 700;
        border-bottom: 2px solid #1A3A6B; padding-bottom: 8px; margin-bottom: 20px;
    }
    .report-container h2 {
        color: #1A3A6B; font-size: 1.1rem; font-weight: 700;
        border-bottom: 1px solid #E2E8F0; padding-bottom: 5px; margin: 24px 0 10px 0;
    }
    .report-container h3 {
        color: #0F2044; font-size: 0.98rem; font-weight: 700; margin: 18px 0 6px 0;
    }
    .report-container p { font-size: 0.91rem; color: #1E293B; line-height: 1.65; margin: 4px 0 8px 0; }
    .report-container ul { padding-left: 20px; margin: 4px 0 10px 0; }
    .report-container li { font-size: 0.9rem; color: #1E293B; margin-bottom: 4px; line-height: 1.55; }
    .report-container strong { color: #0F2044; }
    .report-container hr { border: none; border-top: 1px solid #E2E8F0; margin: 20px 0; }

    /* Scorecard table */
    .report-container table {
        width: 100%; border-collapse: collapse; font-size: 0.86rem; margin: 14px 0 20px 0;
    }
    .report-container thead tr { background: #0F2044; }
    .report-container thead th {
        color: white; padding: 10px 12px; text-align: left;
        font-weight: 600; font-size: 0.8rem; letter-spacing: 0.03em;
    }
    .report-container tbody tr:nth-child(even) { background: #F1F5FB; }
    .report-container tbody tr:hover { background: #E4ECF7; }
    .report-container tbody td {
        padding: 9px 12px; border-bottom: 1px solid #DCE5F0;
        vertical-align: top; line-height: 1.5; color: #1E293B;
    }

    /* Feedback blocks */
    .fb-block {
        border-radius: 0 8px 8px 0; padding: 13px 18px;
        margin: 8px 0 10px 0; border-left-width: 4px; border-left-style: solid;
    }
    .fb-strengths { background: #F0FDF4; border-left-color: #16A34A; }
    .fb-areas     { background: #FFFBEB; border-left-color: #D97706; }
    .fb-next      { background: #EFF6FF; border-left-color: #2563EB; }
    .fb-label {
        font-size: 0.72rem; font-weight: 700; letter-spacing: 0.09em;
        text-transform: uppercase; margin-bottom: 5px;
    }
    .fb-strengths .fb-label { color: #15803D; }
    .fb-areas     .fb-label { color: #B45309; }
    .fb-next      .fb-label { color: #1D4ED8; }
    .fb-block p  { margin: 0; font-size: 0.89rem; color: #1E293B; line-height: 1.6; }
    .fb-block ul { margin: 4px 0 0 0; padding-left: 18px; }
    .fb-block li { font-size: 0.89rem; color: #1E293B; margin-bottom: 3px; }

    /* Badges */
    .badge { display: inline-block; padding: 4px 12px; border-radius: 20px;
              font-size: 0.78rem; font-weight: 600; font-family: 'JetBrains Mono', monospace; }
    .badge-success { background: #DCFCE7; color: #15803D; }

    div[data-testid="stButton"] > button {
        background: linear-gradient(135deg, #1A3A6B, #1E4D8C); color: white;
        border: none; border-radius: 8px; padding: 12px 32px;
        font-family: 'Inter', sans-serif; font-weight: 600; font-size: 0.95rem;
        transition: opacity 0.2s; width: 100%;
    }
    div[data-testid="stButton"] > button:hover { opacity: 0.88; color: white; }

    div[data-testid="stDownloadButton"] > button {
        background: white; color: #1A3A6B; border: 2px solid #1A3A6B;
        border-radius: 8px; padding: 10px 28px;
        font-family: 'Inter', sans-serif; font-weight: 600; font-size: 0.92rem;
        transition: all 0.2s; width: 100%; margin-top: 8px;
    }
    div[data-testid="stDownloadButton"] > button:hover { background: #0F2044; color: white; }
    hr { border: none; border-top: 1px solid #E2E8F0; margin: 24px 0; }
    .stSpinner > div { color: #1A3A6B !important; }
</style>
""", unsafe_allow_html=True)

# ─── Session state ────────────────────────────────────────────────────────────

for key in ("report", "pdf_bytes", "rendered_html"):
    if key not in st.session_state:
        st.session_state[key] = None

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
    if PDF_BACKEND:
        st.markdown(f"- **PDF Engine:** `{PDF_BACKEND}`")
    else:
        st.markdown("- **PDF Engine:** ⚠️ None found")
    st.markdown("---")
    st.markdown(
        f"<p style='font-size:0.75rem;opacity:0.6;'>Automated Practicum Review System · {date.today().year}</p>",
        unsafe_allow_html=True
    )

# ─── Page Header ──────────────────────────────────────────────────────────────

st.markdown("""
<div class="page-header">
    <h1>📋 Practicum Review Report Generator</h1>
    <p>Upload a practicum document and SOP/rubric file — the AI will evaluate the submission
    against your criteria and generate a structured, professional review report.</p>
</div>
""", unsafe_allow_html=True)

if not GEMINI_API_KEY:
    st.error("⚠️ **Gemini API key not found.** Add `GEMINI_API_KEY` to your `.env` file or Streamlit secrets.")
    st.stop()

# ─── File helpers ─────────────────────────────────────────────────────────────

def read_pdf(path: str) -> str:
    text = ""
    with fitz.open(path) as pdf:
        for page in pdf:
            text += page.get_text() + "\n"
    return text.strip()

def read_docx(path: str) -> str:
    doc = Document(path)
    return "\n".join(p.text for p in doc.paragraphs).strip()

def extract_practicum_text(path: str) -> str:
    if path.endswith(".pdf"):  return read_pdf(path)
    if path.endswith(".docx"): return read_docx(path)
    return ""

def read_sop_excel(path: str) -> str:
    xf = pd.ExcelFile(path)
    parts = []
    for sheet in xf.sheet_names:
        df = pd.read_excel(path, sheet_name=sheet)
        parts.append(f"Sheet: {sheet}\n{df.to_string(index=False)}")
    return "\n\n".join(parts)

# ─── AI Prompt ────────────────────────────────────────────────────────────────

def generate_ai_review(practicum_text: str, sop_text: str) -> str:
    practicum_text = practicum_text[:18000]
    sop_text       = sop_text[:12000]
    today          = date.today().strftime("%d %B %Y")
    model          = genai.GenerativeModel("models/gemini-2.5-flash-lite")

    
    prompt = f"""
You are an Automated Practicum Review System.

Review the uploaded practicum, assignment, discussion, or field-work document strictly against the SOP/rubric provided.

Use formal UK English.
Keep the tone professional, balanced, diagnostic, and evaluator-friendly.
Do not assume missing information.
Do not give generic feedback.
Mention strengths first, then areas for improvement.
Give specific evidence from the uploaded document wherever possible.
The report must be deeper than a basic checklist review and should identify pedagogical, structural, assessment, language, and learner-facing issues.

Today's date is {today}.

Scoring rule:
- Score each review dimension out of 5.
- The Overall Quality Score must be the average of the 6 dimension scores.
- Calculate it using this formula:
  Overall Quality Score = total of all 6 dimension scores / 6
- Round the final score to one decimal place.
- Do not change the final score unless the dimension scores change.

SOP AND RUBRIC:
{sop_text}

PRACTICUM / ACTIVITY DOCUMENT:
{practicum_text}

OUTPUT THE REPORT IN EXACTLY THIS STRUCTURE. DO NOT DEVIATE FROM THE FORMAT.

---

# PRACTICUM REVIEW REPORT

**Document Title:** [extract or: Not specified]  
**Assignment Code:** [extract or: Not specified]  
**Course:** [extract or: Not specified]  
**Review Date:** {today}  
**Reviewer:** Automated Practicum Review System  
**Product Type:** [Assignment / Discussion / Field Work / Practicum]

---

## Executive Summary

Write 7–8 sentences covering:
- overall quality of the document
- purpose and relevance of the activity
- strongest aspects
- most important gaps
- whether the document is learner-ready
- whether learning outcomes, activity, procedure, rubric, and submission requirements are aligned
- what type of revision is most needed

---

## Overall Score

**Overall Quality Score:** X / 5  
**Overall Rating:** [Needs Improvement / Developing / Proficient / Exemplary]

---

## Scorecard

| Dimension | Score / 5 | Rating | Short Justification |
|---|---|---|---|
| Relevance and Professional Purpose | X/5 | [rating] | [one specific sentence] |
| Learner-centric Design | X/5 | [rating] | [one specific sentence] |
| Structure and Flow | X/5 | [rating] | [one specific sentence] |
| Learning Outcome Quality and Bloom’s Alignment | X/5 | [rating] | [one specific sentence] |
| Rubric and Evaluation Quality | X/5 | [rating] | [one specific sentence] |
| Clarity, Cognitive Load, Grammar and Readability | X/5 | [rating] | [one specific sentence] |

---

## Detailed Dimension-wise Feedback

CRITICAL RULE:
For every dimension, write each sub-section as a SEPARATE block:
**Strengths:**
**Areas for Improvement:**
**Suggested Next Step:**
Never combine these into one paragraph.

### Dimension 1: Relevance and Professional Purpose

**Judgement:**
Write 1–2 sentences explaining the score.

**Strengths:**
- Identify whether the purpose of the activity is clear.
- Mention whether the task connects to real teaching practice, classroom use, professional development, or learner benefit.
- Cite a specific section, phrase, or feature from the document wherever possible.

**Areas for Improvement:**
- Check whether learner benefit or WIIFM is explicit.
- Check whether real-world classroom context, professional relevance, ethical relevance, or inclusive practice is missing.
- Give specific suggestions using this format: [gap] → Suggestion: [specific improvement].

**Suggested Next Step:**
- Give one actionable revision.

### Dimension 2: Learner-centric Design

**Judgement:**
Write 1–2 sentences explaining the score.

**Strengths:**
- Comment on learner-facing language, autonomy, clarity, tone, and support for independent completion.
- Mention whether the document uses direct address such as “you will” or impersonal language such as “learners will”.

**Areas for Improvement:**
- Check whether the instructions support independent completion.
- Check whether reflection prompts are scaffolded from simple to higher-order thinking.
- Check whether the activity respects prior teaching experience, especially for in-service or trainee teachers.
- Give specific suggestions.

**Suggested Next Step:**
- Give one actionable revision.

### Dimension 3: Structure and Flow

**Judgement:**
Write 1–2 sentences explaining the score.

**Strengths:**
- Comment on the sequence of sections.
- Mention whether description, procedure, discussion, rubric, and submission guidelines are clearly separated.

**Areas for Improvement:**
- Identify missing or weak sections.
- Check whether Learning Materials Required is present.
- Check whether Procedure is step-by-step or only implied.
- Check whether Dos and Don’ts are needed.
- Check whether section transitions are abrupt.
- Check for repetition or redundancy.
- Give specific suggestions.

**Suggested Next Step:**
- Give one actionable revision.

### Dimension 4: Learning Outcome Quality and Bloom’s Alignment

**Judgement:**
Write 1–2 sentences explaining the score.

**Learning Outcome and Taxonomy Analysis:**

Create the analysis in this table format:

| Learning Outcome | Verb Used | Bloom’s Level | Issue Name | Issue Identified | Suggested Revision |
|---|---|---|---|---|---|
| [quote or paraphrase LO1] | [verb] | [Remember / Understand / Apply / Analyse / Evaluate / Create] | **[Measurability / Low Bloom’s Level / Weak Alignment / Broad Wording / Instruction-like Outcome / No Issue]** | [explain the issue clearly in one sentence] | [rewrite the LO using a stronger measurable verb] |
| [quote or paraphrase LO2] | [verb] | [Remember / Understand / Apply / Analyse / Evaluate / Create] | **[Measurability / Low Bloom’s Level / Weak Alignment / Broad Wording / Instruction-like Outcome / No Issue]** | [explain the issue clearly in one sentence] | [rewrite the LO using a stronger measurable verb] |

Rules for this table:
- Keep the Issue Name in bold.
- Do not repeat the same wording in every row.
- If no issue is found, write **No Issue** in the Issue Name column.
- Keep Suggested Revision specific and measurable.
- Use Bloom’s taxonomy accurately.

**Strengths:**
- Mention whether learning outcomes are visible, measurable, and connected to the activity.

**Areas for Improvement:**
- Check whether the verbs are at Apply level or above where needed.
- Check whether the learning outcomes align with the task, procedure, and rubric.
- Check whether any LO is actually an instruction rather than a learning outcome.
- Give specific suggestions.

**Suggested Next Step:**
- Give one actionable revision.

### Dimension 5: Rubric and Evaluation Quality

**Judgement:**
Write 1–2 sentences explaining the score.

**Strengths:**
- Comment on whether the rubric is present, clear, criterion-based, and aligned with the task.
- Mention whether marks total correctly and whether performance bands are meaningful.

**Areas for Improvement:**
- Check whether descriptors are observable and measurable.
- Identify vague terms such as “clear”, “good”, “proper”, “creative”, “appropriate”, or “well-organised” if they are not defined.
- Check whether word count, format, references, evidence, or required submission elements are assessed.
- Check whether the rubric could support consistent marking across evaluators.
- Give specific suggestions.

**Suggested Next Step:**
- Give one actionable revision.

### Dimension 6: Clarity, Cognitive Load, Grammar and Readability

**Judgement:**
Write 1–2 sentences explaining the score.

**Readability Observations:**
- Estimated reading difficulty: [Easy / Moderate / Difficult]
- Sentence length: [Short / Moderate / Long]
- Passive voice: [Low / Moderate / High]
- Cognitive load: [Low / Moderate / High]

**Grammar and Language Issues:**
Write 3-5 bullet points.
For each point, use:
- Name: [grammar / punctuation / spelling / vocabulary / sentence structure]
- Example from text: [short example or write “No major issue detected”]
- Suggested correction: [specific correction or “No correction required”]
Rules:
- Keep the name in bold.
- Use specific issue names such as **Inconsistent terminology**, **Dense sentence**, **Passive construction**, **Unclear learner instruction**, **Long sentence**, **Vocabulary inconsistency**, **No major grammar issue**, or **Punctuation consistency**.
- Do not create false issues just to fill the section.
- If no serious grammar issue is found, still comment on readability, vocabulary consistency, sentence density, or terminology use.
- Keep the points readable and diagnostic, not overly long.
- Pick issues directly from the uploaded activity document.

**Vocabulary Consistency:**
- Comment on consistency of terms such as learner, candidate, student, teacher, participant, assignment, practicum, activity, discussion.
- Mention whether terminology should be standardised.

**Strengths:**
- Mention any strengths in language, tone, chunking, grammar, or UK English usage.

**Areas for Improvement:**
- Mention dense paragraphs, long sentences, unclear instructions, unexplained jargon, or inconsistent terminology if present.

**Suggested Next Step:**
- Give one actionable revision.

---

## Priority Action List

List 5 priority improvements in order of importance.

Each point must include:
- issue
- why it matters
- what should be done

Format:
1. **[Issue title]** — [issue + why it matters + action needed]
2. **[Issue title]** — [issue + why it matters + action needed]

---

## Final Recommendation

Choose one:
**Approved / Approved with Minor Improvements / Needs Revision / Major Revision Required**

Write 3–4 sentences explaining the recommendation.

---

## Review Findings Summary

Write a concise final summary with:
- overall score
- overall quality level
- top strengths
- top improvement areas
- whether the document is ready for learner use

"""

    response = model.generate_content(
        prompt,
        generation_config=genai.types.GenerationConfig(
            temperature=0,
            top_p=0.1,
            top_k=1,
            max_output_tokens=8192
        )
    )

    return response.text
    return response.text

# ─── HTML renderer for Streamlit display ─────────────────────────────────────

def render_report_html(text: str) -> str:
    """
    Convert the raw markdown report into styled HTML for Streamlit display.
    Handles: tables, feedback sub-blocks (Strengths/Areas/Next Step), headings, bullets.
    """
    lines = text.split("\n")
    out   = []
    i     = 0

    def escape(s):
        return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    def inline_md(s):
        # Bold
        s = re.sub(r"\*\*(.*?)\*\*", r"<strong>\1</strong>", s)
        s = re.sub(r"__(.*?)__",     r"<strong>\1</strong>", s)
        # Italic
        s = re.sub(r"\*(.*?)\*",     r"<em>\1</em>", s)
        return s

    while i < len(lines):
        raw  = lines[i]
        line = raw.rstrip()

        # blank
        if line.strip() == "":
            i += 1
            continue

        # HR
        if re.match(r"^\s*-{3,}\s*$", line):
            out.append("<hr>")
            i += 1
            continue

        # Markdown table — collect all consecutive | lines
        if line.strip().startswith("|"):
            table_lines = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                table_lines.append(lines[i].rstrip())
                i += 1
            # filter separator rows
            rows = [l for l in table_lines if not re.match(r"^\s*\|[\s\-:|]+\|\s*$", l)]
            if rows:
                out.append('<table>')
                for ridx, row in enumerate(rows):
                    cells = [c.strip() for c in row.strip().strip("|").split("|")]
                    if ridx == 0:
                        out.append("<thead><tr>" +
                            "".join(f"<th>{escape(c)}</th>" for c in cells) +
                            "</tr></thead><tbody>")
                    else:
                        out.append("<tr>" +
                            "".join(f"<td>{inline_md(escape(c))}</td>" for c in cells) +
                            "</tr>")
                out.append("</tbody></table>")
            continue

        # H1
        if re.match(r"^# ", line):
            out.append(f"<h1>{inline_md(escape(line[2:].strip()))}</h1>")
            i += 1
            continue

        # H2
        if re.match(r"^## ", line):
            out.append(f"<h2>{inline_md(escape(line[3:].strip()))}</h2>")
            i += 1
            continue

        # H3 — criterion heading; collect its sub-blocks
        if re.match(r"^### ", line):
            heading = inline_md(escape(line[4:].strip()))
            out.append(f"<h3>{heading}</h3>")
            i += 1

            # Collect lines belonging to this criterion block
            current_block_type = None   # "strengths" | "areas" | "next"
            current_block_items = []

            def flush_block():
                if current_block_type is None:
                    return
                cls_map   = {"strengths": "fb-strengths", "areas": "fb-areas", "next": "fb-next"}
                label_map = {
                    "strengths": "✅ Strengths",
                    "areas":     "⚠️ Areas for Improvement",
                    "next":      "→ Suggested Next Step",
                }
                cls   = cls_map[current_block_type]
                label = label_map[current_block_type]
                items_html = "".join(f"<li>{inline_md(escape(it))}</li>" for it in current_block_items)
                out.append(
                    f'<div class="fb-block {cls}">'
                    f'<div class="fb-label">{label}</div>'
                    f'<ul>{items_html}</ul>'
                    f'</div>'
                )

            while i < len(lines):
                bline = lines[i].rstrip()

                # Stop when we hit the next criterion/section
                if (re.match(r"^### ", bline) or re.match(r"^## ", bline) or
                        re.match(r"^# ", bline) or re.match(r"^\s*-{3,}\s*$", bline)):
                    break

                stripped = bline.strip()

                if stripped == "":
                    i += 1
                    continue

                # Sub-section heading detection (bold labels)
                clean_stripped = re.sub(r"\*\*(.*?)\*\*", r"\1", stripped)

                if clean_stripped.startswith("Strengths:"):
                    flush_block()
                    current_block_type  = "strengths"
                    current_block_items = []
                    i += 1
                    continue
                if clean_stripped.startswith("Areas for Improvement:"):
                    flush_block()
                    current_block_type  = "areas"
                    current_block_items = []
                    i += 1
                    continue
                if clean_stripped.startswith("Suggested Next Step:"):
                    flush_block()
                    current_block_type  = "next"
                    current_block_items = []
                    i += 1
                    continue

                # Bullet under current block
                if (stripped.startswith("- ") or stripped.startswith("* ")):
                    body = re.sub(r"^[-*]\s+", "", stripped)
                    if current_block_type:
                        current_block_items.append(body)
                    else:
                        out.append(f'<p style="margin:2px 0 2px 16px">• {inline_md(escape(body))}</p>')
                    i += 1
                    continue

                # Plain text — if inside a block treat as item, else paragraph
                if current_block_type:
                    current_block_items.append(stripped)
                else:
                    out.append(f"<p>{inline_md(escape(stripped))}</p>")
                i += 1

            flush_block()
            continue

        # Bullet list item (outside criterion blocks)
        if re.match(r"^\s*[-*] ", line):
            # Collect consecutive bullets
            out.append("<ul>")
            while i < len(lines) and re.match(r"^\s*[-*] ", lines[i]):
                body = re.sub(r"^\s*[-*]\s+", "", lines[i].rstrip())
                out.append(f"<li>{inline_md(escape(body))}</li>")
                i += 1
            out.append("</ul>")
            continue

        # Plain paragraph
        out.append(f"<p>{inline_md(escape(line.strip()))}</p>")
        i += 1

    return "\n".join(out)

# ─── PDF: WeasyPrint backend ──────────────────────────────────────────────────

PDF_CSS = """
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');

@page {
    size: A4;
    margin: 16mm 14mm 18mm 14mm;
    @bottom-center {
        content: "Automated Practicum Review System  ·  Page " counter(page) " of " counter(pages);
        font-family: Arial, sans-serif;
        font-size: 8px;
        color: #94A3B8;
    }
}

body {
    font-family: Arial, sans-serif;
    background: #ffffff;
    color: #1F2937;
    font-size: 10.5px;
    line-height: 1.6;
    margin: 0;
    padding: 0;
}

/* ── Cover strip ── */
.cover-strip {
    background: #0F2044;
    color: white;
    padding: 18px 20px 14px 20px;
    margin: -16mm -14mm 20px -14mm;
}
.cover-strip h1 {
    color: white;
    font-size: 18px;
    font-weight: 700;
    margin: 0 0 4px 0;
    border: none;
    padding: 0;
}
.cover-strip .subtitle {
    color: #93A8C8;
    font-size: 9px;
    margin: 0;
}

h1 {
    color: #0F2044;
    font-size: 17px;
    font-weight: 700;
    border-bottom: 2px solid #0F2044;
    padding-bottom: 7px;
    margin: 20px 0 14px 0;
}

h2 {
    color: #1A3A6B;
    font-size: 13px;
    font-weight: 700;
    border-bottom: 1px solid #D8DEE9;
    padding-bottom: 5px;
    margin: 20px 0 10px 0;
    page-break-after: avoid;
}

h3 {
    color: #0F2044;
    font-size: 11px;
    font-weight: 700;
    margin: 16px 0 6px 0;
    page-break-after: avoid;
}

p  { margin: 4px 0 8px 0; }
hr { border: none; border-top: 1px solid #D8DEE9; margin: 14px 0; }

strong { color: #0F2044; font-weight: 700; }

ul, ol { margin: 4px 0 10px 0; padding-left: 20px; }
li     { margin-bottom: 4px; }

/* ── Scorecard table ── */
table {
    width: 100%;
    border-collapse: collapse;
    margin: 10px 0 16px 0;
    font-size: 9px;
    page-break-inside: auto;
}
thead tr { background-color: #0F2044; }
thead th {
    color: white;
    font-weight: 700;
    padding: 7px 8px;
    text-align: left;
    border: 1px solid #0F2044;
}
tbody tr:nth-child(even) { background-color: #F1F5FB; }
tbody td {
    padding: 6px 8px;
    border: 1px solid #CBD5E1;
    vertical-align: top;
    line-height: 1.4;
}
/* Column widths */
table th:nth-child(1), table td:nth-child(1) { width: 4%;  text-align: center; }
table th:nth-child(2), table td:nth-child(2) { width: 22%; }
table th:nth-child(3), table td:nth-child(3) { width: 9%;  text-align: center; }
table th:nth-child(4), table td:nth-child(4) { width: 12%; }
table th:nth-child(5), table td:nth-child(5) { width: 53%; }

/* ── Feedback blocks ── */
.fb-block {
    border-left-width: 4px;
    border-left-style: solid;
    border-radius: 0 6px 6px 0;
    padding: 10px 14px;
    margin: 7px 0 9px 0;
    page-break-inside: avoid;
}
.fb-strengths { background: #F0FDF4; border-left-color: #16A34A; }
.fb-areas     { background: #FFFBEB; border-left-color: #D97706; }
.fb-next      { background: #EFF6FF; border-left-color: #2563EB; }

.fb-label {
    font-size: 8px;
    font-weight: 700;
    letter-spacing: 0.09em;
    text-transform: uppercase;
    margin-bottom: 5px;
}
.fb-strengths .fb-label { color: #15803D; }
.fb-areas     .fb-label { color: #B45309; }
.fb-next      .fb-label { color: #1D4ED8; }
.fb-block ul  { margin: 0; padding-left: 16px; }
.fb-block li  { font-size: 9.5px; margin-bottom: 3px; line-height: 1.45; }
"""


def html_to_pdf_weasyprint(html_body: str) -> bytes:
    full_html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<style>{PDF_CSS}</style>
</head>
<body>
<div class="cover-strip">
  <h1>Practicum Review Report</h1>
  <p class="subtitle">Automated Practicum Review System &nbsp;·&nbsp; Generated {date.today().strftime('%d %B %Y')}</p>
</div>
{html_body}
</body>
</html>"""
    buf = BytesIO()
    WeasyprintHTML(string=full_html).write_pdf(buf)
    return buf.getvalue()


# ─── PDF: ReportLab backend ───────────────────────────────────────────────────

def html_to_pdf_reportlab(html_body: str, raw_markdown: str) -> bytes:
    """
    ReportLab fallback. Parses the markdown directly (not HTML) for reliability.
    """
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, HRFlowable,
        Table, TableStyle, KeepTogether, ListFlowable, ListItem
    )
    from reportlab.lib.enums import TA_LEFT, TA_CENTER

    buf = BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=14*mm, rightMargin=14*mm,
        topMargin=16*mm, bottomMargin=18*mm,
    )

    NAV   = colors.HexColor("#0F2044")
    BLUE  = colors.HexColor("#1A3A6B")
    DARK  = colors.HexColor("#1F2937")
    G_BG  = colors.HexColor("#F1F5FB")
    WHITE = colors.white

    base = getSampleStyleSheet()
    def S(name, **kw):
        return ParagraphStyle(name, parent=base["Normal"], **kw)

    sty = {
        "h1":   S("h1",   fontSize=15, fontName="Helvetica-Bold", textColor=NAV,  spaceAfter=8, spaceBefore=14),
        "h2":   S("h2",   fontSize=12, fontName="Helvetica-Bold", textColor=BLUE, spaceAfter=6, spaceBefore=14),
        "h3":   S("h3",   fontSize=10, fontName="Helvetica-Bold", textColor=NAV,  spaceAfter=4, spaceBefore=10),
        "body": S("body", fontSize=9,  fontName="Helvetica",      textColor=DARK, spaceAfter=4, leading=14),
        "meta": S("meta", fontSize=9,  fontName="Helvetica",      textColor=DARK, spaceAfter=3, leading=13),
        "bul":  S("bul",  fontSize=9,  fontName="Helvetica",      textColor=DARK, spaceAfter=2, leading=13, leftIndent=12, bulletIndent=0),
        "fblbl":S("fblbl",fontSize=7,  fontName="Helvetica-Bold", spaceAfter=3,   leading=10),
        "fbitem":S("fbitem",fontSize=9,fontName="Helvetica",       textColor=DARK, spaceAfter=2, leading=13, leftIndent=10),
    }

    def clean(s):
        s = re.sub(r"\*\*(.*?)\*\*", r"<b>\1</b>", s)
        s = re.sub(r"\*(.*?)\*",     r"<i>\1</i>", s)
        return s

    story = []

    # Cover
    cover_data = [["PRACTICUM REVIEW REPORT",
                   f"Automated Practicum Review System  ·  {date.today().strftime('%d %B %Y')}"]]
    cover_tbl  = Table([[
        Paragraph('<font color="white" size="14"><b>PRACTICUM REVIEW REPORT</b></font>', base["Normal"]),
        Paragraph(f'<font color="#93A8C8" size="8">Automated Practicum Review System  ·  {date.today().strftime("%d %B %Y")}</font>', base["Normal"]),
    ]], colWidths=[100*mm, 68*mm])
    cover_tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), NAV),
        ("VALIGN",     (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING",(0, 0), (-1, -1), 12),
        ("TOPPADDING", (0, 0), (-1, -1), 12),
        ("BOTTOMPADDING",(0,0),(-1,-1), 12),
    ]))
    story.append(cover_tbl)
    story.append(Spacer(1, 12))

    lines = raw_markdown.split("\n")
    i = 0

    while i < len(lines):
        raw  = lines[i].rstrip()
        line = raw.strip()

        if not line:
            story.append(Spacer(1, 3))
            i += 1
            continue

        if re.match(r"^-{3,}$", line):
            story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#D8DEE9"), spaceAfter=6, spaceBefore=6))
            i += 1
            continue

        # Table
        if line.startswith("|"):
            table_lines = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                table_lines.append(lines[i].strip())
                i += 1
            rows_raw = [l for l in table_lines if not re.match(r"^\|[\s\-:|]+\|$", l)]
            if rows_raw:
                data = []
                for ridx, row in enumerate(rows_raw):
                    cells = [c.strip() for c in row.strip("|").split("|")]
                    if ridx == 0:
                        data.append([Paragraph(f"<b>{c}</b>", S("th", fontSize=8, fontName="Helvetica-Bold", textColor=WHITE)) for c in cells])
                    else:
                        data.append([Paragraph(clean(c), S("td", fontSize=8, fontName="Helvetica", textColor=DARK, leading=11)) for c in cells])
                col_w = [8*mm, 38*mm, 16*mm, 22*mm, 84*mm]
                tbl = Table(data, colWidths=col_w, repeatRows=1)
                tbl.setStyle(TableStyle([
                    ("BACKGROUND",    (0,0), (-1,0),  NAV),
                    ("ROWBACKGROUNDS",(0,1), (-1,-1), [WHITE, G_BG]),
                    ("GRID",          (0,0), (-1,-1), 0.4, colors.HexColor("#CBD5E1")),
                    ("VALIGN",        (0,0), (-1,-1), "TOP"),
                    ("TOPPADDING",    (0,0), (-1,-1), 5),
                    ("BOTTOMPADDING", (0,0), (-1,-1), 5),
                    ("LEFTPADDING",   (0,0), (-1,-1), 5),
                    ("RIGHTPADDING",  (0,0), (-1,-1), 5),
                ]))
                story.append(tbl)
                story.append(Spacer(1, 8))
            continue

        # H1
        if re.match(r"^# ", raw):
            story.append(Paragraph(clean(line[2:]), sty["h1"]))
            story.append(HRFlowable(width="100%", thickness=1.5, color=NAV, spaceAfter=6))
            i += 1
            continue

        # H2
        if re.match(r"^## ", raw):
            story.append(Paragraph(clean(line[3:]), sty["h2"]))
            story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#D8DEE9"), spaceAfter=4))
            i += 1
            continue

        # H3 — collect its feedback sub-blocks
        if re.match(r"^### ", raw):
            story.append(Paragraph(clean(line[4:]), sty["h3"]))
            i += 1

            cur_type  = None
            cur_items = []

            FB_COLORS = {
                "strengths": (colors.HexColor("#F0FDF4"), colors.HexColor("#16A34A"), colors.HexColor("#15803D"), "✓ STRENGTHS"),
                "areas":     (colors.HexColor("#FFFBEB"), colors.HexColor("#D97706"), colors.HexColor("#B45309"), "⚠ AREAS FOR IMPROVEMENT"),
                "next":      (colors.HexColor("#EFF6FF"), colors.HexColor("#2563EB"), colors.HexColor("#1D4ED8"), "→ SUGGESTED NEXT STEP"),
            }

            def flush_rl_block():
                if cur_type is None or not cur_items:
                    return
                bg, border, lbl_col, label = FB_COLORS[cur_type]
                lbl_para  = Paragraph(f'<font color="#{lbl_col.hexval()[2:].upper()}" size="7"><b>{label}</b></font>', base["Normal"])
                item_paras = [Paragraph(f"• {clean(it)}", S("fbi", fontSize=9, fontName="Helvetica", textColor=DARK, leading=13, leftIndent=8)) for it in cur_items]
                inner = [[lbl_para]] + [[p] for p in item_paras]
                inner_tbl = Table(inner, colWidths=[155*mm])
                inner_tbl.setStyle(TableStyle([
                    ("BACKGROUND",   (0,0), (-1,-1), bg),
                    ("LEFTPADDING",  (0,0), (-1,-1), 10),
                    ("RIGHTPADDING", (0,0), (-1,-1), 10),
                    ("TOPPADDING",   (0,0), (0,0),   8),
                    ("BOTTOMPADDING",(0,-1),(-1,-1),  8),
                    ("TOPPADDING",   (0,1), (-1,-1),  3),
                    ("LINEAFTER",    (0,0), (0,-1),   2, border),
                    ("LINEBEFORE",   (0,0), (0,-1),   0.5, bg),
                ]))
                story.append(KeepTogether([inner_tbl, Spacer(1, 4)]))

            while i < len(lines):
                bl = lines[i].rstrip()
                bs = bl.strip()

                if (re.match(r"^### ", bl) or re.match(r"^## ", bl) or
                        re.match(r"^# ", bl) or re.match(r"^-{3,}$", bs)):
                    break

                if not bs:
                    i += 1
                    continue

                clean_bs = re.sub(r"\*\*(.*?)\*\*", r"\1", bs)

                if clean_bs.startswith("Strengths:"):
                    flush_rl_block(); cur_type = "strengths"; cur_items = []; i += 1; continue
                if clean_bs.startswith("Areas for Improvement:"):
                    flush_rl_block(); cur_type = "areas";     cur_items = []; i += 1; continue
                if clean_bs.startswith("Suggested Next Step:"):
                    flush_rl_block(); cur_type = "next";      cur_items = []; i += 1; continue

                if bs.startswith("- ") or bs.startswith("* "):
                    body = re.sub(r"^[-*]\s+", "", bs)
                    if cur_type:
                        cur_items.append(body)
                    else:
                        story.append(Paragraph(f"• {clean(body)}", sty["bul"]))
                    i += 1
                    continue

                if cur_type:
                    cur_items.append(bs)
                else:
                    story.append(Paragraph(clean(bs), sty["body"]))
                i += 1

            flush_rl_block()
            continue

        # Bullet
        if re.match(r"^\s*[-*] ", raw):
            story.append(Paragraph(f"• {clean(re.sub(r'^[-*]\\s+', '', line))}", sty["bul"]))
            i += 1
            continue

        # Plain
        story.append(Paragraph(clean(line), sty["body"]))
        i += 1

    doc.build(story)
    return buf.getvalue()

# ─── Unified PDF entry point ──────────────────────────────────────────────────

def create_pdf_report(report_text: str) -> bytes:
    # Build the shared HTML body (used for weasyprint; reportlab parses markdown directly)
    html_body = render_report_html(report_text)

    if PDF_BACKEND == "weasyprint":
        return html_to_pdf_weasyprint(html_body)
    elif PDF_BACKEND == "reportlab":
        return html_to_pdf_reportlab(html_body, report_text)
    else:
        raise RuntimeError(
            "No PDF library available. Install weasyprint or reportlab:\n"
            "  pip install weasyprint\n  pip install reportlab"
        )

# ─── Upload UI ────────────────────────────────────────────────────────────────

col1, col2 = st.columns(2, gap="large")

with col1:
    st.markdown("""
<div class="step-card">
    <div class="step-label">Step 01</div>
    <div class="step-title">Upload Practicum Document</div>
    <div class="step-desc">Accepted formats: PDF, DOCX</div>
</div>""", unsafe_allow_html=True)
    practicum_file = st.file_uploader(
        "Practicum or assignment file", type=["pdf", "docx"],
        label_visibility="collapsed", key="practicum_upload"
    )
    if practicum_file:
        st.markdown(f'<span class="badge badge-success">✓ {practicum_file.name}</span>', unsafe_allow_html=True)

with col2:
    st.markdown("""
<div class="step-card">
    <div class="step-label">Step 02</div>
    <div class="step-title">Upload SOP / Rubric</div>
    <div class="step-desc">Accepted format: Excel (.xlsx)</div>
</div>""", unsafe_allow_html=True)
    sop_file = st.file_uploader(
        "SOP or rubric Excel file", type=["xlsx"],
        label_visibility="collapsed", key="sop_upload"
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
</div>""", unsafe_allow_html=True)

if st.button("⚡  Generate AI Review", use_container_width=True):
    if not practicum_file:
        st.error("Please upload a practicum or assignment document before generating.")
    elif not sop_file:
        st.error("Please upload the SOP / rubric Excel file before generating.")
    elif not PDF_BACKEND:
        st.error("No PDF library found. Run: `pip install weasyprint` or `pip install reportlab`")
    else:
        with tempfile.NamedTemporaryFile(
            delete=False, suffix=os.path.splitext(practicum_file.name)[1]
        ) as tp:
            tp.write(practicum_file.getbuffer())
            practicum_path = tp.name

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as ts:
            ts.write(sop_file.getbuffer())
            sop_path = ts.name

        try:
            with st.spinner("Extracting text from uploaded files…"):
                practicum_text = extract_practicum_text(practicum_path)
                sop_text       = read_sop_excel(sop_path)

            if not practicum_text:
                st.error("Could not extract text. Ensure the file is not a scanned image.")
                st.stop()

            with st.spinner("Analysing submission and generating review…"):
                report = generate_ai_review(practicum_text, sop_text)

            with st.spinner("Building PDF report…"):
                pdf_bytes     = create_pdf_report(report)
                rendered_html = render_report_html(report)

            st.session_state.report        = report
            st.session_state.pdf_bytes     = pdf_bytes
            st.session_state.rendered_html = rendered_html
            st.success("✅  Review generated successfully.")

        except Exception as err:
            st.error(f"An error occurred: `{err}`")

        finally:
            for path in [practicum_path, sop_path]:
                try: os.unlink(path)
                except OSError: pass

# ─── Report Display & Download ────────────────────────────────────────────────

if st.session_state.rendered_html:
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

    st.markdown(
        f'<div class="report-container">{st.session_state.rendered_html}</div>',
        unsafe_allow_html=True
    )

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