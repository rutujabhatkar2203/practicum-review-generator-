import streamlit as st
import os
import fitz
import pandas as pd
from docx import Document
from dotenv import load_dotenv
from openai import OpenAI


# Load API key from .env
load_dotenv()

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

if OPENAI_API_KEY:
    OPENAI_API_KEY = OPENAI_API_KEY.strip()

client = OpenAI(api_key=OPENAI_API_KEY)


st.set_page_config(
    page_title="Practicum Review Report Generator",
    layout="wide"
)

st.title("AI-Based Practicum Review Report Generator")

st.write(
    "Upload a practicum or assignment document and the SOP/rubric Excel file. "
    "The system will extract the text and generate an AI-based review report."
)


# Create folders safely
if not os.path.exists("uploads"):
    os.makedirs("uploads")

if not os.path.exists("outputs"):
    os.makedirs("outputs")


def read_pdf(file_path):
    text = ""
    pdf = fitz.open(file_path)

    for page in pdf:
        text += page.get_text() + "\n"

    return text


def read_docx(file_path):
    document = Document(file_path)
    text = ""

    for paragraph in document.paragraphs:
        text += paragraph.text + "\n"

    return text


def extract_practicum_text(file_path):
    if file_path.endswith(".pdf"):
        return read_pdf(file_path)

    elif file_path.endswith(".docx"):
        return read_docx(file_path)

    else:
        return "Unsupported practicum file format."


def read_sop_excel(file_path):
    excel_file = pd.ExcelFile(file_path)
    sop_text = ""

    for sheet_name in excel_file.sheet_names:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        sop_text += f"\n\nSheet Name: {sheet_name}\n"
        sop_text += df.to_string(index=False)

    return sop_text


def generate_ai_review(practicum_text, sop_text):
    prompt = f"""
You are an Automated Practicum Review System.

Your task is to review the uploaded practicum/activity document strictly against the SOP and rubric provided.

Use formal UK English.
Keep the tone professional, balanced, diagnostic and evaluator-friendly.
Do not assume missing information.
Do not give generic feedback.
Mention strengths first, then areas for improvement.
Check whether the document follows the SOP structure, required sections, learner-facing language, outcome alignment, rubric quality, submission guidelines and clarity standards.

SOP AND RUBRIC:
{sop_text}

PRACTICUM / ACTIVITY DOCUMENT:
{practicum_text}

Generate the report in this exact structure:

# PRACTICUM REVIEW REPORT

**Document Title:** [Extract from document]
**Assignment Code:** [Extract from document]
**Course:** [Extract from document]
**Review Date:** [Write today's date if available, otherwise write Not specified]
**Reviewer:** Automated Practicum Review System
**Product Type:** [Assignment / Discussion / Field Work / Practicum]

## Executive Summary
Write 5-8 sentences highlighting overall quality, strengths and key improvement areas.

## Overall Score
**Overall Quality Score:** X / 5
**Overall Rating:** [Poor / Needs Improvement / Satisfactory / Good / Excellent]

## Scorecard

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

## Detailed Feedback

For each criterion, write:
- Strengths
- Areas for Improvement
- Suggested Next Step

## Final Recommendation
Choose one:
Approved
Approved with Minor Improvements
Needs Revision
Major Revision Required

## Top 3 Priority Improvements
List only the three most important improvements.
"""

    response = client.responses.create(
        model="gpt-4.1-mini",
        input=prompt
    )

    return response.output_text


st.header("Step 1: Upload Practicum / Assignment Document")

practicum_file = st.file_uploader(
    "Upload practicum or assignment file",
    type=["pdf", "docx"]
)

st.header("Step 2: Upload SOP / Rubric Excel File")

sop_file = st.file_uploader(
    "Upload SOP or rubric Excel file",
    type=["xlsx"]
)


st.header("Step 3: Generate AI Review")

if st.button("Generate AI Review"):

    if not OPENAI_API_KEY:
        st.error("OpenAI API key is missing. Please add OPENAI_API_KEY in your .env file.")

    elif practicum_file is None:
        st.error("Please upload a practicum or assignment file.")

    elif sop_file is None:
        st.error("Please upload the SOP/rubric Excel file.")

    else:
        practicum_path = os.path.join("uploads", practicum_file.name)
        sop_path = os.path.join("uploads", sop_file.name)

        with open(practicum_path, "wb") as f:
            f.write(practicum_file.getbuffer())

        with open(sop_path, "wb") as f:
            f.write(sop_file.getbuffer())

        st.success("Files uploaded successfully.")

        with st.spinner("Extracting text from files..."):
            practicum_text = extract_practicum_text(practicum_path)
            sop_text = read_sop_excel(sop_path)

        st.success("Text extracted successfully.")

        with st.spinner("Generating AI review using OpenAI..."):
            report = generate_ai_review(practicum_text, sop_text)

        st.success("AI review generated successfully.")

        st.subheader("Generated Practicum Review Report")
        st.markdown(report)