import streamlit as st
from docx import Document
from PyPDF2 import PdfReader
from PIL import Image
import pytesseract
from transformers import pipeline
import re
import pandas as pd
import plotly.express as px
# PDF REPORT IMPORTS
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer
)
from reportlab.lib.styles import getSampleStyleSheet
# =====================================================
# LOAD AI SUMMARIZATION MODEL
# =====================================================
summarizer = pipeline(
    "summarization",
    model="facebook/bart-large-cnn"
)
# =====================================================
# TESSERACT OCR PATH
# =====================================================
pytesseract.pytesseract.tesseract_cmd = (
    r"C:\Program Files\Tesseract-OCR\tesseract.exe"
)
# =====================================================
# PAGE CONFIGURATION
# =====================================================
st.set_page_config(
    page_title="AI Resume Analyzer",
    page_icon="📄",
    layout="wide"
)
# =====================================================
# TITLE
# =====================================================
st.title("🚀 Smart AI Resume Analyzer")
st.write("Developed by Harika")
# =====================================================
# JOB DESCRIPTION INPUT
# =====================================================
st.subheader("📌 Paste Job Description")
job_description = st.text_area(
    "Enter Job Description",
    height=200
)
# =====================================================
# FILE UPLOAD
# =====================================================
st.subheader("📄 Upload Resume")
uploaded_file = st.file_uploader(
    "Upload Resume",
    type=["pdf", "docx", "png", "jpg", "jpeg"]
)
# =====================================================
# SKILLS DATABASE
# =====================================================
skills_list = [
    "Python",
    "SQL",
    "Machine Learning",
    "Deep Learning",
    "Artificial Intelligence",
    "Data Analysis",
    "Power BI",
    "Tableau",
    "Excel",
    "AWS",
    "Azure",
    "Java",
    "C++",
    "TensorFlow",
    "NLP",
    "Pandas",
    "NumPy",
    "Streamlit",
    "Docker",
    "Spark",
    "Hadoop",
    "ETL",
    "Data Engineering",
    "GCP",
    "Databricks",
    "Hive",
    "Kafka",
    "Scala"
]
# =====================================================
# CLEAN OCR / TEXT NOISE
# =====================================================
def clean_resume_text(text):
    cleaned_lines = []
    lines = text.split("\n")
    garbage_words = [
        "OAD",
        "VA4a",
        "ssiasivse",
        "ANE",
        "RD",
        "eR",
        "EU cay"
    ]
    for line in lines:
        line = line.strip()
        if not line:
            continue
        if len(line) < 4:
            continue
        if any(word in line for word in garbage_words):
            continue
        strange_count = sum(
            1 for char in line
            if not char.isalnum()
            and char not in " .,()-:/@"
        )
        strange_ratio = strange_count / max(len(line), 1)
        if strange_ratio > 0.30:
            continue
        cleaned_lines.append(line)
    cleaned_text = "\n".join(cleaned_lines)
    cleaned_text = re.sub(r'\s+', ' ', cleaned_text)
    return cleaned_text
# =====================================================
# INTERVIEW QUESTION GENERATOR
# =====================================================
def generate_questions(skill):
    question_bank = {
        "Python": [
            "Explain list comprehension in Python.",
            "What is the difference between list and tuple?",
            "Explain decorators in Python.",
            "What are lambda functions?",
            "Explain exception handling in Python."
        ],
        "SQL": [
            "What is the difference between WHERE and HAVING?",
            "Explain joins in SQL.",
            "What are window functions?",
            "Explain normalization.",
            "What is indexing in SQL?"
        ],
        "Machine Learning": [
            "What is overfitting?",
            "Explain bias vs variance.",
            "Difference between supervised and unsupervised learning?",
            "What is cross validation?",
            "Explain precision and recall."
        ],
        "AWS": [
            "What is EC2?",
            "Explain S3 storage.",
            "Difference between EC2 and Lambda?",
            "What is IAM?",
            "Explain VPC."
        ],
        "Spark": [
            "What is Apache Spark?",
            "Difference between RDD and DataFrame?",
            "Explain Spark architecture.",
            "What is lazy evaluation?",
            "What is Spark SQL?"
        ],
        "Hadoop": [
            "Explain HDFS.",
            "What is MapReduce?",
            "Difference between Hadoop and Spark?",
            "Explain NameNode and DataNode.",
            "What is YARN?"
        ],
        "ETL": [
            "What is ETL process?",
            "Difference between ETL and ELT?",
            "Explain data transformation.",
            "What are ETL pipelines?",
            "How do you handle data validation?"
        ],
        "Power BI": [
            "What are dashboards in Power BI?",
            "Explain DAX functions.",
            "What is Power Query?",
            "Difference between calculated column and measure?",
            "How do you optimize Power BI reports?"
        ],
        "Excel": [
            "What are pivot tables?",
            "Explain VLOOKUP and XLOOKUP.",
            "What are macros?",
            "Explain conditional formatting.",
            "What are Excel formulas?"
        ]
    }
    if skill in question_bank:
        return question_bank[skill]
    else:
        return [
            f"Explain your experience with {skill}.",
            f"What are key concepts in {skill}?",
            f"What projects have you built using {skill}?",
            f"What challenges did you face in {skill}?",
            f"What best practices do you follow in {skill}?"
        ]
# =====================================================
# PDF REPORT GENERATION
# =====================================================
def generate_pdf_report(
    ats_score,
    detected_skills,
    missing_skills,
    predicted_role,
    clean_summary,
    all_questions
):
    doc = SimpleDocTemplate("AI_Resume_Report.pdf")
    styles = getSampleStyleSheet()
    elements = []
    elements.append(
        Paragraph(
            "<b>AI Resume Analysis Report</b>",
            styles['Title']
        )
    )
    elements.append(Spacer(1, 20))
    elements.append(
        Paragraph(
            f"<b>ATS Score:</b> {ats_score}%",
            styles['BodyText']
        )
    )
    elements.append(Spacer(1, 12))
    elements.append(
        Paragraph(
            f"<b>Predicted Role:</b> {predicted_role}",
            styles['BodyText']
        )
    )
    elements.append(Spacer(1, 12))
    skills_text = ", ".join(detected_skills)
    elements.append(
        Paragraph(
            f"<b>Detected Skills:</b> {skills_text}",
            styles['BodyText']
        )
    )
    elements.append(Spacer(1, 12))
    missing_text = ", ".join(missing_skills)
    elements.append(
        Paragraph(
            f"<b>Missing Skills:</b> {missing_text}",
            styles['BodyText']
        )
    )
    elements.append(Spacer(1, 12))
    elements.append(
        Paragraph(
            f"<b>AI Resume Summary:</b><br/>{clean_summary}",
            styles['BodyText']
        )
    )
    elements.append(Spacer(1, 20))
    elements.append(
        Paragraph(
            "<b>Interview Questions</b>",
            styles['Heading2']
        )
    )
    for q in all_questions:
        elements.append(
            Paragraph(f"• {q}", styles['BodyText'])
        )
    doc.build(elements)
# =====================================================
# RESUME TEXT EXTRACTION
# =====================================================
resume_text = ""
if uploaded_file is not None:
    file_type = uploaded_file.name.split(".")[-1].lower()
    # DOCX
    if file_type == "docx":
        doc = Document(uploaded_file)
        for paragraph in doc.paragraphs:
            resume_text += paragraph.text + "\n"
    # PDF
    elif file_type == "pdf":
        pdf_reader = PdfReader(uploaded_file)
        for page in pdf_reader.pages:
            text = page.extract_text()
            if text:
                resume_text += text + "\n"
    # IMAGE OCR
    elif file_type in ["png", "jpg", "jpeg"]:
        image = Image.open(uploaded_file)
        resume_text = pytesseract.image_to_string(image)
    # CLEAN TEXT
    resume_text = clean_resume_text(resume_text)
    # SUCCESS MESSAGE
    st.success("✅ Resume processed successfully!")
    # DISPLAY RESUME TEXT
    st.subheader("📄 Extracted Resume Text")
    st.text_area(
        "Resume Content",
        resume_text,
        height=300
    )
    # =====================================================
    # SKILL DETECTION
    # =====================================================
    detected_skills = []
    for skill in skills_list:
        if skill.lower() in resume_text.lower():
            detected_skills.append(skill)
    # DISPLAY SKILLS
    st.subheader("🧠 Detected Skills")
    if detected_skills:
        for skill in detected_skills:
            st.write("✅", skill)
    else:
        st.write("No skills detected")
    # =====================================================
    # ATS SCORE CALCULATION
    # =====================================================
    ats_score = 0
    matched_skills = []
    jd_skills = []
    # FIND SKILLS PRESENT IN JOB DESCRIPTION
    for skill in skills_list:
        if skill.lower() in job_description.lower():
            jd_skills.append(skill)
    # FIND MATCHED SKILLS
    for skill in detected_skills:
        if skill in jd_skills:
            matched_skills.append(skill)
    # CALCULATE ATS SCORE
    if len(jd_skills) > 0:
        ats_score = int(
            (len(matched_skills) / len(jd_skills)) * 100
        )
    # LIMIT MAX SCORE
    ats_score = min(ats_score, 100)
    # DISPLAY ATS SCORE
    st.subheader("📊 ATS Resume Score")
    st.progress(ats_score / 100)
    st.write(f"ATS Score: {ats_score}%")
    # =====================================================
    # MATCHED SKILLS
    # =====================================================
    st.subheader("✅ Matched Skills")
    if matched_skills:
        for skill in matched_skills:
            st.write("✅", skill)
    else:
        st.write("No matching skills found.")
    # =====================================================
    # SKILL MATCH VISUALIZATION
    # =====================================================
    st.subheader("📈 Skill Match Visualization")
    skill_data = []
    for skill in skills_list:
        if skill in detected_skills:
            score = 100
        else:
            score = 0
        skill_data.append({
            "Skill": skill,
            "Match": score
        })
    df = pd.DataFrame(skill_data)
    fig = px.bar(
        df,
        x="Skill",
        y="Match",
        color="Match",
        title="Skill Match Analysis"
    )
    st.plotly_chart(fig, use_container_width=True)
    # =====================================================
    # PIE CHART
    # =====================================================
    matched_count = len(detected_skills)
    missing_count = len(skills_list) - matched_count
    pie_data = pd.DataFrame({
        "Category": [
            "Matched Skills",
            "Missing Skills"
        ],
        "Count": [
            matched_count,
            missing_count
        ]
    })
    pie_fig = px.pie(
        pie_data,
        names="Category",
        values="Count",
        title="Resume Skill Distribution"
    )
    st.plotly_chart(pie_fig, use_container_width=True)
    # =====================================================
    # AI SUMMARY
    # =====================================================
    st.subheader("🤖 AI Resume Summary")
    clean_summary = ""
    try:
        summary_input = resume_text[:1200]
        result = summarizer(
            summary_input,
            max_length=80,
            min_length=30,
            do_sample=False
        )
        clean_summary = result[0]['summary_text']
        clean_summary = clean_summary.replace(
            "He ",
            "The candidate "
        )
        clean_summary = clean_summary.replace(
            "She ",
            "The candidate "
        )
        clean_summary = clean_summary.replace(
            " he ",
            " the candidate "
        )
        clean_summary = clean_summary.replace(
            " she ",
            " the candidate "
        )
        st.success(clean_summary)
    except Exception as e:
        st.error(f"Summary Error: {e}")
    # =====================================================
    # PREDICTED ROLE
    # =====================================================
    st.subheader("🎯 Predicted Role")
    predicted_role = "Software Developer"
    if "Machine Learning" in detected_skills:
        predicted_role = "Machine Learning Engineer"
    elif "Data Analysis" in detected_skills:
        predicted_role = "Data Analyst"
    elif "Power BI" in detected_skills:
        predicted_role = "Business Intelligence Analyst"
    elif "AWS" in detected_skills:
        predicted_role = "Cloud Engineer"
    elif "Data Engineering" in detected_skills:
        predicted_role = "Data Engineer"
    elif "Spark" in detected_skills:
        predicted_role = "Big Data Engineer"
    st.success(predicted_role)
    # =====================================================
    # MISSING SKILLS
    # =====================================================
    st.subheader("❌ Missing Skills")
    missing_skills = []
    for skill in jd_skills:
        if skill not in detected_skills:
            missing_skills.append(skill)
    if missing_skills:
        for skill in missing_skills:
            st.write("❌", skill)
    else:
        st.write("No major missing skills detected.")
    # =====================================================
    # INTERVIEW QUESTIONS
    # =====================================================
    st.subheader("🎯 AI Generated Interview Questions")
    all_questions = []
    for skill in detected_skills:
        st.write(f"## {skill}")
        generated_questions = generate_questions(skill)
        for q in generated_questions:
            st.write("•", q)
            all_questions.append(q)
    # =====================================================
    # GENERATE PDF REPORT
    # =====================================================
    generate_pdf_report(
        ats_score,
        detected_skills,
        missing_skills,
        predicted_role,
        clean_summary,
        all_questions
    )
    # =====================================================
    # DOWNLOAD PDF BUTTON
    # =====================================================
    with open("AI_Resume_Report.pdf", "rb") as pdf_file:
        st.download_button(
            label="📥 Download Resume Analysis Report",
            data=pdf_file,
            file_name="AI_Resume_Report.pdf",
            mime="application/pdf"
        )