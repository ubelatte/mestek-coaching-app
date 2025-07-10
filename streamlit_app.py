import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import openai
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from io import BytesIO
import smtplib
from email.message import EmailMessage
import datetime

# === PASSWORD GATE ===
st.title("🔐 Secure Access")
PASSWORD = "WFHQmestek413"
if st.text_input("Enter password", type="password") != PASSWORD:
    st.warning("Access denied. Please enter the correct password.")
    st.stop()
st.success("Access granted!")

# === SETUP ===
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
service_account_info = st.secrets["gcp_service_account"]
creds = Credentials.from_service_account_info(service_account_info, scopes=scope)
client = gspread.authorize(creds)
SHEET_NAME = "Automated Supervisor Report"
sheet = client.open(SHEET_NAME).sheet1

# Correct OpenAI key loading per your secrets.toml layout
openai.api_key = st.secrets["openai"]["api_key"]

# Email credentials
SENDER_EMAIL = st.secrets["sender_email"]["sender_email"]
SENDER_PASSWORD = st.secrets["sender_password"]["sender_password"]

# === QUESTIONS ===
categories = [
    "Feedback & Conflict Resolution",
    "Communication & Team Support",
    "Reliability & Productivity",
    "Adaptability & Quality Focus",
    "Safety Commitment",
    "Documentation & Procedures"
]

prompts = [
    "How does this employee typically respond to feedback — especially when it differs from their own opinion? Do they apply it constructively, and do they help others do the same when it comes to resolving conflict and promoting cooperation?",
    "How effectively does this employee communicate with others? How well does this employee support their team - including their willingness to shift focus, assist other teams, or go beyond their assigned duties?",
    "How reliable is this employee in terms of attendance and use of time? Does this employee consistently meet or exceed productivity standards, follow company policies, and actively contribute ideas for improving standard work?",
    "When your team encounters workflow disruptions or shifting priorities, how does this employee typically respond? How does this employee contribute to maintaining and improving product quality?",
    "In what ways does this employee demonstrate commitment to safety and workplace organization? Can you provide an example of how they follow safety procedures and apply 5S principles (Sort, Set in Order, Shine, Standardize, Sustain) in their work area?",
    "How effectively does this employee use technical documentation and operate equipment according to established procedures? Please describe how they access and apply information (e.g., blueprints, work orders), and how confidently they handle equipment and tools in their role."
]

from openai import OpenAI

client_ai = OpenAI(api_key=st.secrets["openai"]["api_key"])

def analyze_feedback(category, response):
    prompt = f"""
Evaluate the following feedback for the category \"{category}\". Provide:
1. A rating from 1 to 5 (1 = Poor, 5 = Excellent)
2. A brief 1–2 sentence explanation

Feedback:
{response}

Respond in this format:
Rating: x/5
Explanation: your summary here.
"""
    completion = client_ai.chat.completions.create(
        model="gpt-3.5-turbo",
        messages=[
            {"role": "system", "content": "You are a performance coach generating professional ratings and summaries."},
            {"role": "user", "content": prompt}
        ],
        temperature=0.3
    )
    return completion.choices[0].message.content.strip()

# === DOCX GENERATION ===
def create_report(employee_name, supervisor_name, review_date, department, responses, ai_feedbacks):
    doc = Document()
    doc.add_heading('MESTEK – Hourly Performance Appraisal', level=1)
    doc.add_heading('Employee Information', level=2)

    info_fields = [
        ("Employee Name", employee_name),
        ("Department", department),
        ("Supervisor Name", supervisor_name),
        ("Date of Review", review_date)
    ]

    for label, value in info_fields:
        p = doc.add_paragraph()
        p.add_run(f"• {label}: ").bold = True
        p.add_run(str(value))
        p.paragraph_format.space_after = Pt(0)

    header = doc.add_heading('Core Performance Categories', level=2)
    header.runs[0].font.size = Pt(10)
    rating_description = doc.add_paragraph("1 – Poor | 2 – Needs Improvement | 3 – Meets Expectations | 4 – Exceeds Expectations | 5 – Outstanding")
    for run in rating_description.runs:
        run.font.size = Pt(8)  # Smaller font for rating description

    table = doc.add_table(rows=1, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    widths = [Inches(1.0), Inches(0.5), Inches(5.0)]  # Adjust widths: smaller middle column
    for i, width in enumerate(widths):
        for cell in table.columns[i].cells:
            cell.width = width

    hdr = table.rows[0].cells
    hdr[0].text = "Category"
    hdr[1].text = "Rating (1–5)"
    hdr[2].text = "Supervisor Comments"
    for cell in hdr:
        for p in cell.paragraphs:
            p.runs[0].bold = True

    ratings = []
    for i, cat in enumerate(categories):
        row = table.add_row().cells
        row[0].text = cat
        rating = explanation = "N/A"
        for line in ai_feedbacks[i].splitlines():
            if line.lower().startswith("rating:"):
                rating = line.split(":")[1].strip().split("/")[0]
            elif line.lower().startswith("explanation:"):
                explanation = line.split(":", 1)[1].strip()
        row[1].text = rating
        for p in row[1].paragraphs:
            for r in p.runs:
                r.font.size = Pt(9)
            p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        row[2].text = responses[i]
        try:
            ratings.append(float(rating))
        except:
            pass

    tbl = table._tbl
    for cell in tbl.iter_tcs():
        tcPr = cell.get_or_add_tcPr()
        borders = OxmlElement('w:tcBorders')
        for edge in ('top', 'left', 'bottom', 'right'):
            edge_el = OxmlElement(f'w:{edge}')
            edge_el.set(qn('w:val'), 'single')
            edge_el.set(qn('w:sz'), '6')
            edge_el.set(qn('w:space'), '0')
            edge_el.set(qn('w:color'), '000000')
            borders.append(edge_el)
        tcPr.append(borders)

    avg = round(sum(ratings)/len(ratings), 2) if ratings else 0
    summary = (
        f"The employee shows patterns that indicate coaching is needed in multiple performance areas. "
        f"Common themes include feedback resistance, limited communication, poor reliability, and low adaptability. "
        f"The overall performance score is {avg}/5."
    )

    doc.add_heading('Performance Summary', level=2)
    doc.add_paragraph(summary)

    doc.add_heading('Goals for Next Review Period', level=2)
    for i in range(1, 4):
        doc.add_paragraph(f"{i}. ____________________________________________")

    doc.add_heading('Sign-Offs', level=2)
    doc.add_paragraph("Employee Signature: ____________________      Date: __________")
    doc.add_paragraph("Supervisor Signature: __________________      Date: __________")

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# === EMAIL REPORT ===
def send_email(to_email, subject, body, attachment_bytes, attachment_filename):
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = SENDER_EMAIL
    msg['To'] = to_email
    msg.set_content(body)
    msg.add_attachment(attachment_bytes.read(), maintype='application',
                       subtype='vnd.openxmlformats-officedocument.wordprocessingml.document',
                       filename=attachment_filename)
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(SENDER_EMAIL, SENDER_PASSWORD)
        smtp.send_message(msg)

# === UI ===
st.title("Supervisor Feedback - Auto Evaluation")

tab1, tab2 = st.tabs(["Submit New Coaching Report", "View Past Coaching Reports"])

with tab1:
    with st.form("coaching_form"):
        email = st.text_input("Employee Email *")
        employee_name = st.text_input("Employee Name")
        supervisor_name = st.text_input("Supervisor Name")
        review_date = st.date_input("Date of Review", value=datetime.date.today())
        department = st.selectbox("Department", [
            "Rough In", "Paint Line", "Commercial Fabrication", "Baseboard Accessories"
        ])
        responses = [st.text_area(q) for q in prompts]
        submit_button = st.form_submit_button("Submit")

    if submit_button:
        if not email or not all(responses):
            st.warning("Please complete all fields.")
        else:
            st.info("Analyzing with AI...")
            ai_feedbacks = [analyze_feedback(cat, resp) for cat, resp in zip(categories, responses)]
            ratings = [f.splitlines()[0].split(':')[-1].split('/')[0] for f in ai_feedbacks]
            sheet.append_row([email, employee_name, supervisor_name, str(review_date), department,
                              *responses, *ratings, *ai_feedbacks])
            report = create_report(employee_name, supervisor_name, str(review_date), department, responses, ai_feedbacks)
            send_email(email, f"Coaching Report for {employee_name}",
                       "Attached is your performance report from Mestek.",
                       report, f"{employee_name}_report.docx")
            st.success("✅ Report emailed and saved successfully!")

with tab2:
    st.header("View Past Coaching Reports")
    data = sheet.get_all_values()
    if len(data) <= 1:
        st.info("No reports yet.")
    else:
        supervisors = sorted(set(row[2].strip() for row in data[1:] if row[2].strip()))
        selected = st.selectbox("Select Supervisor", ["--Select--"] + supervisors)
        if selected != "--Select--":
            filtered = [r for r in data[1:] if r[2].strip().lower() == selected.lower()]
            for i, row in enumerate(filtered, 1):
                st.markdown(f"### Report {i}")
                st.write(f"Date: {row[3]}, Department: {row[4]}")
                for j, cat in enumerate(categories):
                    st.markdown(f"**{cat}**")
                    st.write(f"- Comment: {row[5 + j]}")
                    st.write(f"- Rating: {row[5 + len(categories) + j]}/5")
                    st.write(f"- AI Summary: {row[5 + 2 * len(categories) + j]}")
                st.markdown("---")
