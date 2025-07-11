import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from openai import OpenAI
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
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

# ✅ Confirm connection to sheet
try:
    sheet = client.open("Automated Supervisor Report").sheet1
    st.success("✅ Successfully connected to Google Sheet")
except Exception as e:
    st.error(f"❌ Sheet connection error: {e}")

client_openai = OpenAI(api_key=st.secrets["openai"]["api_key"])
SENDER_EMAIL = st.secrets["sender_email"]["sender_email"]
SENDER_PASSWORD = st.secrets["sender_password"]["sender_password"]

# === CATEGORIES & PROMPTS ===
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

# === FORMATTED SHEET WRITER ===
def update_formatted_sheet(email, employee_name, supervisor_name, review_date, department, responses, ratings, ai_score, ai_summary):
    timestamp = datetime.datetime.now().strftime("%m/%d/%Y %H:%M:%S")
    formatted_row = [
        timestamp,
        email,
        employee_name,
        supervisor_name,
        str(review_date),
        department,
        responses[0], ratings[0],
        responses[1], ratings[1],
        responses[2], ratings[2],
        responses[3], ratings[3],
        responses[4], ratings[4],
        responses[5], ratings[5],
        "", ratings[5],  # Placeholder column + score repeat
        ai_score,
        ai_summary,
        "✔️"
    ]
    sheet.append_row(formatted_row)
    st.success("✅ Google Form-style row saved in correct column order!")


# === AI ANALYSIS ===
def analyze_feedback(category, response):
    prompt = (
        f"You are an HR analyst. Rate the employee's response on '{category}' from 1 to 5. "
        f"Then summarize in 1-2 sentences. Format: Rating: X/5\nSummary: ...\n\nResponse: {response}"
    )
    try:
        completion = client_openai.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.7,
        )
        return completion.choices[0].message.content.strip()
    except Exception as e:
        return f"Rating: 3/5\nSummary: AI error: {e}"

def summarize_overall_feedback(employee_name, feedbacks):
    joined = "\n\n".join(feedbacks)
    prompt = (
        f"Summarize overall performance for {employee_name} based on the following evaluations.\n"
        f"Write a 2–3 sentence paragraph that highlights strengths and any improvement areas.\n"
        f"At the end, include an overall score out of 5 in the format: 'Overall performance score: X.XX/5'.\n\n{joined}"
    )
    try:
        completion = client_openai.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.7,
        )
        return completion.choices[0].message.content.strip()
    except Exception as e:
        return f"(Summary unavailable: {e})"

# === REST OF SCRIPT UNCHANGED === ...

# === REPORT GENERATOR ===
def create_report(employee, supervisor, review_date, department, categories, ratings, comments, summary):
    doc = Document()
    doc.add_heading("MESTEK – Hourly Performance Appraisal", level=1).alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    doc.add_heading("Employee Information", level=2)
    info = doc.add_paragraph()
    info.add_run(f"• Employee Name: {employee}\n")
    info.add_run(f"• Department: {department}\n")
    info.add_run(f"• Supervisor Name: {supervisor}\n")
    info.add_run(f"• Date of Review: {review_date}\n")

    doc.add_heading("Core Performance Categories", level=2)
    rating_note = doc.add_paragraph()
    run = rating_note.add_run("1 – Poor | 2 – Needs Improvement | 3 – Meets Expectations | 4 – Exceeds Expectations | 5 – Outstanding")
    run.font.size = Pt(9)
    rating_note.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

    table = doc.add_table(rows=1, cols=3)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Category'
    hdr_cells[1].text = 'Rating (1–5)'
    hdr_cells[2].text = 'Supervisor Comments'

    for row in table.rows:
        for i, cell in enumerate(row.cells):
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER if i == 1 else WD_PARAGRAPH_ALIGNMENT.LEFT

    for cat, rating, comment in zip(categories, ratings, comments):
        row_cells = table.add_row().cells
        row_cells[0].text = cat
        row_cells[1].text = str(rating)
        row_cells[2].text = comment
        row_cells[1].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    doc.add_paragraph("\nPerformance Summary", style='Heading 2')
    doc.add_paragraph(summary)

    doc.add_paragraph("\nGoals for Next Review Period", style='Heading 2')
    doc.add_paragraph("1. ________________________________")
    doc.add_paragraph("2. ________________________________")
    doc.add_paragraph("3. ________________________________")

    doc.add_paragraph("\nSign-Offs", style='Heading 2')
    doc.add_paragraph("Employee Signature: ________________    Date: ____________")
    doc.add_paragraph("Supervisor Signature: ________________  Date: ____________")

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer



# === EMAIL SENDER ===
def send_email(to_address, subject, body, attachment, filename):
    msg = EmailMessage()
    msg["From"] = SENDER_EMAIL
    msg["To"] = to_address
    msg["Subject"] = subject
    msg.set_content(body)
    msg.add_attachment(attachment.read(), maintype="application", subtype="vnd.openxmlformats-officedocument.wordprocessingml.document", filename=filename)
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(SENDER_EMAIL, SENDER_PASSWORD)
        smtp.send_message(msg)

# === SESSION INIT ===
if 'responses' not in st.session_state:
    st.session_state.responses = [""] * len(prompts)

# === MAIN FORM ===
with st.form("coaching_form"):
    email = st.text_input("Employee Email *")
    employee_name = st.text_input("Employee Name")
    supervisor_name = st.text_input("Supervisor Name")
    review_date = st.date_input("Date of Review", value=datetime.date.today())
    department = st.selectbox("Department", ["Rough In", "Paint Line", "Commercial Fabrication", "Baseboard Accessories"])

    for i, prompt in enumerate(prompts):
        st.session_state.responses[i] = st.text_area(prompt, value=st.session_state.responses[i])

    submit_button = st.form_submit_button("Submit")

    if submit_button:
        if not email or not all(st.session_state.responses):
            st.warning("Please complete all fields.")
        else:
            st.info("Analyzing with AI...")
            feedbacks = [analyze_feedback(cat, resp) for cat, resp in zip(categories, st.session_state.responses)]
            ratings = [f.splitlines()[0].split(":")[-1].split("/")[0].strip() for f in feedbacks]
            summaries = [f.split("Summary:")[-1].strip() for f in feedbacks]
            overall_summary = summarize_overall_feedback(employee_name, feedbacks)

            sheet.append_row([email, employee_name, supervisor_name, str(review_date), department,
                              *st.session_state.responses, *ratings, *summaries])

            report = create_report(employee_name, supervisor_name, str(review_date), department,
                                   categories, ratings, summaries, overall_summary)

            send_email(email, f"Coaching Report for {employee_name}",
                       "Attached is your performance report from Mestek.",
                       report, f"{employee_name}_report.docx")

            st.success("✅ Report emailed and saved successfully!")
            st.session_state.responses = [""] * len(prompts)


