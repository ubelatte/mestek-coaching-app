import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from openai import OpenAI  # ✅ new SDK usage
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
client_gsheets = gspread.authorize(creds)
SHEET_NAME = "Automated Supervisor Report"
sheet = client_gsheets.open(SHEET_NAME).sheet1

client_openai = OpenAI(api_key=st.secrets["openai"]["api_key"])  # ✅ new client

SENDER_EMAIL = st.secrets["sender_email"]["sender_email"]
SENDER_PASSWORD = st.secrets["sender_password"]["sender_password"]

# === TEST GPT CONNECTION ===
if st.checkbox("Test GPT response"):
    try:
        test = client_openai.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": "Say hello from Streamlit"}]
        )
        st.success(test.choices[0].message.content)
    except Exception as e:
        st.error(f"❌ GPT error: {e}")

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

# === ANALYSIS FUNCTION ===
def analyze_feedback(category, response):
    prompt = (
        f"You are an HR performance analyst. Rate the following employee comment related to '{category}' "
        f"on a scale of 1 to 5, and explain why. Provide your output in this format:\n\n"
        f"Rating: X/5\nSummary: ...\n\n"
        f"Comment:\n{response}"
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

# === REPORT GENERATOR ===
def create_report(employee, supervisor, review_date, department, responses, ai_feedbacks):
    doc = Document()
    doc.add_heading(f'Coaching Report: {employee}', 0)

    table = doc.add_table(rows=1, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Category'
    hdr_cells[1].text = 'Rating'
    hdr_cells[2].text = 'Explanation'

    for category, ai_result in zip(categories, ai_feedbacks):
        lines = ai_result.splitlines()
        rating = next((line for line in lines if "Rating" in line), "Rating: N/A")
        summary = next((line for line in lines if "Summary" in line), "Summary: N/A")
        row_cells = table.add_row().cells
        row_cells[0].text = category
        row_cells[1].text = rating.replace("Rating:", "").strip()
        row_cells[2].text = summary.replace("Summary:", "").strip()

    doc.add_paragraph("\nDevelopment Goals:")
    doc.add_paragraph("1. ____________________________________", style='List Number')
    doc.add_paragraph("2. ____________________________________", style='List Number')
    doc.add_paragraph("3. ____________________________________", style='List Number')
    doc.add_paragraph("\nEmployee Signature: ____________________________    Date: ____________")
    doc.add_paragraph("Supervisor Signature: __________________________  Date: ____________")

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

# === SESSION STATE INIT ===
if 'responses' not in st.session_state:
    st.session_state.responses = [""] * len(prompts)

# === MAIN FORM ===
with st.form("coaching_form"):
    email = st.text_input("Employee Email *")
    employee_name = st.text_input("Employee Name")
    supervisor_name = st.text_input("Supervisor Name")
    review_date = st.date_input("Date of Review", value=datetime.date.today())
    department = st.selectbox("Department", [
        "Rough In", "Paint Line", "Commercial Fabrication", "Baseboard Accessories"
    ])

    for i, prompt in enumerate(prompts):
        st.session_state.responses[i] = st.text_area(prompt, value=st.session_state.responses[i])

    submit_button = st.form_submit_button("Submit")

    if submit_button:
        if not email or not all(st.session_state.responses):
            st.warning("Please complete all fields.")
        else:
            st.info("Analyzing with AI...")
            ai_feedbacks = [analyze_feedback(cat, resp) for cat, resp in zip(categories, st.session_state.responses)]
            ratings = [f.splitlines()[0].split(':')[-1].split('/')[0] for f in ai_feedbacks]

            sheet.append_row([email, employee_name, supervisor_name, str(review_date), department,
                              *st.session_state.responses, *ratings, *ai_feedbacks])

            report = create_report(employee_name, supervisor_name, str(review_date), department, st.session_state.responses, ai_feedbacks)
            send_email(email, f"Coaching Report for {employee_name}",
                       "Attached is your performance report from Mestek.",
                       report, f"{employee_name}_report.docx")

            st.success("✅ Report emailed and saved successfully!")
            st.session_state.responses = [""] * len(prompts)

# === VIEW PAST REPORTS ===
with st.expander("View Past Coaching Reports"):
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
