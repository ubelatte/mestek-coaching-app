import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import openai
from docx import Document
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

openai.api_key = st.secrets["openai"]["api_key"]
SENDER_EMAIL = st.secrets["sender_email"]["sender_email"]
SENDER_PASSWORD = st.secrets["sender_password"]["sender_password"]

categories = [
    "Feedback & Conflict Resolution",
    "Communication & Team Support",
    "Reliability & Productivity",
    "Adaptability & Quality Focus",
    "Safety Commitment",
    "Documentation & Procedures"
]

prompts = [
    "How does this employee typically respond to feedback...",
    # Add the other prompts here
]

if 'responses' not in st.session_state:
    st.session_state.responses = [""] * len(prompts)

def analyze_feedback(category, response):
    prompt = f"Evaluate feedback for {category}. Feedback: {response}"
    # OpenAI request logic here...
    return openai.Completion.create(model="gpt-3.5-turbo", prompt=prompt, max_tokens=150).choices[0].text.strip()

# Form handling for data input
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
            ai_feedbacks = [analyze_feedback(cat, resp) for cat, resp in zip(categories, st.session_state.responses)]
            ratings = [f.splitlines()[0].split(':')[-1].split('/')[0] for f in ai_feedbacks]
            sheet.append_row([email, employee_name, supervisor_name, str(review_date), department, *st.session_state.responses, *ratings, *ai_feedbacks])
            report = create_report(employee_name, supervisor_name, str(review_date), department, st.session_state.responses, ai_feedbacks)
            send_email(email, f"Coaching Report for {employee_name}", "Attached is your performance report.", report, f"{employee_name}_report.docx")
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
