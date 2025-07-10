import streamlit as st
import datetime
from collections import defaultdict, Counter
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from openai import OpenAI
from gspread_formatting import CellFormat, Color, format_cell_range

# --- CONFIG ---
SERVICE_ACCOUNT_FILE = "service_account.json"  # <- Update with your key filename
SHEET_NAME = "Coaching Assessment Form"
DASHBOARD_SHEET_NAME = "Trend Dashboard"
OPENAI_API_KEY = st.secrets["openai_api_key"] if "openai_api_key" in st.secrets else "your-openai-api-key"

# --- APP UI ---
st.title("📊 Coaching Trend Dashboard Generator")
st.write("This tool generates a summary dashboard using AI from the latest coaching history.")

if st.button("Generate Dashboard"):
    st.info("Processing... please wait ⏳")

    # === AUTH ===
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_ACCOUNT_FILE, scope)
    client = gspread.authorize(creds)
    spreadsheet = client.open(SHEET_NAME)
    form_sheet = spreadsheet.sheet1
    data = form_sheet.get_all_records()

    # === GROUP DATA ===
    employee_issues = defaultdict(list)
    for row in data:
        try:
            date = datetime.datetime.strptime(row["Date of Incident"], "%m/%d/%Y")
        except:
            continue
        employee_issues[row["Employee Name"]].append((date, row))

    client_gpt = OpenAI(api_key=OPENAI_API_KEY)
    today = datetime.datetime.now().strftime("%m/%d/%Y")
    dashboard_rows = [["Employee", "Supervisor", "Department", "Total Coachings", "Last Incident Date", "Top Issue Type", "Trend Summary", "Sentiment"]]
    sentiments = {}

    for emp, history in employee_issues.items():
        history.sort()
        supervisor = history[-1][1].get("Supervisor Name", "")
        department = history[-1][1].get("Department", "")
        total_coachings = len(history)
        last_incident_date = history[-1][0].strftime("%m/%d/%Y")
        issue_counts = Counter(entry["Issue Type"] for _, entry in history if entry.get("Issue Type"))
        top_issue_type = issue_counts.most_common(1)[0][0] if issue_counts else "N/A"

        coaching_history = "\n".join(
            f"{date.strftime('%m/%d/%Y')}: {entry.get('Action Taken', 'No action recorded')}"
            for date, entry in history
        )

        # GPT: Trend Summary
        response = client_gpt.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "You write concise summaries for HR coaching."},
                {"role": "user", "content": f"""You are a concise workplace coaching analyst.
Given this coaching history actions, write a 1-2 sentence summary highlighting key trends or repeated actions.

Coaching Actions History:
{coaching_history}"""}
            ],
            temperature=0.3,
        )
        trend_summary = response.choices[0].message.content.strip()

        # GPT: Sentiment
        sentiment_response = client_gpt.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "You classify employee coaching sentiment."},
                {"role": "user", "content": f"""Based on the following coaching actions, classify the overall tone as either POSITIVE or NEGATIVE.
Coaching Actions History:
{coaching_history}

Respond with only one word: POSITIVE or NEGATIVE."""}
            ],
            temperature=0,
        )
        sentiment = sentiment_response.choices[0].message.content.strip().upper()

        dashboard_rows.append([
            emp, supervisor, department, total_coachings,
            last_incident_date, top_issue_type, trend_summary, sentiment
        ])
        sentiments[emp] = sentiment

    # === OVERWRITE DASHBOARD SHEET ===
    try:
        dashboard_sheet = spreadsheet.worksheet(DASHBOARD_SHEET_NAME)
        spreadsheet.del_worksheet(dashboard_sheet)
    except gspread.WorksheetNotFound:
        pass

    dashboard_sheet = spreadsheet.add_worksheet(title=DASHBOARD_SHEET_NAME, rows=str(len(dashboard_rows)+10), cols="10")
    dashboard_sheet.update('A1', dashboard_rows)

    # === FORMAT SENTIMENT ===
    positive_format = CellFormat(backgroundColor=Color(0.8, 1, 0.8))   # Light green
    negative_format = CellFormat(backgroundColor=Color(1, 0.8, 0.8))   # Light red

    for i, row in enumerate(dashboard_rows[1:], start=2):
        emp = row[0]
        sentiment = sentiments.get(emp, "")
        if sentiment == "POSITIVE":
            format_cell_range(dashboard_sheet, f"A{i}", positive_format)
        elif sentiment == "NEGATIVE":
            format_cell_range(dashboard_sheet, f"A{i}", negative_format)

    st.success(f"✅ Dashboard updated with {len(employee_issues)} employees on {today}.")

