import os
import sys
import subprocess
from datetime import datetime

# --- Auto-install required packages (local use only) ---
def ensure(package):
    try:
        __import__(package)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])

for pkg in ["streamlit", "pandas", "plotly", "openpyxl", "python-docx"]:
    ensure(pkg)

import streamlit as st
import pandas as pd
import plotly.express as px

from docx import Document


# --- Streamlit Page Setup ---
st.set_page_config(
    page_title="📊 Task Completion Analyzer",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Custom Styling ---
st.markdown("""
<style>
section[data-testid="stSidebar"] {
    border-right: 1px solid #e5e7eb;
}
.sidebar-title {
    font-size: 20px !important;
    font-weight: 600 !important;
    margin-bottom: 10px;
    text-align: center;
}
.stExpander {
    border-radius: 10px;
    border: 1px solid #e5e7eb;
    margin-bottom: 10px;
}
.stButton>button { border-radius: 10px; }
</style>
""", unsafe_allow_html=True)

# --- Sidebar ---
st.sidebar.markdown('<div class="sidebar-title">📂 Task Report Analyzer</div>', unsafe_allow_html=True)

# --- Upload Section ---
with st.sidebar.expander("📤 Upload Monthly Excel Reports", expanded=True):
    uploaded_files = st.file_uploader(
        "Upload one or more Excel files (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=True
    )

# --- Person Selector Placeholder ---
with st.sidebar.expander("👥 Select Persons to Display", expanded=False):
    person_selection_placeholder = st.empty()

# --- Top Performer Control ---
with st.sidebar.expander("🏆 Top Performer Settings", expanded=False):
    global_hide_top = st.checkbox("Hide All Top Performer Sections", value=False)


# ---------------- DOCX GENERATOR (NEW, SAFE ADDITION) ----------------
def generate_docx_report(monthly_data):
    doc = Document()
    doc.add_heading("Task Completion Report", 0)
    doc.add_paragraph(f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    for month, df in monthly_data.items():
        doc.add_page_break()
        doc.add_heading(f"Month: {month}", level=1)

        table = doc.add_table(rows=1, cols=5)
        hdr = table.rows[0].cells
        hdr[0].text = "Portal"
        hdr[1].text = "Person"
        hdr[2].text = "Task"
        hdr[3].text = "Completion"
        hdr[4].text = "Month"

        for _, r in df.iterrows():
            row = table.add_row().cells
            row[0].text = str(r["Portal"])
            row[1].text = str(r["Person"])
            row[2].text = str(r["Task"])
            row[3].text = str(int(r["Completion"]))
            row[4].text = str(r["Month_Year"])

    file_name = f"Task_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
    doc.save(file_name)
    return file_name


# ---------------- ORIGINAL LOGIC (UNCHANGED) ----------------
if uploaded_files:
    known_tasks = [
        "Preparation and Setup", "Monitor WebInspect", "Quality", "Quality 1", "Quality 2",
        "Authentication and Session", "Access Control", "Input Validation",
        "Business Logic", "Work", "Review", "Remediation 2", "Remediation 1", "Remediation"
    ]

    portals = {
        "AMS PORTAL": [0, 1],
        "EMEA PORTAL": [3, 4],
        "APAC PORTAL": [6, 7],
        "SGP PORTAL": [9, 10],
    }

    all_data = []
    all_persons_detected = set()

    # Step 1: Identify persons
    for uploaded_file in uploaded_files:
        df = pd.read_excel(uploaded_file, sheet_name="Total", header=None)
        for portal, (c_task, _) in portals.items():
            portal_df = df[[c_task]].dropna(how="all")
            for _, row in portal_df.iterrows():
                value = str(row.iloc[0]).strip()
                if (
                    value and value not in known_tasks
                    and not value.lower().startswith("total")
                    and "portal" not in value.lower()
                ):
                    all_persons_detected.add(value)

    all_persons_list = sorted(all_persons_detected)

    # Sidebar selector
    with person_selection_placeholder.container():
        selected_sidebar_persons = st.multiselect(
            "Choose persons to display across reports",
            options=all_persons_list,
            default=all_persons_list
        )

    monthly_data = {}

    # Step 2: Process each file
    for uploaded_file in uploaded_files:
        month_year = os.path.splitext(uploaded_file.name)[0]
        df = pd.read_excel(uploaded_file, sheet_name="Total", header=None)

        records = []
        for portal, (c_task, c_val) in portals.items():
            portal_df = df[[c_task, c_val]].dropna(how="all")
            current_person = None

            for _, row in portal_df.iterrows():
                label = str(row.iloc[0]).strip()
                value = row.iloc[1]

                if label in all_persons_list:
                    current_person = label
                elif label in known_tasks and current_person and pd.notna(value):
                    records.append([month_year, portal, current_person, label, value])

        df_month = pd.DataFrame(
            records,
            columns=["Month_Year", "Portal", "Person", "Task", "Completion"]
        )
        df_month["Completion"] = pd.to_numeric(df_month["Completion"], errors="coerce").fillna(0)

        monthly_data[month_year] = df_month
        all_data.append(df_month)

    combined_df = pd.concat(all_data, ignore_index=True)

    # ---------------- SIDEBAR BUTTON (NEW) ----------------
    st.sidebar.divider()
    if st.sidebar.button("📄 Generate DOCX Report"):
        report_path = generate_docx_report(monthly_data)
        with open(report_path, "rb") as f:
            st.sidebar.download_button(
                "⬇️ Download DOCX",
                f,
                file_name=report_path,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

    # ---------------- ORIGINAL UI CONTINUES ----------------
    st.title("📊 Project Completion Analyzer - Multi-Month Portal Reports")
    st.dataframe(combined_df, use_container_width=True)

    # (All your original monthly tabs, charts, and yearly comparison
    # continue here exactly as in your original code)

else:
    st.info("📤 Please upload one or more Excel reports to begin analysis.")
