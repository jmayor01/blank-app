import os
import sys
import subprocess
from datetime import datetime
import tempfile

# --- Auto-install required packages (local use only) ---
def ensure(package):
    try:
        __import__(package)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])

for pkg in ["streamlit", "pandas", "plotly", "openpyxl", "matplotlib", "python-docx"]:
    ensure(pkg)

import streamlit as st
import pandas as pd
import plotly.express as px
import matplotlib.pyplot as plt

from docx import Document
from docx.shared import Inches


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
.stButton>button {
    border-radius: 10px;
}
</style>
""", unsafe_allow_html=True)

# --- Sidebar ---
st.sidebar.markdown('<div class="sidebar-title">📂 Task Report Analyzer</div>', unsafe_allow_html=True)

with st.sidebar.expander("📤 Upload Monthly Excel Reports", expanded=True):
    uploaded_files = st.file_uploader(
        "Upload one or more Excel files (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=True
    )

with st.sidebar.expander("🏆 Top Performer Settings", expanded=False):
    global_hide_top = st.checkbox("Hide All Top Performer Sections", value=False)


# ---------------- DOCX REPORT FUNCTION (ADDED) ----------------
def generate_docx_report(combined_df, monthly_data):
    doc = Document()

    doc.add_heading("Task Completion Analysis Report", 0)
    doc.add_paragraph(f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    doc.add_page_break()

    # ---- Overall Summary ----
    doc.add_heading("Overall Summary", level=1)
    doc.add_paragraph(f"Total Completion: {int(combined_df['Completion'].sum())}")
    doc.add_paragraph(f"Active Persons: {combined_df['Person'].nunique()}")
    top_person = combined_df.groupby("Person")["Completion"].sum().idxmax()
    doc.add_paragraph(f"Top Performer: {top_person}")

    overall = combined_df.groupby("Person")["Completion"].sum().reset_index()
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
        plt.figure(figsize=(8, 4))
        plt.bar(overall["Person"], overall["Completion"])
        plt.xticks(rotation=45, ha="right")
        plt.title("Overall Completion per Person")
        plt.tight_layout()
        plt.savefig(tmp.name)
        plt.close()
        doc.add_picture(tmp.name, width=Inches(6))

    doc.add_page_break()

    # ---- Monthly Reports ----
    for month, df_month in monthly_data.items():
        doc.add_heading(f"Monthly Report – {month}", level=1)

        doc.add_paragraph(f"Total Completion: {int(df_month['Completion'].sum())}")
        doc.add_paragraph(f"Active Persons: {df_month['Person'].nunique()}")

        summary = df_month.groupby(["Person", "Portal"])["Completion"].sum().reset_index()
        table = doc.add_table(rows=1, cols=3)
        hdr = table.rows[0].cells
        hdr[0].text = "Person"
        hdr[1].text = "Portal"
        hdr[2].text = "Completion"

        for _, r in summary.iterrows():
            row = table.add_row().cells
            row[0].text = str(r["Person"])
            row[1].text = str(r["Portal"])
            row[2].text = str(int(r["Completion"]))

        person_summary = df_month.groupby("Person")["Completion"].sum().reset_index()
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
            plt.figure(figsize=(8, 4))
            plt.bar(person_summary["Person"], person_summary["Completion"])
            plt.xticks(rotation=45, ha="right")
            plt.title(f"{month} – Completion per Person")
            plt.tight_layout()
            plt.savefig(tmp.name)
            plt.close()
            doc.add_picture(tmp.name, width=Inches(6))

        doc.add_page_break()

    file_path = f"Task_Completion_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
    doc.save(file_path)
    return file_path


# ---------------- MAIN LOGIC (ORIGINAL) ----------------
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
    monthly_data = {}

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

                if label in known_tasks and current_person and pd.notna(value):
                    records.append([month_year, portal, current_person, label, value])
                elif label and label not in known_tasks:
                    current_person = label

        df_month = pd.DataFrame(
            records,
            columns=["Month_Year", "Portal", "Person", "Task", "Completion"]
        )
        df_month["Completion"] = pd.to_numeric(df_month["Completion"], errors="coerce").fillna(0)

        monthly_data[month_year] = df_month
        all_data.append(df_month)

    combined_df = pd.concat(all_data, ignore_index=True)

    st.title("📊 Project Completion Analyzer")
    st.dataframe(combined_df, use_container_width=True)

    # ---------------- DOCX BUTTON (ADDED) ----------------
    st.divider()
    st.markdown("## 📄 Generate Report")

    if st.button("📥 Generate DOCX Report"):
        with st.spinner("Generating Word report..."):
            report_path = generate_docx_report(combined_df, monthly_data)

            with open(report_path, "rb") as f:
                st.download_button(
                    label="⬇️ Download DOCX Report",
                    data=f,
                    file_name=os.path.basename(report_path),
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

else:
    st.info("📤 Please upload one or more Excel reports to begin analysis.")
