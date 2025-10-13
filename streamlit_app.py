import os
import sys
import subprocess

# --- Auto-install required packages ---
def ensure(package):
    try:
        __import__(package)
    except ImportError:
        print(f"[INFO] Installing {package} ...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])

for pkg in ["streamlit", "pandas", "plotly"]:
    ensure(pkg)

import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime

# --- Streamlit Page Setup ---
st.set_page_config(page_title="📊 Task Completion Analyzer", layout="wide")
st.markdown(
    """
    <style>
    /* --- Custom Styling --- */
    h1, h2, h3, h4 { color: #0A4D68; font-weight: 700; }
    .stDataFrame, .stPlotlyChart { border-radius: 10px !important; }
    .metric-container { background-color: #f0f9ff; padding: 10px; border-radius: 10px; }
    .stButton>button { border-radius: 10px; background-color: #007ACC; color: white; }
    .stButton>button:hover { background-color: #005EA6; }
    .stTabs [data-baseweb="tab-list"] { gap: 6px; }
    .stTabs [data-baseweb="tab"] { border-radius: 8px; padding: 6px 12px; background: #f4f4f4; color: #333; }
    .stTabs [aria-selected="true"] { background: #007ACC; color: white; font-weight: 600; }
    </style>
    """,
    unsafe_allow_html=True
)

st.title("📊 Project Completion Analyzer - Multi-Month Portal Reports")
st.caption("Analyze completion performance across multiple portals and months effortlessly.")

# --- File Upload ---
st.markdown("### 📤 Upload Monthly Excel Reports")
uploaded_files = st.file_uploader(
    "Upload one or more Excel files (.xlsx) containing monthly task reports",
    type=["xlsx"],
    accept_multiple_files=True
)

if uploaded_files:
    known_tasks = [
        "Preparation and Setup", "Monitor WebInspect", "Quality", "Quality 1", "Quality 2",
        "Authentication and Session", "Access Control", "Input Validation", "Business Logic",
        "Work", "Review", "Remediation 2", "Remediation 1", "Remediation"
    ]

    portals = {
        "AMS PORTAL": [0, 1],
        "EMEA PORTAL": [3, 4],
        "APAC PORTAL": [6, 7],
        "SGP PORTAL": [9, 10],
    }

    all_data = []
    all_persons_detected = set()

    # --- Step 1: Capture all person names first ---
    for uploaded_file in uploaded_files:
        try:
            df = pd.read_excel(uploaded_file, sheet_name="Total", header=None)
            for portal, (col_task, col_value) in portals.items():
                portal_df = df[[col_task, col_value]].dropna(how="all")
                for idx, row in portal_df.iterrows():
                    value = str(row.iloc[0]).strip()
                    # --- Skip invalid "person-like" entries ---
                    if (
                        value
                        and value not in known_tasks
                        and value.lower() != "row labels"
                        and "portal" not in value.lower()
                        and not value.lower().startswith("total")
                        and "grand total" not in value.lower()
                        and not value.replace(" ", "").isdigit()
                    ):
                        all_persons_detected.add(value)
        except Exception as e:
            st.warning(f"⚠️ Error scanning {uploaded_file.name}: {e}")

    all_persons_list = sorted(list(all_persons_detected))

    # --- Step 2: Process data for each file using detected person list ---
    for uploaded_file in uploaded_files:
        file_name = uploaded_file.name
        month_year = os.path.splitext(file_name)[0]

        try:
            df = pd.read_excel(uploaded_file, sheet_name="Total", header=None)
            all_records = []

            for portal, (col_task, col_value) in portals.items():
                try:
                    portal_df = df[[col_task, col_value]].dropna(how="all")
                    current_person = None

                    for _, row in portal_df.iterrows():
                        cell_value = str(row.iloc[0]).strip()
                        completion_value = row.iloc[1]

                        if not cell_value or cell_value.lower() in ["total", "grand total"]:
                            continue

                        if cell_value in all_persons_list:
                            current_person = cell_value
                        elif cell_value in known_tasks and current_person and pd.notna(completion_value):
                            all_records.append([month_year, portal, current_person, cell_value, completion_value])
                except Exception as e:
                    st.warning(f"Skipping portal {portal} in {file_name}: {e}")

            if all_records:
                df_records = pd.DataFrame(all_records, columns=["Month_Year", "Portal", "Person", "Task", "Completion"])
                df_records["Completion"] = pd.to_numeric(df_records["Completion"], errors="coerce").fillna(0)
                all_data.append(df_records)
            else:
                st.warning(f"📄 No valid data found in {file_name}. Verify format.")

        except Exception as e:
            st.error(f"❌ Error reading {file_name}: {e}")

    # --- Step 3: Combine all data and display ---
    if all_data:
        combined_df = pd.concat(all_data, ignore_index=True)

        def parse_month_year(m):
            try:
                return datetime.strptime(m, "%B %Y")
            except:
                return None

        combined_df["Month_Order"] = combined_df["Month_Year"].apply(parse_month_year)
        combined_df = combined_df.sort_values("Month_Order")
        month_list = combined_df["Month_Year"].unique().tolist()

        month_tabs = st.tabs(list(month_list) + ["📈 Yearly Comparison"])
        monthly_data = {}

        # --- Monthly Tabs ---
        for i, month_year in enumerate(month_list):
            with month_tabs[i]:
                st.markdown(f"## 📅 {month_year} Summary")

                df_month = combined_df[combined_df["Month_Year"] == month_year]
                total_completion = int(df_month["Completion"].sum())
                active_persons = df_month["Person"].nunique()
                top_performer = (
                    df_month.groupby("Person")["Completion"].sum().idxmax()
                    if not df_month.empty else "N/A"
                )

                with st.container():
                    hide_top = st.checkbox(f"🙈 Hide Top Performer ({month_year})", value=False)
                    c1, c2, c3 = st.columns(3)
                    with c1:
                        st.metric("✅ Total Completions", f"{total_completion}")
                    with c2:
                        st.metric("👥 Active Persons", f"{active_persons}")
                    with c3:
                        if not hide_top:
                            st.metric("🏆 Top Performer", top_performer)
                        else:
                            st.empty()

                st.markdown("---")
                st.subheader("🎯 Select Persons to Display")
                col1, col2 = st.columns([1, 5])
                with col1:
                    select_all = st.button(f"✅ Select All ({month_year})")
                    deselect_all = st.button(f"❌ Deselect All ({month_year})")

                if "selected_persons" not in st.session_state:
                    st.session_state.selected_persons = all_persons_list

                if select_all:
                    st.session_state.selected_persons = all_persons_list
                elif deselect_all:
                    st.session_state.selected_persons = []

                selected_persons = st.multiselect(
                    f"Choose persons to include for {month_year}:",
                    options=all_persons_list,
                    default=st.session_state.selected_persons
                )
                st.session_state.selected_persons = selected_persons
                df_filtered = df_month[df_month["Person"].isin(selected_persons)]

                with st.expander("👤 Task Completion Summary per Person and Portal", expanded=True):
                    st.dataframe(
                        df_filtered.groupby(["Person", "Portal"])["Completion"]
                        .sum()
                        .reset_index(),
                        use_container_width=True
                    )

                st.plotly_chart(
                    px.bar(
                        df_filtered.groupby(["Person", "Portal"])["Completion"].sum().reset_index(),
                        x="Person",
                        y="Completion",
                        color="Portal",
                        barmode="group",
                        title=f"Task Completions per Person ({month_year})",
                        text_auto=True,
                        color_discrete_sequence=px.colors.qualitative.Set2
                    ),
                    use_container_width=True
                )

                with st.expander("🧩 Task Completion per Task Type", expanded=False):
                    st.dataframe(
                        df_filtered.groupby("Task")["Completion"].sum().reset_index(),
                        use_container_width=True
                    )
                    st.plotly_chart(
                        px.bar(
                            df_filtered.groupby("Task")["Completion"].sum().reset_index(),
                            x="Task",
                            y="Completion",
                            title=f"Task Type Breakdown ({month_year})",
                            text_auto=True,
                            color="Completion",
                            color_continuous_scale="Viridis"
                        ),
                        use_container_width=True
                    )

                monthly_data[month_year] = df_filtered

        # --- Yearly Comparison ---
        with month_tabs[-1]:
            st.markdown("## 🏆 Yearly & Monthly Comparison Overview")

            combined_filtered = pd.concat(monthly_data.values(), ignore_index=True)
            total_completion_year = int(combined_filtered["Completion"].sum())
            total_active_persons = combined_filtered["Person"].nunique()
            top_performer_year = (
                combined_filtered.groupby("Person")["Completion"].sum().idxmax()
                if not combined_filtered.empty else "N/A"
            )

            hide_top = st.checkbox("🙈 Hide Top Performer (Year)", value=False)
            c1, c2, c3 = st.columns(3)
            c1.metric("📅 Total Yearly Completions", f"{total_completion_year}")
            c2.metric("👥 Total Active Persons", f"{total_active_persons}")
            if not hide_top:
                c3.metric("🏆 Top Performer (Year)", top_performer_year)
            else:
                c3.empty()

            st.markdown("---")
            monthly_summary = combined_filtered.groupby(["Month_Year", "Person"])["Completion"].sum().reset_index()
            monthly_summary["Month_Order"] = monthly_summary["Month_Year"].apply(parse_month_year)
            monthly_summary = monthly_summary.sort_values("Month_Order")

            st.markdown("### 📈 Monthly Completion Trend per Person")
            trend_fig = px.line(
                monthly_summary,
                x="Month_Year",
                y="Completion",
                color="Person",
                markers=True,
                title="Completion Trend per Person (Chronological)",
                color_discrete_sequence=px.colors.qualitative.Set2
            )
            trend_fig.update_xaxes(categoryorder="array", categoryarray=monthly_summary["Month_Year"].unique())
            st.plotly_chart(trend_fig, use_container_width=True)

            portal_summary = combined_filtered.groupby(["Month_Year", "Portal"])["Completion"].sum().reset_index()
            portal_summary["Month_Order"] = portal_summary["Month_Year"].apply(parse_month_year)
            portal_summary = portal_summary.sort_values("Month_Order")

            st.markdown("### 🧱 Portal Completion per Month")
            portal_fig = px.bar(
                portal_summary,
                x="Month_Year",
                y="Completion",
                color="Portal",
                barmode="stack",
                title="Portal Completion per Month (Chronological)",
                color_discrete_sequence=px.colors.qualitative.Set3
            )
            portal_fig.update_xaxes(categoryorder="array", categoryarray=portal_summary["Month_Year"].unique())
            st.plotly_chart(portal_fig, use_container_width=True)

            st.markdown("### 🏅 Leaderboard - Top Performers of the Year")
            leaderboard = (
                combined_filtered.groupby("Person")["Completion"]
                .sum()
                .reset_index()
                .sort_values(by="Completion", ascending=False)
            )
            st.dataframe(leaderboard, use_container_width=True)
            st.plotly_chart(
                px.bar(
                    leaderboard,
                    x="Person",
                    y="Completion",
                    title="Top Performers (All Months Combined)",
                    text_auto=True,
                    color="Completion",
                    color_continuous_scale="Viridis"
                ),
                use_container_width=True
            )

    else:
        st.warning("⚠️ No valid completion data found in uploaded files.")
else:
    st.info("📤 Please upload one or more Excel reports to begin analysis.")
