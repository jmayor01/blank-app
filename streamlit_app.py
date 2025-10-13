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
st.title("📊 Project Completion Analyzer - Multi-Month Portal Reports")

# --- File Upload ---
uploaded_files = st.file_uploader("Upload monthly Excel reports", type=["xlsx"], accept_multiple_files=True)

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
                    if (
                        value
                        and value not in known_tasks
                        and value.lower() != "row labels"
                        and "portal" not in value.lower()
                        and not value.lower().startswith("total")
                    ):
                        all_persons_detected.add(value)
        except Exception as e:
            st.warning(f"Error scanning {uploaded_file.name}: {e}")

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

                    for idx, row in portal_df.iterrows():
                        cell_value = str(row.iloc[0]).strip()
                        completion_value = row.iloc[1]

                        if not cell_value or cell_value.lower() == "total":
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
                st.warning(f"No valid data found in {file_name}. Please verify format.")

        except Exception as e:
            st.error(f"Error reading {file_name}: {e}")

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

        # Tabs per month
        month_tabs = st.tabs(list(month_list) + ["📈 Yearly Comparison"])
        monthly_data = {}

        for i, month_year in enumerate(month_list):
            with month_tabs[i]:
                df_month = combined_df[combined_df["Month_Year"] == month_year]
                st.markdown(f"## 📅 {month_year} Summary")

                total_completion = int(df_month["Completion"].sum())
                active_persons = df_month["Person"].nunique()
                top_performer = (
                    df_month.groupby("Person")["Completion"].sum().idxmax()
                    if not df_month.empty else "N/A"
                )

                hide_top = st.checkbox(f"🙈 Hide Top Performer ({month_year})", value=False)

                c1, c2, c3 = st.columns(3)
                c1.metric("✅ Total Completions", f"{total_completion}")
                c2.metric("👥 Active Persons", f"{active_persons}")
                if not hide_top:
                    c3.metric("🏆 Top Performer", top_performer)
                else:
                    c3.empty()

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

                # --- Clean Summary Table (No background color) ---
                st.markdown("### 👤 Task Completion Summary per Person and Portal")
                st.dataframe(
                    df_filtered.groupby(["Person", "Portal"])["Completion"]
                    .sum()
                    .reset_index()
                    .style.set_table_styles(
                        [{'selector': 'th', 'props': [('font-weight', 'bold'), ('text-align', 'center')]},
                         {'selector': 'td', 'props': [('text-align', 'center')]}]
                    ),
                    use_container_width=True
                )

                # --- Bar Chart ---
                st.markdown("#### 📊 Bar Chart - Completion per Person")
                bar_fig = px.bar(
                    df_filtered.groupby(["Person", "Portal"])["Completion"].sum().reset_index(),
                    x="Person",
                    y="Completion",
                    color="Portal",
                    barmode="group",
                    title=f"Task Completions per Person ({month_year})",
                    text_auto=True,
                    color_discrete_sequence=px.colors.qualitative.Set2
                )
                st.plotly_chart(bar_fig, use_container_width=True)

                # --- Task Summary Table ---
                st.markdown("### 🧩 Task Completion per Task Type")
                st.dataframe(
                    df_filtered.groupby("Task")["Completion"]
                    .sum()
                    .reset_index()
                    .style.set_table_styles(
                        [{'selector': 'th', 'props': [('font-weight', 'bold'), ('text-align', 'center')]},
                         {'selector': 'td', 'props': [('text-align', 'center')]}]
                    ),
                    use_container_width=True
                )

                # --- Task Chart ---
                task_fig = px.bar(
                    df_filtered.groupby("Task")["Completion"].sum().reset_index(),
                    x="Task",
                    y="Completion",
                    title=f"Task Type Breakdown ({month_year})",
                    text_auto=True,
                    color="Completion",
                    color_continuous_scale="Viridis"
                )
                st.plotly_chart(task_fig, use_container_width=True)

                monthly_data[month_year] = df_filtered

        # --- Yearly Comparison Tab ---
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
            st.dataframe(
                leaderboard.style.set_table_styles(
                    [{'selector': 'th', 'props': [('font-weight', 'bold'), ('text-align', 'center')]},
                     {'selector': 'td', 'props': [('text-align', 'center')]}]
                ),
                use_container_width=True
            )

            leader_fig = px.bar(
                leaderboard,
                x="Person",
                y="Completion",
                title="Top Performers (All Months Combined)",
                text_auto=True,
                color="Completion",
                color_continuous_scale="Viridis"
            )
            st.plotly_chart(leader_fig, use_container_width=True)

    else:
        st.warning("No valid completion data found in uploaded files.")
else:
    st.info("📤 Please upload one or more Excel reports to begin analysis.")
