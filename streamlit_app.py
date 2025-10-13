import os
import sys
import subprocess

# --- Auto-install required packages ---
def ensure(package):
    """Ensure a Python package is installed."""
    try:
        __import__(package)
    except ImportError:
        print(f"[INFO] Installing {package} ...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])

# Ensure dependencies
for pkg in ["streamlit", "pandas", "plotly", "openpyxl"]:
    ensure(pkg)

# --- Imports (safe after installation) ---
import streamlit as st
import pandas as pd
import plotly.express as px

# --- Page Config ---
st.set_page_config(page_title="📊 Task Reports Dashboard", layout="wide")

# --- Header ---
st.title("📊 Task Reports Dashboard")
st.markdown("Easily visualize and track your monthly task performance.")

# --- Sidebar ---
with st.sidebar:
    st.header("⚙️ Dashboard Controls")

    # Collapsible Section: Upload Excel
    with st.expander("📂 Upload Monthly Excel Reports", expanded=True):
        uploaded_files = st.file_uploader(
            "Upload Excel files (one or multiple months)",
            type=["xlsx"],
            accept_multiple_files=True
        )

    # Collapsible Section: Select Persons
    with st.expander("👥 Select Persons to Display", expanded=True):
        person_selection_placeholder = st.empty()

    # Collapsible Section: Hide Top Performer
    with st.expander("🏆 Top Performer Settings", expanded=False):
        hide_top_performer = st.checkbox("Hide Top Performer Section", value=False)

# --- Main Logic ---
if uploaded_files:
    dfs = []
    for file in uploaded_files:
        try:
            df = pd.read_excel(file, engine="openpyxl")
            dfs.append(df)
        except Exception as e:
            st.error(f"❌ Error reading {file.name}: {e}")

    if dfs:
        combined_df = pd.concat(dfs, ignore_index=True)
        st.success(f"✅ Successfully loaded {len(dfs)} file(s).")

        # Data Preview
        st.subheader("📋 Combined Task Data")
        st.dataframe(combined_df, use_container_width=True)

        # --- Filter by Person ---
        if "Person" in combined_df.columns:
            persons = combined_df["Person"].dropna().unique().tolist()
            with person_selection_placeholder.container():
                selected_persons = st.multiselect(
                    "Select Person(s) to display:",
                    options=persons,
                    default=persons
                )
            combined_filtered = combined_df[combined_df["Person"].isin(selected_persons)]
        else:
            st.warning("⚠️ 'Person' column not found in uploaded file(s).")
            combined_filtered = combined_df

        # --- Visualization ---
        if "Completion" in combined_filtered.columns:
            st.subheader("📈 Task Completion Summary")
            fig = px.bar(
                combined_filtered,
                x="Person",
                y="Completion",
                color="Person",
                title="Task Completion by Person",
                text="Completion",
                color_discrete_sequence=px.colors.qualitative.Bold
            )
            fig.update_traces(textposition="outside")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.warning("⚠️ 'Completion' column not found in uploaded file(s).")

        # --- Top Performer Section ---
        if not hide_top_performer:
            if "Person" in combined_filtered.columns and "Completion" in combined_filtered.columns:
                try:
                    top_performer = (
                        combined_filtered.groupby("Person")["Completion"].sum().idxmax()
                    )
                    total_completion = (
                        combined_filtered.groupby("Person")["Completion"].sum().max()
                    )
                    st.markdown("---")
                    st.subheader("🏅 Top Performer")
                    st.success(f"**{top_performer}** with total completion of **{total_completion} tasks**!")
                except Exception:
                    st.info("No valid data to compute top performer.")
        else:
            st.info("🏆 Top Performer section is hidden.")
    else:
        st.warning("⚠️ No valid Excel data found.")
else:
    st.info("📥 Please upload one or more Excel files to start.")

# --- Footer ---
st.markdown("---")
st.caption("Developed with ❤️ using Streamlit + Plotly + Pandas")
