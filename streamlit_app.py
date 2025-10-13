import streamlit as st
import pandas as pd
import plotly.express as px
import numpy as np

# --- Page Configuration ---
st.set_page_config(
    page_title="Monthly Task Report",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Custom Sidebar Styling ---
st.markdown("""
    <style>
    /* Sidebar styling */
    section[data-testid="stSidebar"] {
        background-color: #f8fafc;
        border-right: 1px solid #e5e7eb;
    }
    .sidebar-title {
        font-size: 20px !important;
        font-weight: 600 !important;
        color: #1f2937;
        margin-bottom: 10px;
        text-align: center;
    }
    .sidebar-icon {
        margin-right: 8px;
        font-size: 18px;
    }
    .stExpander {
        border-radius: 10px;
        border: 1px solid #e5e7eb;
        margin-bottom: 10px;
        background-color: white;
    }
    .stExpander:hover {
        border-color: #3b82f6;
        box-shadow: 0px 0px 8px rgba(59,130,246,0.2);
    }
    </style>
""", unsafe_allow_html=True)

# --- Sidebar Title ---
st.sidebar.markdown('<div class="sidebar-title">📊 Task Report Dashboard</div>', unsafe_allow_html=True)

# --- Sidebar Expander 1: Upload ---
with st.sidebar.expander("📂 Upload Monthly Excel Reports", expanded=True):
    uploaded_files = st.file_uploader(
        "Upload one or more Excel files (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=True
    )

# Placeholder for dataframe
all_data = pd.DataFrame()

if uploaded_files:
    for file in uploaded_files:
        df = pd.read_excel(file)
        df["Source File"] = file.name
        all_data = pd.concat([all_data, df], ignore_index=True)
    st.sidebar.success(f"✅ {len(uploaded_files)} file(s) uploaded successfully!")
else:
    st.sidebar.warning("⚠️ Please upload at least one Excel file to continue.")

# --- Sidebar Expander 2: Select Persons ---
with st.sidebar.expander("👥 Select Persons to Display", expanded=False):
    if not all_data.empty and "Person" in all_data.columns:
        persons = sorted(all_data["Person"].dropna().unique())
        selected_persons = st.multiselect(
            "Select one or more persons",
            persons,
            default=persons[:3] if len(persons) >= 3 else persons
        )
    else:
        selected_persons = []

# --- Main Content Area ---
st.title("📈 Monthly Task Summary")

if not all_data.empty:
    filtered_data = (
        all_data[all_data["Person"].isin(selected_persons)]
        if selected_persons
        else all_data
    )

    # --- Task Table ---
    st.subheader("📋 Task Summary Table")
    st.dataframe(filtered_data, use_container_width=True)

    # --- Bar Chart: Task Count per Person ---
    if "Person" in filtered_data.columns:
        st.subheader("📊 Task Count per Person")
        task_count = filtered_data["Person"].value_counts().reset_index()
        task_count.columns = ["Person", "Task Count"]

        fig = px.bar(
            task_count,
            x="Person",
            y="Task Count",
            color="Person",
            text="Task Count",
            title="Task Distribution by Person",
            template="plotly_white"
        )
        fig.update_layout(showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

    # --- Pie Chart: Task Status (if available) ---
    if "Status" in filtered_data.columns:
        st.subheader("📉 Task Status Overview")
        status_summary = (
            filtered_data.groupby("Status")["Status"]
            .count()
            .reset_index(name="Count")
        )

        fig2 = px.pie(
            status_summary,
            names="Status",
            values="Count",
            title="Overall Task Status Distribution",
            hole=0.3
        )
        st.plotly_chart(fig2, use_container_width=True)

else:
    st.info("👆 Upload your Excel files from the sidebar to view reports.")

# --- Footer ---
st.markdown("---")
st.markdown(
    "<center>📅 Developed for Monthly Task Insights | Powered by Streamlit + Plotly</center>",
    unsafe_allow_html=True
)
