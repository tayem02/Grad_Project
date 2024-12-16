
import os
import fitz  # PyMuPDF for reading PDFs
import pandas as pd
import openai
from datetime import datetime
import streamlit as st  # type: ignore
from faker import Faker

# Set up Faker
fake = Faker()

# Set page configuration for full-screen mode
st.set_page_config(
    page_title="Project Management Dashboard",
    page_icon="📊",
    layout="wide",  # Full-screen layout
    initial_sidebar_state="expanded"  # Sidebar starts expanded
)

# Paths and sheet names
excel_file_path = "https://raw.githubusercontent.com/tayem02/Grad_Project/main/Results_(3.5_gpt)_16_Dec_2024.xlsx"
sheet_names = {
    "Projects": "Projects",
    "Goals": "Goals",
    "Tasks": "Tasks",
    "Stakeholders_Projects": "Stakeholders_Projects",
    "Stakeholders_Details": "Stakeholders_Details"
}

# Read data from the Excel file
dataframes = {}
for sheet, sheet_name in sheet_names.items():
    try:
        dataframes[sheet] = pd.read_excel(excel_file_path, sheet_name=sheet_name)
    except Exception as e:
        st.error(f"Error reading sheet {sheet_name}: {e}")

# Custom CSS for styling tables, text wrapping, and layout adjustments
st.markdown(
    """
<style>
    .centered {
        display: flex;
        justify-content: center;
        align-items: center;
        flex-direction: column;
    }
    .dataframe-container {
        overflow-x: auto;
        width: 100%;
    }
    .dataframe th, .dataframe td {
        white-space: pre-wrap !important; /* Ensure text wraps properly */
        word-wrap: break-word;
        text-align: left;
    }
    .block-container {
        padding-top: 0rem !important; /* Remove white space at the top */
    }
</style>
""",
    unsafe_allow_html=True
)

# Streamlit UI
# Move project filter to the top
st.title("📊 Project Management Dashboard")

if "Projects" in dataframes:
    project_df = dataframes["Projects"]
    project_names = project_df["Project Name"].unique()

    # Filter at the top
    st.write("### Select a Project")
    selected_project = st.selectbox("Choose a project:", project_names, key="top_project_filter")

    # Filter DataFrames based on selected project
    selected_project_data = project_df[project_df["Project Name"] == selected_project]

    # Layout: First row - Project Details and Tasks in two columns
    if "Tasks" in dataframes:
        tasks_df = dataframes["Tasks"]
        filtered_tasks = tasks_df[tasks_df["Project ID"].isin(selected_project_data["Project ID"])]

        col1, col2 = st.columns(2)

        with col1:
            st.write("### Project Details")
            st.dataframe(selected_project_data.transpose(), height=400)

        with col2:
            st.write("### Tasks Details")
            st.dataframe(filtered_tasks, height=400)

    # Layout: Second row - Goals, Stakeholders Details, and Stakeholders Projects in three columns
    if "Stakeholders_Projects" in dataframes and "Stakeholders_Details" in dataframes and "Goals" in dataframes:
        stakeholders_projects_df = dataframes["Stakeholders_Projects"]
        stakeholders_details_df = dataframes["Stakeholders_Details"]
        goals_df = dataframes["Goals"]

        project_ids = selected_project_data["Project ID"].values

        # Filter Goals
        filtered_goals = goals_df[goals_df["Project ID"].isin(project_ids)]

        # Filter Stakeholders Projects
        linked_stakeholders = stakeholders_projects_df[stakeholders_projects_df["Project ID"].isin(project_ids)]
        stakeholder_ids = linked_stakeholders["Stakeholder ID"].values

        # Filter Stakeholders Details
        filtered_stakeholders_details = stakeholders_details_df[stakeholders_details_df["ID"].isin(stakeholder_ids)]

        col1, col2, col3 = st.columns(3)

        with col1:
            st.write("### Goals")
            st.dataframe(filtered_goals, height=400)

        with col2:
            st.write("### Stakeholders Details")
            st.dataframe(filtered_stakeholders_details, height=400)

        with col3:
            st.write("### Stakeholders Projects")
            st.dataframe(linked_stakeholders, height=400)
else:
    st.error("No Projects data available.")
