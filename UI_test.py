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

# Function to get a list of Excel files in the app's directory
def get_excel_files(directory):
    return [f for f in os.listdir(directory) if f.endswith('.xlsx')]

# Paths and sheet names
excel_file_directory = "./"  # Directory containing the Excel files
sheet_names = {
    "Projects": "Projects",
    "Goals": "Goals",
    "Tasks": "Tasks",
    "Stakeholders_Projects": "Stakeholders_Projects",
    "Stakeholders_Details": "Stakeholders_Details"
}

# Function to clean empty rows and rows with specific placeholders
def clean_empty_rows(df):
    return (df.dropna(how='all')
              .replace(r'^\s*$', pd.NA, regex=True)
              .dropna(how='all')
              .loc[~df.apply(lambda row: row.astype(str).str.contains('------------------------------------')).any(axis=1)])

# Load Generated Resources Data
resources_file = "Generated_Resources_Data.csv"
resources_df = None
if os.path.exists(resources_file):
    resources_df = pd.read_csv(resources_file)
    resources_df = clean_empty_rows(resources_df)

# Custom CSS for styling tables, cards, text wrapping, and layout adjustments
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
        word-wrap: break-word !important;
        text-align: left;
    }
    div.stDataFrame {
        width: 100% !important;
    }
    .card {
        background-color: #000000;
        color: #FFFFFF;
        border-radius: 10px;
        padding: 20px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        margin: 10px;
        border: 2px solid #ddd;
        text-align: center;
        height: auto;
        overflow: hidden;
    }
    .metric-value {
        font-size: 1.5em;
        font-weight: bold;
        margin-top: 10px;
    }
</style>
""",
    unsafe_allow_html=True
)

# Streamlit UI
st.title("📊 Project Management Dashboard")

# Model and Project selection in the same row
selection_col1, selection_col2 = st.columns(2)
with selection_col1:
    excel_files = get_excel_files(excel_file_directory)
    selected_excel_file = st.selectbox("Choose an Excel file:", excel_files, key="excel_file_selector")

if selected_excel_file:
    # Load data from the selected file
    dataframes = {}
    excel_file_path = os.path.join(excel_file_directory, selected_excel_file)
    
    for sheet, sheet_name in sheet_names.items():
        try:
            df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
            dataframes[sheet] = clean_empty_rows(df)  # Remove empty rows and rows with placeholder lines
        except Exception as e:
            st.error(f"Error reading sheet {sheet_name}: {e}")

    with selection_col2:
        project_df = dataframes.get("Projects", pd.DataFrame())
        project_names = project_df["Project Name"].unique() if not project_df.empty else []
        selected_project = st.selectbox("Choose a project:", project_names, key="project_filter")

    # Cards for KPIs after selection
    metrics_col1, metrics_col2 = st.columns(2)
    with metrics_col1:
        num_projects = len(dataframes.get("Projects", []))
        st.markdown(f"""
            <h4 class='card'>Number of Projects<br><br>
            <span class='metric-value'>{num_projects}</span></h4>
        """, unsafe_allow_html=True)
    with metrics_col2:
        num_resources = len(resources_df) if resources_df is not None else 0
        st.markdown(f"""
            <h4 class='card'>Number of Resources<br><br>
            <span class='metric-value'>{num_resources}</span></h4>
        """, unsafe_allow_html=True)

    if "Projects" in dataframes:
        selected_project_data = project_df[project_df["Project Name"] == selected_project]

        # Layout: First row - Project Details and Tasks in two columns
        if "Tasks" in dataframes:
            tasks_df = dataframes["Tasks"]
            filtered_tasks = tasks_df[tasks_df["Project ID"].isin(selected_project_data["Project ID"])]

            col1, col2 = st.columns(2)

            with col1:
                st.write("### Project Details")
                st.dataframe(selected_project_data.transpose(), use_container_width=True, height=400)

            with col2:
                st.write("### Tasks Details")
                st.dataframe(filtered_tasks, use_container_width=True, height=400)

        # Layout: Second row - Goals in a single expanded row
        if "Goals" in dataframes:
            goals_df = dataframes["Goals"]
            project_ids = selected_project_data["Project ID"].values
            filtered_goals = goals_df[goals_df["Project ID"].isin(project_ids)]
            st.write("### Goals")
            st.dataframe(filtered_goals, use_container_width=True, height=400)

        # Layout: Third row - Stakeholders Details and Stakeholders Projects in expanded columns
        if "Stakeholders_Projects" in dataframes and "Stakeholders_Details" in dataframes:
            stakeholders_projects_df = dataframes["Stakeholders_Projects"]
            stakeholders_details_df = dataframes["Stakeholders_Details"]
            linked_stakeholders = stakeholders_projects_df[stakeholders_projects_df["Project ID"].isin(project_ids)]
            stakeholder_ids = linked_stakeholders["Stakeholder ID"].values
            filtered_stakeholders_details = stakeholders_details_df[stakeholders_details_df["ID"].isin(stakeholder_ids)]

            col1, col2 = st.columns([1, 1])

            with col1:
                st.write("### Stakeholders Details")
                st.dataframe(filtered_stakeholders_details, use_container_width=True, height=400)

            with col2:
                st.write("### Stakeholders Projects")
                st.dataframe(linked_stakeholders, use_container_width=True, height=400)

        # New row: Generated Resources Data Table
        if resources_df is not None:
            st.write("### Resources Data")
            st.dataframe(resources_df, use_container_width=True, height=400)
    else:
        st.error("No Projects data available.")
