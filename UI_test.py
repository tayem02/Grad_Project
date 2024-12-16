import os
#import fitz  # PyMuPDF for reading PDFs
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
excel_file_path = "C:\\Graduation Project\\OpenAI_APIs\\Results_ModelName_16_Dec_2024.xlsx"
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

# Custom CSS for styling tables and removing extra whitespace
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
# Centered container for all content
st.markdown('<div class="centered">', unsafe_allow_html=True)

st.title("📊 Project Management Dashboard")

# Sidebar for selecting a project
st.sidebar.header("Select a Project")

if "Projects" in dataframes and "Stakeholders_Projects" in dataframes and "Stakeholders_Details" in dataframes:
    project_df = dataframes["Projects"]
    stakeholders_projects_df = dataframes["Stakeholders_Projects"]
    stakeholders_details_df = dataframes["Stakeholders_Details"]

    project_names = project_df["Project Name"].unique()
    selected_project = st.sidebar.selectbox("Choose a project:", project_names)

    # Filter DataFrames based on selected project
    selected_project_data = project_df[project_df["Project Name"] == selected_project]

    # Display Project Details (Transposed)
    st.write("### Project Details")
    with st.container():
        st.dataframe(selected_project_data.transpose())

    # Filter and display Stakeholders Details linked to the project
    project_ids = selected_project_data["Project ID"].values
    linked_stakeholders = stakeholders_projects_df[stakeholders_projects_df["Project ID"].isin(project_ids)]
    stakeholder_ids = linked_stakeholders["Stakeholder ID"].values
    filtered_stakeholders_details = stakeholders_details_df[stakeholders_details_df["ID"].isin(stakeholder_ids)]

    st.write("### Stakeholders Details")
    with st.container():
        st.dataframe(filtered_stakeholders_details)

    # Display data from other sheets
    for sheet, df in dataframes.items():
        if sheet not in ["Projects", "Stakeholders_Projects", "Stakeholders_Details"]:
            st.write(f"### {sheet.replace('_', ' ')}")
            filtered_df = df[df["Project ID"].isin(selected_project_data["Project ID"])]
            with st.container():
                st.dataframe(filtered_df)
else:
    st.error("No Projects or Stakeholders data available.")

# Close the centered container
st.markdown('</div>', unsafe_allow_html=True)
