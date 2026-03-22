import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Case Study Charts", layout="wide")

st.title("Case Study Chart Generator")
st.write("Upload the Excel workbook once, then open any page from the left sidebar.")

uploaded_file = st.file_uploader(
    "Upload Excel file",
    type=["xlsx", "xls", "xlsm"]
)

sheet_name = st.text_input("Current-state sheet name", value="FPC_Current State")
precedence_sheet = st.text_input("Precedence sheet name", value="Precedence Network")
resource_sheet = st.text_input("Resource sheet name", value="Resources")

if uploaded_file is not None:
    file_bytes = uploaded_file.getvalue()
    st.session_state["excel_file_bytes"] = file_bytes
    st.session_state["excel_file_name"] = uploaded_file.name
    st.session_state["sheet_name"] = sheet_name
    st.session_state["precedence_sheet"] = precedence_sheet
    st.session_state["resource_sheet"] = resource_sheet

    try:
        xls = pd.ExcelFile(io.BytesIO(file_bytes))
        st.session_state["excel_sheet_names"] = xls.sheet_names
        st.success(f"Loaded file: {uploaded_file.name}")
        st.write("Detected sheets:", xls.sheet_names)
    except Exception as e:
        st.error(f"File uploaded, but workbook could not be read: {e}")
else:
    st.info("Please upload your Excel file.")

if "excel_file_name" in st.session_state:
    st.write(f"Current file: **{st.session_state['excel_file_name']}**")

if "sheet_name" in st.session_state:
    st.write(f"Current current-state sheet: **{st.session_state['sheet_name']}**")

if "precedence_sheet" in st.session_state:
    st.write(f"Current precedence sheet: **{st.session_state['precedence_sheet']}**")

if "resource_sheet" in st.session_state:
    st.write(f"Current resource sheet: **{st.session_state['resource_sheet']}**")

st.markdown(
    """
    ### Pages
    Use the left sidebar to open:
    - Resource Utilization
    - Impact vs Effort
    - Pareto Frequency
    - Pareto Total Time
    - Scatter Plot
    - Waste Analysis
    - NLP Normalization
    - IE Ontology
    - Lean Rule Engine
    - Macro Tasks
    - Precedence Network
    - RCPSP Schedule
    - Analytics

    ### Expected current-state headers
    - Step
    - Description
    - Activity
    - Start time
    - End time
    - Duration (Sec)
    - Activity Type
    - Resources
    """
)
