import io

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Case Study", layout="wide")
st.title("Case Study")
st.write("Upload the Excel workbook once, then open any page from the left sidebar.")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls", "xlsm"])
sheet_name = st.text_input("Current-state sheet name", value="FPC_Current State")
precedence_sheet = st.text_input("Precedence sheet name", value="Precedence Network")
resource_sheet = st.text_input("Resources sheet name", value="Resources")

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
    except Exception as exc:
        st.error(f"Workbook could not be read: {exc}")
else:
    st.info("Please upload your Excel file.")

if "excel_file_name" in st.session_state:
    st.write(f'Current file: **{st.session_state["excel_file_name"]}**')

st.markdown(
    """
### Pipeline
- Waste Analysis
- NLP Normalization
- IE Ontology
- Lean Rule Engine
- Macro Tasks
- Precedence Network
- RCPSP Schedule
- Analytics

### Future-state discovery
Current-state workbook → NLP normalization → IE ontology → Lean rules → automatic macro-task discovery → inferred precedence → RCPSP

### Notes
- No ManufactuBERT
- No hard-coded planner template mode
- No fixed duration dictionary
- No fixed predecessor dictionary
"""
)
