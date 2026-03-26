import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Case Study Charts", layout="wide")
st.title("Case Study Chart Generator")
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
    except Exception as e:
        st.error(f"Workbook could not be read: {e}")
else:
    st.info("Please upload your Excel file.")

if "excel_file_name" in st.session_state:
    st.write(f'Current file: **{st.session_state["excel_file_name"]}**')

st.markdown(
    """
### Pipeline
1. Waste Analysis
2. NLP Normalization
3. IE Ontology
4. Lean Rule Engine
5. Macro Tasks
6. Precedence Network
7. RCPSP Schedule
8. Analytics

### Future-state modes
- **Automatic discovery mode**: current-state NLP → ontology/rules → macro tasks → RCPSP
- **Planner template mode**: fixed planner-style future-state durations for comparison

### NLP stack
- ManufactuBERT semantic hints (optional)
- Zero-shot waste classification (optional)
- Manufacturing ontology and Lean rules
- OR-Tools RCPSP scheduler
"""
)
