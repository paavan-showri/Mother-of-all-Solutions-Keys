import streamlit as st

st.set_page_config(page_title="Industrial Engineering Optimizer", layout="wide")
st.title("Industrial Engineering Optimizer")
st.write("Upload one Excel workbook in the required format, then open any page in the sidebar.")

uploaded = st.file_uploader("Upload Excel workbook", type=["xlsx"])
if uploaded is not None:
    st.session_state["uploaded_excel"] = uploaded
    st.success("Workbook uploaded and saved in session state.")
else:
    st.info("Upload an Excel workbook to continue.")

st.markdown("""
Expected sheets:
- `FPC_Current State`
- `Precedence Network`
- `Resources`

Expected current-state headers:
- Step
- Description
- Activity
- Start time
- End time
- Duration (Sec)
- Activity Type
- Resources
""")
