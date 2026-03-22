import streamlit as st
import pandas as pd
from ie_modules_v2.m01_workbook_loader import load_precedence_tasks, load_current_state_steps
from ie_modules_v2.m02_nlp_normalization import normalize_steps
from ie_modules_v2.m05_macro_tasks import generate_macro_tasks

st.title("10 Macro Tasks")
if "uploaded_excel" not in st.session_state:
    st.warning("Upload the workbook in Home first.")
    st.stop()

file = st.session_state["uploaded_excel"]
manual_tasks = load_precedence_tasks(file)
st.subheader("Macro tasks from Precedence Network sheet")
st.dataframe(pd.DataFrame([vars(t) for t in manual_tasks]), use_container_width=True)

steps = load_current_state_steps(file)
normalized = normalize_steps(steps)
generated = generate_macro_tasks(normalized)
st.subheader("Macro tasks generated from FPC rows")
st.dataframe(pd.DataFrame([vars(t) for t in generated]), use_container_width=True)
