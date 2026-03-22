import streamlit as st
import pandas as pd
from ie_modules_v2.m01_workbook_loader import load_current_state_steps
from ie_modules_v2.m02_nlp_normalization import normalize_steps

st.title("7 NLP Normalization")
if "uploaded_excel" not in st.session_state:
    st.warning("Upload the workbook in Home first.")
    st.stop()

steps = load_current_state_steps(st.session_state["uploaded_excel"])
normalized = normalize_steps(steps)
df = pd.DataFrame([vars(x) for x in normalized])
st.dataframe(df, use_container_width=True)
