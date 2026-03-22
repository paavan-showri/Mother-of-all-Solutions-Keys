import streamlit as st
from ie_modules_v2.m01_workbook_loader import load_current_state_steps
from ie_modules_v2.m02_nlp_normalization import normalize_steps
from ie_modules_v2.m03_ie_ontology import map_steps_to_ie_ontology

st.title("8 IE Ontology")
if "uploaded_excel" not in st.session_state:
    st.warning("Upload the workbook in Home first.")
    st.stop()

steps = load_current_state_steps(st.session_state["uploaded_excel"])
normalized = normalize_steps(steps)
ontology = map_steps_to_ie_ontology(normalized)
st.dataframe(ontology, use_container_width=True)
