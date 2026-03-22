import streamlit as st

from session_utils import require_workbook
from modules.m01_workbook_loader import load_current_state_steps
from modules.m02_nlp_normalization import normalize_steps
from modules.m03_ie_ontology import map_steps_to_ie_ontology

st.set_page_config(page_title="IE Ontology", layout="wide")
st.title("8 IE Ontology")

ctx = require_workbook()
steps = load_current_state_steps(ctx["excel_file"], sheet_name=ctx["sheet_name"])
normalized = normalize_steps(steps)
ontology = map_steps_to_ie_ontology(normalized)

st.dataframe(ontology, use_container_width=True)
