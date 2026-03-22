import streamlit as st

from session_utils import require_workbook
from modules.m01_workbook_loader import load_current_state_steps
from modules.m02_nlp_normalization import normalize_steps
from modules.m03_ie_ontology import map_steps_to_ie_ontology
from modules.m04_rule_engine import apply_rule_engine

st.set_page_config(page_title="Lean Rule Engine", layout="wide")
st.title("9 Lean Rule Engine")

ctx = require_workbook()
steps = load_current_state_steps(ctx["excel_file"], sheet_name=ctx["sheet_name"])
normalized = normalize_steps(steps)
ontology = map_steps_to_ie_ontology(normalized)
ruled = apply_rule_engine(ontology)

st.dataframe(ruled, use_container_width=True)
