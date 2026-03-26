import streamlit as st

from session_utils import require_workbook
from modules.m01_workbook_loader import load_current_state_steps
from modules.m02_nlp_normalization import normalize_steps
from modules.m03_ie_ontology import map_steps_to_ie_ontology
from modules.m04_rule_engine import apply_rule_engine

st.set_page_config(page_title='Lean Rule Engine', layout='wide')
st.title('9 Lean Rule Engine')
ctx = require_workbook()
use_zero_shot = st.checkbox('Use zero-shot waste classification', value=False)
use_manufactubert = st.checkbox('Use ManufactuBERT semantic hints', value=False)
steps = load_current_state_steps(ctx['excel_file'], sheet_name=ctx['sheet_name'])
normalized = normalize_steps(steps, use_zero_shot=use_zero_shot, use_manufactubert=use_manufactubert)
ontology = map_steps_to_ie_ontology(normalized)
ruled = apply_rule_engine(ontology)
st.dataframe(ruled, use_container_width=True)
