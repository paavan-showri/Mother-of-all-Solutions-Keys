import streamlit as st

from session_utils import require_workbook
from modules.m01_workbook_loader import load_current_state_steps
from modules.m02_nlp_normalization import normalize_steps, normalized_to_df

st.set_page_config(page_title='NLP Normalization', layout='wide')
st.title('7 NLP Normalization')
ctx = require_workbook()
use_zero_shot = st.checkbox('Use zero-shot waste classification', value=False)
use_manufactubert = st.checkbox('Use ManufactuBERT semantic hints', value=False)
steps = load_current_state_steps(ctx['excel_file'], sheet_name=ctx['sheet_name'])
normalized = normalize_steps(steps, use_zero_shot=use_zero_shot, use_manufactubert=use_manufactubert)
st.dataframe(normalized_to_df(normalized), use_container_width=True)
