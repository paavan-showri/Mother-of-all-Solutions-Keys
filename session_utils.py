import io

import streamlit as st


def require_workbook():
    if "excel_file_bytes" not in st.session_state:
        st.warning("Please upload the Excel file first on the Home page.")
        st.stop()
    return {
        "excel_file": io.BytesIO(st.session_state["excel_file_bytes"]),
        "excel_file_name": st.session_state.get("excel_file_name", "uploaded.xlsx"),
        "sheet_name": st.session_state.get("sheet_name", "FPC_Current State"),
        "precedence_sheet": st.session_state.get("precedence_sheet", "Precedence Network"),
        "resource_sheet": st.session_state.get("resource_sheet", "Resources"),
    }
