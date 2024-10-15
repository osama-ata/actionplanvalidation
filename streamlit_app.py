import streamlit as st
from openpyxl import load_workbook

worksheet_name = 'Action_Plan'
Action_Plan_table_name = 'Action_Plan_Table'

st.title("Action Plan Validator and Compiler")
st.write(
    "This app was deveopped by [Osama Ata](https://osamata.com/)."
)


uploaded_files = st.file_uploader("Choose a file(s)",
                                 type=['xlsx'],
                                 accept_multiple_files=True)


if uploaded_files:
    for uploaded_file in uploaded_files:
        # Process each uploaded file
        wb = load_workbook(uploaded_file)
        # Access the first sheet by default (adjust sheet name if needed)
        ws = wb[worksheet_name]
        Action_Plan_table = ws.tables[Action_Plan_table_name]
        Action_Plan_table_range = Action_Plan_table.ref
        st.success(f"{uploaded_file.name} has Action_Plan_table in {Action_Plan_table_range}")