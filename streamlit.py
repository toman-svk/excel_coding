import streamlit as st
import pandas as pd
from excel_scripts import Excel_scripts

def main():
    st.title("Dashboard maker")

    uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])

    if uploaded_file:
        try:
            excel_path = uploaded_file.name
            ex = Excel_scripts(excel_path)
            st.write(f'Uploaded file with name {excel_path}')
            st.write(f'There are: {ex.get_sheet_count()} sheets.')

            ex.find_all_tables()

            # excel_data = pd.read_excel(uploaded_file, sheet_name=None)

            # # Display sheet names
            # sheet_names = list(excel_data.keys())
            # selected_sheet = st.selectbox("Select Sheet", sheet_names)

            # # Display data from the selected sheet
            # st.write(f"Showing data from sheet '{selected_sheet}':")
            # st.dataframe(excel_data[selected_sheet])
        
        except Exception as e:
            st.error(f"Error: {e}")
    
    else:
        st.info("Please upload an Excel file.")

if __name__ == "__main__":
    main()