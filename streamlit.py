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


            ex.find_all_tables()

            st.write(f'Uploaded file with name {excel_path}')
            st.write(f'There are: {ex.get_sheet_count()} sheet(s).')
            st.write(f'There are: {len(ex.table_dict.keys())} table(s).')

            for key, values in ex.table_dict.items():
                st.text_input("", f"{key}")
                st.write(values[0])

        
        except Exception as e:
            st.error(f"Error: {e}")
    
    else:
        st.info("Please upload an Excel file.")

if __name__ == "__main__":
    main()