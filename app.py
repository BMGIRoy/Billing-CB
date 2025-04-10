import streamlit as st
import pandas as pd

# File uploader
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

# Check if file is uploaded
if uploaded_file is not None:
    try:
        # Read the Excel file
        excel_file = pd.ExcelFile(uploaded_file)
        
        # Display sheet names for debugging
        st.write("Sheet names in the uploaded file:", excel_file.sheet_names)
        
        # If both required sheets exist, process them
        if 'Contracts' in excel_file.sheet_names and 'Consultant Billing' in excel_file.sheet_names:
            # Process the 'Contracts' sheet
            contracts_df = pd.read_excel(uploaded_file, sheet_name="Contracts")
            contracts_df = process_contract_data(contracts_df)

            # Process the 'Consultant Billing' sheet
            consultant_df = pd.read_excel(uploaded_file, sheet_name="Consultant Billing")
            consultant_df = process_consultant_billing_data(consultant_df)

            # Display Contract Data
            st.header("Contract Summary")
            st.write(contracts_df)

            # Display Consultant Billing Data
            st.header("Consultant Billing Summary")
            st.write(consultant_df)
            
        else:
            st.error("The uploaded file does not contain the required sheets ('Contracts' and 'Consultant Billing'). Please check the file.")
    except Exception as e:
        st.error(f"An error occurred: {e}")
else:
    st.warning("Please upload a file to proceed.")
