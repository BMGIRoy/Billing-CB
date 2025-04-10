import streamlit as st
import pandas as pd

# Function to process the contract sheet
def process_contract_data(df):
    # Selecting relevant columns
    contracts_df = df[["Client Name", "Type of Work", "PO No.", "Business Head", "Total PO Value", "PO Balance"]]

    # Calculate Utilization
    contracts_df['Utilization'] = ((contracts_df['Total PO Value'] - contracts_df['PO Balance']) / contracts_df['Total PO Value']) * 100
    return contracts_df

# Function to process the consultant billing sheet
def process_consultant_billing_data(df):
    # Selecting relevant columns
    consultant_df = df[["Business Head", "Consultant", "Client Name", "Month", "T Amt", "Ded", "N Amt", "Days"]]
    return consultant_df

# Streamlit app UI setup
st.title("Contract & Consultant Billing Dashboard")

# File uploader
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

# Check if file is uploaded
if uploaded_file is not None:
    try:
        # Read the Excel file
        excel_file = pd.ExcelFile(uploaded_file)
        
        # Check if both required sheets exist in the file
        if 'contracts' in excel_file.sheet_names and 'consultant billing' in excel_file.sheet_names:
            # Process the 'contracts' sheet
            contracts_df = pd.read_excel(uploaded_file, sheet_name="contracts")
            contracts_df = process_contract_data(contracts_df)

            # Process the 'consultant billing' sheet
            consultant_df = pd.read_excel(uploaded_file, sheet_name="consultant billing")
            consultant_df = process_consultant_billing_data(consultant_df)

            # Display Contract Data
            st.header("Contract Summary")
            st.write(contracts_df)

            # Display Consultant Billing Data
            st.header("Consultant Billing Summary")
            st.write(consultant_df)
            
        else:
            st.error("The uploaded file does not contain the required sheets ('contracts' and 'consultant billing'). Please check the file.")
    except Exception as e:
        st.error(f"An error occurred: {e}")
else:
    st.warning("Please upload a file to proceed.")
