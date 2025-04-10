import streamlit as st
import pandas as pd

# Function to process contract data
def process_contract_data(contracts_df):
    # Filter the necessary columns
    contracts_df = contracts_df[['Client Name', 'Type of Work', 'P O No.', 'Business Head', 'Total PO Value', 'PO Balance']]
    
    # Calculate Utilization
    contracts_df['Utilization'] = ((contracts_df['Total PO Value'] - contracts_df['PO Balance']) / contracts_df['Total PO Value']) * 100
    
    # Convert numbers to Indian currency format
    contracts_df['Total PO Value'] = contracts_df['Total PO Value'].apply(lambda x: f'₹ {x:,.0f}')
    contracts_df['PO Balance'] = contracts_df['PO Balance'].apply(lambda x: f'₹ {x:,.0f}')
    
    # Return the cleaned data
    return contracts_df

# Function to process consultant billing data
def process_consultant_billing_data(consultant_df):
    # Filter the necessary columns
    consultant_df = consultant_df[['Business Head', 'Consultant', 'Client Name', 'T Amt', 'Ded', 'N Amt', 'Days']]
    
    # Convert numbers to Indian currency format
    consultant_df['T Amt'] = consultant_df['T Amt'].apply(lambda x: f'₹ {x:,.0f}')
    consultant_df['Ded'] = consultant_df['Ded'].apply(lambda x: f'₹ {x:,.0f}')
    consultant_df['N Amt'] = consultant_df['N Amt'].apply(lambda x: f'₹ {x:,.0f}')
    
    # Return the cleaned data
    return consultant_df

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
