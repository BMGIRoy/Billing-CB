import streamlit as st
import pandas as pd
import altair as alt

# Function to process the contract data
def process_contract_data(df):
    # Selecting relevant columns
    contracts_df = df[["Client Name", "Type of Work", "PO No.", "Business Head", "Total PO Value", "PO Balance"]]
    
    # Calculate Utilization
    contracts_df['Utilization'] = ((contracts_df['Total PO Value'] - contracts_df['PO Balance']) / contracts_df['Total PO Value']) * 100
    
    # Convert numbers to Indian currency format
    contracts_df['Total PO Value'] = contracts_df['Total PO Value'].apply(lambda x: f'₹ {x:,.0f}')
    contracts_df['PO Balance'] = contracts_df['PO Balance'].apply(lambda x: f'₹ {x:,.0f}')
    
    return contracts_df

# Function to process the consultant billing data
def process_consultant_billing_data(df):
    # Selecting relevant columns
    consultant_df = df[["Business Head", "Consultant", "Client Name", "T Amt", "Ded", "N Amt", "Days"]]
    
    # Convert numbers to Indian currency format
    consultant_df['T Amt'] = consultant_df['T Amt'].apply(lambda x: f'₹ {x:,.0f}')
    consultant_df['Ded'] = consultant_df['Ded'].apply(lambda x: f'₹ {x:,.0f}')
    consultant_df['N Amt'] = consultant_df['N Amt'].apply(lambda x: f'₹ {x:,.0f}')
    
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
            
            # Contract Utilization Visualization (Bar Chart)
            contract_utilization_chart = alt.Chart(contracts_df).mark_bar().encode(
                x='Client Name:N',
                y='Utilization:Q',
                color='Business Head:N',
                tooltip=['Client Name', 'Utilization']
            )
            st.altair_chart(contract_utilization_chart, use_container_width=True)

            # Consultant Billing by Client Visualization (Bar Chart)
            consultant_billing_chart = alt.Chart(consultant_df).mark_bar().encode(
                x='Client Name:N',
                y='N Amt:Q',
                color='Consultant:N',
                tooltip=['Client Name', 'N Amt']
            )
            st.altair_chart(consultant_billing_chart, use_container_width=True)

        else:
            st.error("The uploaded file does not contain the required sheets ('Contracts' and 'Consultant Billing'). Please check the file.")
    except Exception as e:
        st.error(f"An error occurred: {e}")
else:
    st.warning("Please upload a file to proceed.")
