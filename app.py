import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

# Load the Excel file
@st.cache
def load_data(file):
    # Read the Excel file with the specific sheet names
    xls = pd.ExcelFile(file)
    contracts_df = pd.read_excel(xls, 'Contracts')
    consultant_billing_df = pd.read_excel(xls, 'Consultant Billing')
    return contracts_df, consultant_billing_df

# Process the Contracts sheet
def process_contracts_data(contracts_df):
    # Select relevant columns based on your input
    contracts_df = contracts_df[['Client', 'Work', 'PO No.', 'BH', 'Total Value (F+V)', 'Fixed Balance']]

    # Remove any rows with missing data in these columns
    contracts_df = contracts_df.dropna()

    # Calculate PO Utilization based on the formula
    contracts_df['Utilization (%)'] = ((contracts_df['Total Value (F+V)'] - contracts_df['Fixed Balance']) / contracts_df['Total Value (F+V)']) * 100

    # Convert to Indian currency format
    contracts_df['Total Value (F+V)'] = contracts_df['Total Value (F+V)'].apply(lambda x: f"₹{x:,.2f}")
    contracts_df['Fixed Balance'] = contracts_df['Fixed Balance'].apply(lambda x: f"₹{x:,.2f}")
    
    return contracts_df

# Process the Consultant Billing sheet
def process_consultant_billing_data(consultant_billing_df):
    # Data Processing for Consultant Billing: we will pivot the data based on hierarchy and months
    consultant_billing_df = consultant_billing_df[['Business Head', 'Consultant', 'Client', 'T Amt', 'Ded', 'N Amt', 'Days']]
    consultant_billing_df = consultant_billing_df.dropna()

    # Aggregating the data based on Business Head, Consultant and Client
    summary_df = consultant_billing_df.groupby(['Business Head', 'Consultant', 'Client']).agg({
        'T Amt': 'sum',
        'Ded': 'sum',
        'N Amt': 'sum',
        'Days': 'sum'
    }).reset_index()

    # Convert to Indian currency format
    summary_df['T Amt'] = summary_df['T Amt'].apply(lambda x: f"₹{x:,.2f}")
    summary_df['Ded'] = summary_df['Ded'].apply(lambda x: f"₹{x:,.2f}")
    summary_df['N Amt'] = summary_df['N Amt'].apply(lambda x: f"₹{x:,.2f}")

    return summary_df

# Display Data and Visualizations in Streamlit
def display_data(contracts_df, consultant_billing_df):
    st.title("Contract & Consultant Billing Dashboard (FY 2024-25)")

    # Display Filters
    client_filter = st.selectbox('Select Client', contracts_df['Client'].unique())
    bh_filter = st.selectbox('Select Business Head', contracts_df['BH'].unique())

    # Filter data based on the selections
    filtered_contracts_df = contracts_df[(contracts_df['Client'] == client_filter) & (contracts_df['BH'] == bh_filter)]
    filtered_consultant_billing_df = consultant_billing_df[(consultant_billing_df['Client'] == client_filter) & (consultant_billing_df['Business Head'] == bh_filter)]

    # Display the data
    st.subheader('Contract Summary')
    st.dataframe(filtered_contracts_df)

    st.subheader('Consultant Billing Summary')
    st.dataframe(filtered_consultant_billing_df)

    # Graphical Representation
    st.subheader('PO Utilization by Business Head')
    utilization_data = filtered_contracts_df.groupby('BH')['Utilization (%)'].mean()
    st.bar_chart(utilization_data)

    st.subheader('Monthly Trend of Net Amount for Consultant')
    monthly_trend_data = filtered_consultant_billing_df.groupby('Consultant')['N Amt'].sum()
    st.bar_chart(monthly_trend_data)

# Streamlit app execution
def main():
    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])
    
    if uploaded_file is not None:
        # Load Data
        contracts_df, consultant_billing_df = load_data(uploaded_file)

        # Process Data
        contracts_df = process_contracts_data(contracts_df)
        consultant_billing_df = process_consultant_billing_data(consultant_billing_df)

        # Display data and visuals
        display_data(contracts_df, consultant_billing_df)

if __name__ == "__main__":
    main()
