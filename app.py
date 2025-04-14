import streamlit as st
import pandas as pd

# Function to process the 'Contracts' sheet
def process_contract_data(df):
    # Select relevant columns and rename for clarity
    contracts_df = df[['Client', 'Work', 'PO No.', 'BH', 'Total Value (F+V)', 'Fixed Balance']]
    contracts_df.columns = ['Client Name', 'Type of Work', 'PO No.', 'Business Head', 'Total PO Value', 'PO Balance']
    
    # Calculate PO Utilization
    contracts_df['PO Utilization (%)'] = ((contracts_df['Total PO Value'] - contracts_df['PO Balance']) / contracts_df['Total PO Value']) * 100
    
    # Format the amounts as Indian Rupees (₹)
    contracts_df['Total PO Value'] = contracts_df['Total PO Value'].apply(lambda x: f'₹ {x:,.0f}')
    contracts_df['PO Balance'] = contracts_df['PO Balance'].apply(lambda x: f'₹ {x:,.0f}')
    
    return contracts_df

# Function to process the 'BillBook' sheet
def process_consultant_billing_data(df):
    # Select the necessary columns (CM to CX)
    billing_df = df.iloc[:, 89:106]
    
    # Rename columns for clarity
    billing_df.columns = ['Business Head', 'Consultant', 'Client', 'Apr_T_Amt', 'Apr_Ded', 'Apr_N_Amt', 'Apr_Days',
                          'May_T_Amt', 'May_Ded', 'May_N_Amt', 'May_Days']
    
    # Fill Business Head, Consultant, and Client based on hierarchy
    billing_df['Consultant'] = billing_df['Consultant'].fillna(method='ffill')
    billing_df['Client'] = billing_df['Client'].fillna(method='ffill')
    billing_df['Business Head'] = billing_df['Business Head'].fillna(method='ffill')
    
    # Convert the monthly values to numeric for aggregation
    billing_df[['Apr_T_Amt', 'Apr_Ded', 'Apr_N_Amt', 'Apr_Days', 'May_T_Amt', 'May_Ded', 'May_N_Amt', 'May_Days']] = \
        billing_df[['Apr_T_Amt', 'Apr_Ded', 'Apr_N_Amt', 'Apr_Days', 'May_T_Amt', 'May_Ded', 'May_N_Amt', 'May_Days']].apply(pd.to_numeric, errors='coerce')
    
    # Calculate totals for T Amt, Ded, N Amt, and Days
    billing_df['Total_T_Amt'] = billing_df[['Apr_T_Amt', 'May_T_Amt']].sum(axis=1)
    billing_df['Total_Ded'] = billing_df[['Apr_Ded', 'May_Ded']].sum(axis=1)
    billing_df['Total_N_Amt'] = billing_df[['Apr_N_Amt', 'May_N_Amt']].sum(axis=1)
    billing_df['Total_Days'] = billing_df[['Apr_Days', 'May_Days']].sum(axis=1)
    
    # Format the totals as Indian Rupees (₹)
    billing_df['Total_T_Amt'] = billing_df['Total_T_Amt'].apply(lambda x: f'₹ {x:,.0f}')
    billing_df['Total_Ded'] = billing_df['Total_Ded'].apply(lambda x: f'₹ {x:,.0f}')
    billing_df['Total_N_Amt'] = billing_df['Total_N_Amt'].apply(lambda x: f'₹ {x:,.0f}')
    
    return billing_df

# Streamlit app
def main():
    st.title("Contract and Consultant Billing Dashboard")
    
    # File uploader
    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])
    
    if uploaded_file is not None:
        try:
            # Read the Excel file
            excel_file = pd.ExcelFile(uploaded_file)
            
            # Display sheet names for debugging
            st.write("Sheet names in the uploaded file:", excel_file.sheet_names)
            
            # Check if required sheets exist
            if 'Contracts' in excel_file.sheet_names and 'BillBook' in excel_file.sheet_names:
                # Process the 'Contracts' sheet
                contracts_df = pd.read_excel(uploaded_file, sheet_name="Contracts")
                contracts_df = process_contract_data(contracts_df)

                # Process the 'BillBook' sheet
                billbook_df = pd.read_excel(uploaded_file, sheet_name="BillBook")
                billbook_df = process_consultant_billing_data(billbook_df)

                # Display Contract Data
                st.header("Contract Summary")
                st.write(contracts_df)

                # Display Consultant Billing Data
                st.header("Consultant Billing Summary")
                st.write(billbook_df)
                
                # Contract Utilization Visualization (Bar Chart)
                st.subheader('PO Utilization by Business Head')
                utilization_data = contracts_df.groupby('Business Head')['PO Utilization (%)'].mean()
                st.bar_chart(utilization_data)

                # Monthly Trend of Net Amount for Consultant
                st.subheader('Monthly Trend of Net Amount for Consultant')
                monthly_trend_data = billbook_df.groupby('Consultant')['Total_N_Amt'].sum()
                st.bar_chart(monthly_trend_data)

            else:
                st.error("The uploaded file does not contain the required sheets ('Contracts' and 'BillBook'). Please check the file.")
        except Exception as e:
            st.error(f"An error occurred: {e}")
    else:
        st.warning("Please upload a file to proceed.")

if __name__ == "__main__":
    main()
