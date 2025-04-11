import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import io
from data_processor import (
    process_excel_data, 
    filter_data, 
    get_fiscal_periods,
    get_fiscal_year_for_date
)
from visualization import (
    create_time_series_chart,
    create_hierarchy_chart,
    create_comparison_chart,
    create_quarterly_chart,
    create_annual_chart,
    create_consultant_performance_chart
)

# App configuration
st.set_page_config(
    page_title="Business Analytics Dashboard",
    page_icon="📊",
    layout="wide"
)

# App title
st.title("Business Analytics Dashboard")
st.markdown("Upload your business data to generate interactive visualizations")

# Sidebar for filters and upload
with st.sidebar:
    st.header("Upload Data")
    uploaded_file = st.file_uploader("Upload Excel file", type=['xlsx', 'xls'])
    
    # Initialize session state for filters
    if 'data_loaded' not in st.session_state:
        st.session_state.data_loaded = False
    
    if 'business_heads' not in st.session_state:
        st.session_state.business_heads = []
    
    if 'consultants' not in st.session_state:
        st.session_state.consultants = []
    
    if 'clients' not in st.session_state:
        st.session_state.clients = []
        
    if 'fiscal_periods' not in st.session_state:
        st.session_state.fiscal_periods = []
    
    # Filter options (only show when data is loaded)
    if st.session_state.data_loaded:
        st.header("Filters")
        
        selected_fiscal_period = st.selectbox(
            "Select Fiscal Period",
            options=st.session_state.fiscal_periods,
            index=0
        )
        
        selected_business_heads = st.multiselect(
            "Business Heads",
            options=st.session_state.business_heads,
            default=st.session_state.business_heads
        )
        
        selected_consultants = st.multiselect(
            "Consultants",
            options=st.session_state.consultants,
            default=st.session_state.consultants
        )
        
        selected_clients = st.multiselect(
            "Clients",
            options=st.session_state.clients,
            default=st.session_state.clients
        )
        
        # Apply filters button
        if st.button("Apply Filters"):
            st.session_state.filtered_data = filter_data(
                st.session_state.billing_data,
                selected_business_heads,
                selected_consultants,
                selected_clients,
                selected_fiscal_period
            )
            st.rerun()

# Main content area
if uploaded_file is not None:
    try:
        # Show processing message
        processing_msg = st.empty()
        processing_msg.info("Processing your Excel file, please wait...")
        
        # Process the uploaded Excel file
        billing_data, contracts_data, business_heads, consultants, clients, fiscal_periods = process_excel_data(uploaded_file)
        
        # Clear processing message
        processing_msg.empty()
        
        # Store data and filter options in session state
        st.session_state.billing_data = billing_data
        st.session_state.contracts_data = contracts_data
        st.session_state.business_heads = business_heads
        st.session_state.consultants = consultants
        st.session_state.clients = clients
        st.session_state.fiscal_periods = fiscal_periods
        st.session_state.data_loaded = True
        
        # Initial filtered data (all data)
        if 'filtered_data' not in st.session_state:
            st.session_state.filtered_data = billing_data
            
        # Display dashboard
        if st.session_state.data_loaded:
            # Contracts section
            st.header("Contract/PO Information")
            
            # Display specific columns from contracts_data with better formatting
            contracts_display = st.session_state.contracts_data.copy()
            
            # Check if the expected columns exist (using the mapping from the data_processor)
            display_columns = ['Client Name', 'Type of Work', 'PO No', 'Business Head', 'Total PO Value', 'PO Balance']
            existing_columns = [col for col in display_columns if col in contracts_display.columns]
            
            if existing_columns:
                st.dataframe(contracts_display[existing_columns])
            else:
                st.dataframe(contracts_display)
            
            # Summary metrics
            st.header("Summary Metrics")
            col1, col2, col3, col4 = st.columns(4)
            
            total_t_amt = st.session_state.filtered_data['T Amt'].sum()
            total_n_amt = st.session_state.filtered_data['N Amt'].sum()
            total_consultants = st.session_state.filtered_data['Consultant'].nunique()
            total_clients = st.session_state.filtered_data['Client'].nunique()
            
            col1.metric("Total T Amount", f"${total_t_amt:,.2f}")
            col2.metric("Total N Amount", f"${total_n_amt:,.2f}")
            col3.metric("Total Consultants", total_consultants)
            col4.metric("Total Clients", total_clients)
            
            # Time series visualization
            st.header("Monthly Financial Performance")
            time_series_chart = create_time_series_chart(st.session_state.filtered_data)
            st.plotly_chart(time_series_chart, use_container_width=True)
            
            # Quarterly analysis
            st.header("Quarterly Analysis")
            quarterly_chart = create_quarterly_chart(st.session_state.filtered_data)
            st.plotly_chart(quarterly_chart, use_container_width=True)
            
            # Annual trends
            st.header("Annual Trends")
            annual_chart = create_annual_chart(st.session_state.filtered_data)
            st.plotly_chart(annual_chart, use_container_width=True)
            
            # Hierarchical visualization
            col1, col2 = st.columns(2)
            
            with col1:
                st.header("Business Hierarchy Performance")
                hierarchy_chart = create_hierarchy_chart(st.session_state.filtered_data)
                st.plotly_chart(hierarchy_chart, use_container_width=True)
            
            with col2:
                st.header("T Amt vs N Amt Comparison")
                comparison_chart = create_comparison_chart(st.session_state.filtered_data)
                st.plotly_chart(comparison_chart, use_container_width=True)
            
            # Consultant performance
            st.header("Consultant Performance")
            consultant_chart = create_consultant_performance_chart(st.session_state.filtered_data)
            st.plotly_chart(consultant_chart, use_container_width=True)
            
            # Data viewer
            with st.expander("View Filtered Data"):
                st.dataframe(st.session_state.filtered_data)
                
                # Download filtered data as CSV
                csv = st.session_state.filtered_data.to_csv(index=False)
                st.download_button(
                    label="Download Filtered Data as CSV",
                    data=csv,
                    file_name="filtered_data.csv",
                    mime="text/csv"
                )
    
    except Exception as e:
        error_message = str(e)
        st.error(f"An error occurred while processing the file: {error_message}")
        
        # Add more specific error handling based on error type
        if "missing required columns" in error_message.lower():
            st.info("The required columns are missing in your Excel file. Please ensure your 'Consultant Billing' sheet contains: Business Head, Consultant, Client, Date, T Amt, and N Amt columns.")
            st.expander("See Column Details", expanded=True).markdown("""
            - **Business Head**: The name of the business unit head
            - **Consultant**: The name of the consultant
            - **Client**: The client name
            - **Date**: The billing date
            - **T Amt**: Total Amount
            - **N Amt**: Net Amount
            """)
        elif "required sheet" in error_message.lower():
            st.info("Please ensure your Excel file contains both 'Contracts' and 'Consultant Billing' sheets with the correct structure.")
        else:
            st.info("Please check your Excel file structure. Your file might have an unexpected format.")
            
        with st.expander("Show Detailed Error Information"):
            st.code(error_message)
else:
    # Show placeholder when no file is uploaded
    st.info("Please upload an Excel file to begin analysis.")
    
    # Sample structure explanation
    with st.expander("Expected Excel File Structure", expanded=True):
        st.markdown("""
        Your Excel file should contain the following sheets:
        
        1. **Contracts Sheet** - Contains details about contracts and POs with columns:
           - Client (Client Name)
           - Work (Type of Work)
           - PO No. (PO Number)
           - BH (Business Head)
           - Total Value (F+V) (Total PO Value)
           - Fixed Balance (PO Balance)
           
        2. **Consultant Billing Sheet** - Contains hierarchical pivot data with:
           - Business Heads (rows in ALL CAPS)
           - Consultants (indented rows under Business Heads)
           - Clients (further indented rows under Consultants)
           - Monthly billing columns with T Amt, N Amt
        
        The app will automatically process the data and generate visualizations based on the fiscal year (April to March).
        """)
        
        st.warning("Note: The app is designed to work with hierarchical consultant billing data in pivot format. Business heads are typically in ALL CAPS, with consultants and clients as indented rows.")
        
        # Example structure without nesting expanders
        st.subheader("Sample Data Structure")
        st.markdown("""
        ### Pivot Table Structure in Consultant Billing sheet:
        
        | Row Labels      | Apr-22 |       |       | May-22 |       |       | ... |
        |-----------------|--------|-------|-------|--------|-------|-------|-----|
        |                 | T Amt  | Ded   | N Amt | T Amt  | Ded   | N Amt | ... |
        | JOHN SMITH      |        |       |       |        |       |       | ... |
        | Mark Davis      | 5000   | 500   | 4500  | 6000   | 600   | 5400  | ... |
        | ABC Inc         | 3000   | 300   | 2700  | 3500   | 350   | 3150  | ... |
        | XYZ Corp        | 2000   | 200   | 1800  | 2500   | 250   | 2250  | ... |
        | Jane Doe        | 4500   | 450   | 4050  | 5200   | 520   | 4680  | ... |
        | DEF Ltd         | 4500   | 450   | 4050  | 5200   | 520   | 4680  | ... |
        | SARAH JOHNSON   |        |       |       |        |       |       | ... |
        | Tom Brown       | 6000   | 600   | 5400  | 7000   | 700   | 6300  | ... |
        | GHI Co          | 6000   | 600   | 5400  | 7000   | 700   | 6300  | ... |
        
        ### Contracts sheet structure example:
        
        | Client    | Work          | PO No.    | BH          | Total Value (F+V) | Fixed Balance |
        |-----------|---------------|-----------|-------------|-------------------|---------------|
        | ABC Inc   | Development   | PO-12345  | JOHN SMITH  | 50000             | 35000         |
        | XYZ Corp  | Consulting    | PO-67890  | SARAH JOHNSON | 75000           | 45000         |
        
        The app will try to automatically detect this structure and process the hierarchical data correctly.
        """)
