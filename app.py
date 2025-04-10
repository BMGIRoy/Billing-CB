
import streamlit as st
import pandas as pd
import altair as alt

# --- Data loading and preprocessing ---
@st.cache
def load_data():
    # Data loading and cleaning for Contract Balance and Consultant Billing will be done here
    # You should upload your final cleaned datasets here, as we are not able to run previous steps in Streamlit directly
    contracts_clean_df = pd.read_csv("path_to_contracts_data.csv")
    consultant_billing_df = pd.read_csv("path_to_consultant_billing_data.csv")
    return contracts_clean_df, consultant_billing_df

# --- Streamlit app UI ---
st.title("Contract & Consultant Billing Dashboard")

# Data loading
contracts_clean_df, consultant_billing_df = load_data()

# --- Contract Balance Tab ---
with st.beta_expander("Contract Balance"):
    # Filter options for Contract Balance
    client_filter = st.selectbox("Select Client", contracts_clean_df['Client Name'].unique())
    business_head_filter = st.selectbox("Select Business Head", contracts_clean_df['Business Head'].unique())

    # Filter data based on selected filters
    filtered_contracts = contracts_clean_df[
        (contracts_clean_df['Client Name'] == client_filter) & 
        (contracts_clean_df['Business Head'] == business_head_filter)
    ]

    st.write(f"Contract Balance for {client_filter} - {business_head_filter}")
    st.dataframe(filtered_contracts)

    # PO Utilization Visualization
    utilization_chart = alt.Chart(filtered_contracts).mark_bar().encode(
        x='Client Name:N',
        y='Utilization:Q',
        color='Business Head:N',
        tooltip=['Client Name', 'Utilization']
    )
    st.altair_chart(utilization_chart)

# --- Consultant Billing Tab ---
with st.beta_expander("Consultant Billing"):
    # Filter options for Consultant Billing
    consultant_filter = st.selectbox("Select Consultant", consultant_billing_df['Consultant'].unique())
    month_filter = st.selectbox("Select Month", ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"])

    # Filter data based on selected filters
    filtered_billing = consultant_billing_df[
        (consultant_billing_df['Consultant'] == consultant_filter)
    ]

    st.write(f"Consultant Billing for {consultant_filter} - {month_filter}")
    st.dataframe(filtered_billing)

    # Monthly Billing Trend Visualization
    trend_chart = alt.Chart(filtered_billing).mark_line().encode(
        x='Month:N',
        y='N Amt:Q',
        color='Consultant:N',
        tooltip=['Consultant', 'Month', 'N Amt']
    )
    st.altair_chart(trend_chart)

# --- Download button ---
st.download_button(
    label="Download Data",
    data=contracts_clean_df.to_csv(),
    file_name="contracts_data.csv",
    mime="text/csv"
)

