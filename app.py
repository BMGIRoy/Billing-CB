
import streamlit as st
import pandas as pd
import altair as alt
from io import BytesIO

st.set_page_config(page_title="Contract & Billing Dashboard", layout="wide")
st.title("📊 Contract & Consultant Billing Dashboard")

def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

def get_monthly_columns(columns, keyword):
    return [col for col in columns if isinstance(col, str) and keyword.lower() in col.lower()]

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])
if uploaded_file:
    try:
        data = pd.read_excel(uploaded_file, sheet_name=["Contracts", "Consultant Billing"])

        # --- Contracts Sheet ---
        contracts_raw = data["Contracts"]
        contracts_raw.columns = contracts_raw.iloc[3]
        contracts_df = contracts_raw[4:].copy()
        contracts_df.columns = contracts_df.columns.map(str).str.strip().str.lower()
        contracts_df = contracts_df.rename(columns={
            "client name": "client",
            "po no.": "po no.",
            "total value (f +v)": "contract_value"
        })
        contracts_df = contracts_df[["client", "po no.", "contract_value"]].dropna(subset=["client", "po no."])
        contracts_df["contract_value"] = pd.to_numeric(contracts_df["contract_value"], errors="coerce")
        contracts_df = contracts_df[contracts_df["contract_value"].notna()]
        contracts_df["po no."] = contracts_df["po no."].astype(str).str.strip()

        # --- Consultant Billing Sheet ---
        billing_raw = data["Consultant Billing"]
        billing_raw.columns = billing_raw.iloc[8]
        billing_df = billing_raw[9:].copy()
        billing_df = billing_df.rename(columns={billing_df.columns[0]: "consultant"})
        billing_df.columns = billing_df.columns.map(str).str.strip().str.lower()

        # Identify monthly columns
        t_amt_cols = get_monthly_columns(billing_df.columns, "t amt")
        n_amt_cols = get_monthly_columns(billing_df.columns, "n amt")
        day_cols = get_monthly_columns(billing_df.columns, "day")

        billing_df["billed_amount"] = billing_df[t_amt_cols].apply(pd.to_numeric, errors='coerce').sum(axis=1)
        billing_df["net_amount"] = billing_df[n_amt_cols].apply(pd.to_numeric, errors='coerce').sum(axis=1)
        billing_df["billed_days"] = billing_df[day_cols].apply(pd.to_numeric, errors='coerce').sum(axis=1)

        if "business head" not in billing_df.columns:
            billing_df["business head"] = "Unassigned"

        billing_summary = billing_df.groupby(["business head", "consultant"]).agg({
            "billed_amount": "sum",
            "net_amount": "sum",
            "billed_days": "sum"
        }).reset_index()

        head_summary = billing_summary.groupby("business head").agg({
            "billed_amount": "sum",
            "net_amount": "sum",
            "billed_days": "sum"
        }).reset_index()

        monthly_trend_cols = get_monthly_columns(billing_df.columns, "t amt")
        monthly_trend = billing_df[["consultant", "business head"] + monthly_trend_cols].copy()
        monthly_trend = monthly_trend.melt(id_vars=["consultant", "business head"], var_name="Month", value_name="Billing")
        monthly_trend.dropna(subset=["Billing"], inplace=True)
        monthly_trend["Billing"] = pd.to_numeric(monthly_trend["Billing"], errors="coerce")

        # Layout
        tab1, tab2, tab3, tab4 = st.tabs([
            "📄 Contract Summary",
            "👥 Consultant Summary",
            "📆 Monthly Trends",
            "🧑‍💼 Head Summary"
        ])

        with tab1:
            st.subheader("📄 Contract Summary (Unlinked)")
            st.dataframe(contracts_df.reset_index(drop=True))
            st.download_button(
                label="Download Contract Summary (Excel)",
                data=convert_df_to_excel(contracts_df),
                file_name="contract_summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        with tab2:
            st.subheader("👥 Consultant Billing Summary by Business Head")
            selected_heads = st.multiselect("Filter by Business Head", billing_summary["business head"].unique(), default=billing_summary["business head"].unique())
            st.dataframe(billing_summary[billing_summary["business head"].isin(selected_heads)].reset_index(drop=True))
            st.download_button(
                label="Download Consultant Summary (Excel)",
                data=convert_df_to_excel(billing_summary),
                file_name="consultant_summary_by_head.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        with tab3:
            st.subheader("📆 Monthly Billing Trends by Consultant and Business Head")
            trend_chart = alt.Chart(monthly_trend).mark_line(point=True).encode(
                x="Month:N",
                y="Billing:Q",
                color="consultant:N",
                tooltip=["business head", "consultant", "Month", "Billing"]
            ).properties(width=900)
            st.altair_chart(trend_chart, use_container_width=True)

        with tab4:
            st.subheader("🧑‍💼 Rollup Summary by Business Head")
            st.dataframe(head_summary)
            st.download_button(
                label="Download Head Summary (Excel)",
                data=convert_df_to_excel(head_summary),
                file_name="head_summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Something went wrong: {e}")
else:
    st.info("Please upload the Excel file with 'Contracts' and 'Consultant Billing' sheets.")
