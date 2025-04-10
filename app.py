
import streamlit as st
import pandas as pd
import altair as alt
from io import BytesIO

st.set_page_config(page_title="Contract & Billing Dashboard", layout="wide")
st.title("📊 Contract & Consultant Billing Dashboard (FY 2024-25)")

def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

def get_month_order():
    return ["apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec", "jan", "feb", "mar"]

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])
if uploaded_file:
    try:
        # Load sheets
        data = pd.read_excel(uploaded_file, sheet_name=["Contracts", "Consultant Billing"])

        # --- CONTRACTS SHEET ---
        contracts_raw = data["Contracts"]
        contracts_raw.columns = contracts_raw.iloc[3]
        contracts_df = contracts_raw[4:].copy()
        contracts_df.columns = contracts_df.columns.map(str).str.strip().str.lower()

        # Select key fields
        contracts_df = contracts_df.rename(columns={
            "client name": "client",
            "po no.": "po no.",
            "total value (f +v)": "contract_value",
            "n": "billed_amount"
        })

        contracts_df = contracts_df[["client", "po no.", "contract_value", "billed_amount"]].dropna(subset=["client", "contract_value"])
        contracts_df["contract_value"] = pd.to_numeric(contracts_df["contract_value"], errors="coerce")
        contracts_df["billed_amount"] = pd.to_numeric(contracts_df["billed_amount"], errors="coerce").fillna(0)
        contracts_df["po no."] = contracts_df["po no."].astype(str).str.strip()
        contracts_df["utilization %"] = (contracts_df["billed_amount"] / contracts_df["contract_value"] * 100).round(1)
        contracts_df["balance"] = contracts_df["contract_value"] - contracts_df["billed_amount"]

        client_contribution = contracts_df.groupby("client").agg({
            "billed_amount": "sum"
        }).sort_values("billed_amount", ascending=False).reset_index()

        # --- CONSULTANT BILLING SHEET ---
        billing_raw = data["Consultant Billing"]
        billing_raw.columns = billing_raw.iloc[8]
        billing_df = billing_raw[9:].copy()
        billing_df = billing_df.rename(columns={billing_df.columns[0]: "consultant"})
        billing_df.columns = billing_df.columns.map(str).str.strip().str.lower()

        # Monthly columns
        t_amt_cols = [col for col in billing_df.columns if "t amt" in col.lower()]
        n_amt_cols = [col for col in billing_df.columns if "n amt" in col.lower()]
        day_cols = [col for col in billing_df.columns if "day" in col.lower() and "total" not in col.lower()]

        billing_df["billed_amount"] = billing_df[t_amt_cols].apply(pd.to_numeric, errors="coerce").sum(axis=1)
        billing_df["net_amount"] = billing_df[n_amt_cols].apply(pd.to_numeric, errors="coerce").sum(axis=1)
        billing_df["billed_days"] = billing_df[day_cols].apply(pd.to_numeric, errors="coerce").sum(axis=1)

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

        # Monthly trend: sort by financial year order
        monthly_cols = [col for col in t_amt_cols if any(m in col for m in get_month_order())]
        monthly_trend = billing_df[["consultant", "business head"] + monthly_cols].copy()
        monthly_trend = monthly_trend.melt(id_vars=["consultant", "business head"], var_name="Month", value_name="Billing")
        monthly_trend.dropna(subset=["Billing"], inplace=True)
        monthly_trend["Billing"] = pd.to_numeric(monthly_trend["Billing"], errors="coerce")

        # Fix month order
        month_order = get_month_order()
        monthly_trend["Month"] = monthly_trend["Month"].str.extract(r"(" + "|".join(month_order) + ")")[0]
        monthly_trend["Month"] = pd.Categorical(monthly_trend["Month"], categories=month_order, ordered=True)
        monthly_trend = monthly_trend.dropna(subset=["Month"])

        # --- Layout ---
        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            "📄 Contracts Summary",
            "💰 Client Contribution",
            "👥 Consultant Summary",
            "📆 Monthly Trends",
            "🧑‍💼 Head Summary"
        ])

        with tab1:
            st.subheader("📄 Client & PO Level Contract Summary")
            st.dataframe(contracts_df)
            st.download_button("Download Contracts Summary", convert_df_to_excel(contracts_df), "contracts_summary.xlsx")

        with tab2:
            st.subheader("💰 Client-Wise Contribution (₹ Billed)")
            st.dataframe(client_contribution)
            st.download_button("Download Client Contribution", convert_df_to_excel(client_contribution), "client_contribution.xlsx")

        with tab3:
            st.subheader("👥 Consultant Billing Summary by Business Head")
            st.dataframe(billing_summary)
            st.download_button("Download Consultant Summary", convert_df_to_excel(billing_summary), "consultant_summary.xlsx")

        with tab4:
            st.subheader("📆 Monthly Billing Trend (FY 2024-25)")
            chart = alt.Chart(monthly_trend).mark_line(point=True).encode(
                x=alt.X("Month:N", sort=month_order),
                y="Billing:Q",
                color="consultant:N",
                tooltip=["consultant", "business head", "Month", "Billing"]
            ).properties(width=900)
            st.altair_chart(chart, use_container_width=True)

        with tab5:
            st.subheader("🧑‍💼 Rollup by Business Head")
            st.dataframe(head_summary)
            st.download_button("Download Head Summary", convert_df_to_excel(head_summary), "head_summary.xlsx")

    except Exception as e:
        st.error(f"Something went wrong: {e}")
else:
    st.info("Please upload the Excel file with 'Contracts' and 'Consultant Billing' sheets.")
