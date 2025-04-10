
import streamlit as st
import pandas as pd
import altair as alt
from io import BytesIO

st.set_page_config(page_title="Contract & Consultant Billing Dashboard", layout="wide")

def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

st.title("📊 Contract & Consultant Billing Dashboard (FY 2024-25)")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])
if uploaded_file:
    try:
        data = pd.read_excel(uploaded_file, sheet_name=["Contracts", "Consultant Billing"])

        # CONTRACTS
        contracts_raw = data["Contracts"]
        contracts_raw.columns = contracts_raw.iloc[3]
        contracts_df = contracts_raw[4:].copy()
        contracts_df.columns = contracts_df.columns.map(str).str.strip().str.lower()
        contracts_df = contracts_df.rename(columns={
            "client name": "client",
            "po no.": "po no.",
            "total value (f +v)": "contract_value",
            "billed current year": "billed_amount"
        })

        contracts_df = contracts_df[["client", "po no.", "contract_value", "billed_amount"]].dropna(subset=["client", "contract_value"])
        contracts_df["contract_value"] = pd.to_numeric(contracts_df["contract_value"], errors="coerce")
        contracts_df["billed_amount"] = pd.to_numeric(contracts_df["billed_amount"], errors="coerce").fillna(0)
        contracts_df["utilization %"] = (contracts_df["billed_amount"] / contracts_df["contract_value"] * 100).round(1)
        contracts_df["balance"] = contracts_df["contract_value"] - contracts_df["billed_amount"]

        client_contribution = contracts_df.groupby("client").agg({
            "billed_amount": "sum"
        }).sort_values("billed_amount", ascending=False).reset_index()

        # CONSULTANT BILLING PIVOTED
        billing_df = data["Consultant Billing"].copy()
        billing_df.columns = billing_df.columns.map(str).str.strip().str.lower()
        billing_df = billing_df.rename(columns={billing_df.columns[0]: "business head", billing_df.columns[1]: "consultant"})
        billing_df = billing_df[billing_df["consultant"].notna()]

        t_amt_cols = [c for c in billing_df.columns if "t amt" in c]
        n_amt_cols = [c for c in billing_df.columns if "n amt" in c]
        days_cols = [c for c in billing_df.columns if "days" in c]

        billing_df["billed_amount"] = billing_df[t_amt_cols].apply(pd.to_numeric, errors="coerce").sum(axis=1)
        billing_df["net_amount"] = billing_df[n_amt_cols].apply(pd.to_numeric, errors="coerce").sum(axis=1)
        billing_df["billed_days"] = billing_df[days_cols].apply(pd.to_numeric, errors="coerce").sum(axis=1)

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

        # Monthly Trend
        trend_cols = t_amt_cols
        monthly = billing_df[["consultant", "business head"] + trend_cols].copy()
        monthly = monthly.melt(id_vars=["consultant", "business head"], var_name="month", value_name="billing")
        monthly["month"] = monthly["month"].str.extract(r"(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)")[0]
        monthly["month"] = pd.Categorical(monthly["month"],
                                          categories=["apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec", "jan", "feb", "mar"],
                                          ordered=True)
        monthly["billing"] = pd.to_numeric(monthly["billing"], errors="coerce")

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
            st.subheader("📆 Monthly Billing Trend")
            chart = alt.Chart(monthly.dropna()).mark_line(point=True).encode(
                x="month:N",
                y="billing:Q",
                color="consultant:N",
                tooltip=["consultant", "business head", "month", "billing"]
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
