
import streamlit as st
import pandas as pd
import altair as alt
from io import BytesIO

st.set_page_config(page_title="Contract & Consultant Billing Dashboard (FY 2024-25)", layout="wide")
st.title("📊 Contract & Consultant Billing Dashboard (FY 2024-25)")

def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file:
    try:
        data = pd.read_excel(uploaded_file, sheet_name=["Contracts", "Consultant Billing"])

        # --- CONTRACTS SHEET ---
        contracts_raw = data["Contracts"]
        contracts_raw.columns = contracts_raw.iloc[3]
        contracts_df = contracts_raw[4:].copy()
        contracts_df.columns = contracts_df.columns.map(str).str.strip().str.lower()

        contracts_df = contracts_df.rename(columns={
            "client name": "client",
            "po no.": "po_no",
            "total value (f +v)": "contract_value",
            "billed current year": "billed_current_year",
            "billed last year": "billed_last_year",
            "balance": "balance"
        })

        # Calculate utilization based on Balance and PO Value
        contracts_df["utilization %"] = ((1 - (contracts_df["balance"] / contracts_df["contract_value"])) * 100).round(1)
        contracts_df["balance"] = contracts_df["contract_value"] - contracts_df["billed_current_year"] - contracts_df["billed_last_year"]

        # Client Contribution (based on billed current year)
        client_contribution = contracts_df.groupby("client").agg({"billed_current_year": "sum"}).sort_values("billed_current_year", ascending=False).reset_index()

        # --- CONSULTANT BILLING SHEET ---
        billing_df = data["Consultant Billing"]
        billing_df.columns = billing_df.columns.map(str).str.strip().str.lower()
        billing_df = billing_df.rename(columns={billing_df.columns[0]: "business_head", billing_df.columns[1]: "consultant"})

        billing_df = billing_df[billing_df["consultant"].notna()]
        billing_df["consultant"] = billing_df["consultant"].astype(str).str.strip()

        # Identify T Amt and N Amt columns
        t_amt_cols = [col for col in billing_df.columns if "t amt" in col]
        n_amt_cols = [col for col in billing_df.columns if "n amt" in col]
        day_cols = [col for col in billing_df.columns if "days" in col]

        # Sum the T Amt, N Amt, Days columns
        billing_df["billed_amount"] = billing_df[t_amt_cols].apply(pd.to_numeric, errors="coerce").sum(axis=1)
        billing_df["net_amount"] = billing_df[n_amt_cols].apply(pd.to_numeric, errors="coerce").sum(axis=1)
        billing_df["billed_days"] = billing_df[day_cols].apply(pd.to_numeric, errors="coerce").sum(axis=1)

        # Assign business head from consultant level
        billing_df["business_head"] = billing_df["business_head"].where(billing_df["billed_amount"].isna()).ffill()
        billing_df = billing_df[~billing_df["business_head"].isna()]

        # Consultant summary
        consultant_summary = billing_df.groupby(["business_head", "consultant"]).agg({
            "billed_amount": "sum",
            "net_amount": "sum",
            "billed_days": "sum"
        }).reset_index()

        # Monthly trend (T Amt columns)
        month_order = ["apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec", "jan", "feb", "mar"]
        monthly_trends = []

        for _, row in billing_df.iterrows():
            consultant_name = row["consultant"]
            for col in t_amt_cols:
                month = col.split()[0].lower()
                value = pd.to_numeric(row[col], errors="coerce")
                if pd.notna(value):
                    monthly_trends.append({
                        "consultant": consultant_name,
                        "month": month,
                        "billing": value
                    })

        trend_df = pd.DataFrame(monthly_trends)
        trend_df["month"] = pd.Categorical(trend_df["month"], categories=month_order, ordered=True)

        # --- Dashboard Tabs ---
        tab1, tab2, tab3, tab4 = st.tabs([
            "📄 Contracts Summary",
            "💰 Client Contribution",
            "👥 Consultant Summary",
            "📆 Monthly Trends"
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
            st.dataframe(consultant_summary)
            st.download_button("Download Consultant Summary", convert_df_to_excel(consultant_summary), "consultant_summary.xlsx")

        with tab4:
            st.subheader("📆 Monthly Billing Trend (FY 2024-25)")
            chart = alt.Chart(trend_df).mark_line(point=True).encode(
                x="month:N",
                y="billing:Q",
                color="consultant:N",
                tooltip=["consultant", "month", "billing"]
            ).properties(width=900)
            st.altair_chart(chart, use_container_width=True)

    except Exception as e:
        st.error(f"Something went wrong: {e}")
else:
    st.info("Please upload an Excel file with 'Contracts' and 'Consultant Billing' sheets.")
