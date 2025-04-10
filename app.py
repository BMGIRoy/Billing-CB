
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

st.set_page_config(page_title="Contract & Consultant Billing Dashboard (FY 2024-25)", layout="wide")

st.title("📊 Contract & Consultant Billing Dashboard (FY 2024-25)")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file:
    xl = pd.ExcelFile(uploaded_file)
    if "Contracts" in xl.sheet_names and "Consultant Billing" in xl.sheet_names:
        contracts_df = xl.parse("Contracts", skiprows=9)
        billing_df = xl.parse("Consultant Billing", skiprows=10)

        contracts_df.columns = contracts_df.columns.str.lower().str.strip()
        contracts_df.rename(columns={"po no.": "po_no", "billed current year": "billed_amount"}, inplace=True)
        contracts_df["utilization"] = contracts_df["billed_amount"] / contracts_df["contract_value"]

        st.subheader("📄 Contracts Summary")
        st.dataframe(contracts_df[["client", "po_no", "contract_value", "billed_amount", "utilization"]])

        st.subheader("💰 Client Contribution")
        client_contribution = contracts_df.groupby("client")["billed_amount"].sum().reset_index()
        client_contribution = client_contribution.sort_values("billed_amount", ascending=False)
        st.dataframe(client_contribution)

        st.subheader("📈 Monthly Billing Trend")
        month_cols = billing_df.columns[billing_df.columns.str.contains("T Amt|Net Amt", case=False)]
        billing_df_clean = billing_df[~billing_df.iloc[:, 0].astype(str).str.contains("Total", na=False)]
        billing_df_clean = billing_df_clean.rename(columns={billing_df_clean.columns[0]: "consultant"})
        trend_data = billing_df_clean[["consultant"] + list(month_cols)].copy()
        trend_data = trend_data.dropna(subset=["consultant"])
        trend_data = trend_data.groupby("consultant")[month_cols].sum().T
        st.line_chart(trend_data)

        st.subheader("👥 Consultant Summary by Business Head")
        summary_df = billing_df_clean.copy()
        summary_df["business_head"] = ""
        current_head = ""
        for i, row in summary_df.iterrows():
            if pd.isna(row["consultant"]) or row["consultant"].strip() == "":
                continue
            if row["consultant"].isupper():
                current_head = row["consultant"]
                summary_df.at[i, "business_head"] = current_head
            else:
                summary_df.at[i, "business_head"] = current_head
        summary_df = summary_df[~summary_df["consultant"].isin(summary_df["business_head"].unique())]
        summary_df["billed_amount"] = summary_df.filter(like="T Amt").sum(axis=1)
        summary_df["net_amount"] = summary_df.filter(like="Net Amt").sum(axis=1)
        summary_df["billed_days"] = summary_df.filter(like="Days").sum(axis=1)
        final_summary = summary_df[["business_head", "consultant", "billed_amount", "net_amount", "billed_days"]]
        st.dataframe(final_summary)

        st.download_button("Download Consultant Summary", final_summary.to_csv(index=False), file_name="consultant_summary.csv")
    else:
        st.error("Required sheets 'Contracts' and 'Consultant Billing' not found in the Excel file.")
