
# Contract & Consultant Billing Dashboard

This Streamlit app provides an interactive dashboard for analyzing:

- 📄 Contract data (client, PO, contract value)
- 👥 Consultant billing (with monthly breakdowns)
- 📈 Trends over time and rollups by Business Head

---

### 📁 How to Use

1. Upload your Excel file with the following sheets:
   - `Contracts` — should contain: Client, PO No., Total Value
   - `Consultant Billing` — monthly `T Amt`, `N Amt`, and `Days` columns

2. Navigate tabs to:
   - View contract summaries
   - Analyze consultant performance
   - Explore monthly billing trends
   - See business head rollups

---

### ⚙️ Requirements

Dependencies are listed in `requirements.txt`:
```
streamlit
pandas
openpyxl
altair
xlsxwriter
```

---

### 🌐 Deployment

To run this on [Streamlit Cloud](https://streamlit.io/cloud):

- Push this repo to GitHub (public)
- Link your GitHub repo on Streamlit Cloud
- Make sure the main file is named `app.py`
- Upload your Excel workbook to view your dashboard

---

Made with ❤️ by your consulting team.
