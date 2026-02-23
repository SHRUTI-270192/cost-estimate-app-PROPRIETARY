import streamlit as st
import pandas as pd
from openpyxl import Workbook
import io

st.title("Cost Estimate Generator")

uploaded_file = st.file_uploader("Upload Budgetary Offer (Excel)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    wb = Workbook()
    ws = wb.active
    ws.title = "Cost Estimate"

    ws.append(["Cost Estimate Sheet"])
    ws.append([])

    headers = ["Sl No","UCS Code","Item description","Unit",
               "Installed qty","PR Qty","Budgetary offer",
               "Estimated rate (Rs.)","Amount (Rs.)"]
    ws.append(headers)

    total = 0

    for i,row in df.iterrows():
        est_rate = row["Budgetary offer"]
        amount = est_rate * row["PR Qty"]
        total += amount

        ws.append([
            i+1,
            row["UCS Code"],
            row["Item description"],
            row["Unit"],
            row["Installed qty"],
            row["PR Qty"],
            est_rate,
            est_rate,
            amount
        ])

    ws.append(["","","","","","","","Total",total])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    st.download_button(
        label="Download Cost Estimate Excel",
        data=output,
        file_name="Cost_Estimate.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
