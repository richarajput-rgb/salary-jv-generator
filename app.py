import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="Salary JV Generator", layout="wide")

st.title("📊 Salary JV Upload Generator")

st.write("""
Upload your **SALARY ENTRY.xlsx** file and download the generated  
**Salary_JV_Upload.xlsx** with three sheets:
- All_Branches_JV  
- HO001_JV  
- HO001_Adjustment_JV  
""")

uploaded_file = st.file_uploader("Upload SALARY ENTRY.xlsx", type=["xlsx"])

if uploaded_file:

    # -------- Read Excel (FIX FOR STREAMLIT CLOUD) --------
    file_bytes = uploaded_file.read()
    df = pd.read_excel(file_bytes, header=None, engine="openpyxl")

    # -------- Extract C1–K1 accounts for other branches --------
    account_headers = df.iloc[0, 2:11].values  # C1 to K1

    # Start from row 2 (as per your sheet)
    df = df.iloc[2:].reset_index(drop=True)

    # -------- Document date & reference --------
    doc_date = st.date_input("Select Document Date", datetime.today()).strftime("%d-%m-%Y")
    month_year = datetime.strptime(doc_date, "%d-%m-%Y").strftime("%B %Y")
    ref = f"BEING SALARY ENTRY POSTED FOR {month_year}"

    # ============================
    # COLUMN POSITIONS
    # ============================

    COL_BCODES = 1          # B
    START_EXP_COL = 2       # C
    END_EXP_COL = 10        # K

    COL_ACCOUNT = 28        # AC
    COL_SUB_ACCOUNT = 29    # AD
    COL_DEBIT = 31          # AF
    COL_CREDIT = 32         # AG

    # Trim spaces
    df[COL_BCODES] = df[COL_BCODES].astype(str).str.strip()

    # ---- FILTERS ----
    df_others = df[df[COL_BCODES].str.upper() != "HO001"].copy()
    df_ho = df[df[COL_BCODES].str.upper() == "HO001"].copy()

    # ============================
    # OUTPUT COLUMNS
    # ============================
    columns = [
        "Journal Code",
        "Sequence",
        "Account",
        "Sub Account",
        "Department",
        "Document Date",
        "Debit",
        "Credit",
        "Supplier Id",
        "Customer Id",
        "SAC/HSN",
        "Reference",
        "Branch Id",
        "Invoice Num",
        "Comments"
    ]

    # ============================
    # 1️⃣ OTHER BRANCHES LOGIC
    # ============================

    all_rows = []

    for _, r in df_others.iterrows():

        branch = r[COL_BCODES]

        if branch == "" or pd.isna(branch):
            continue

        for i, col in enumerate(range(START_EXP_COL, END_EXP_COL + 1)):

            account = account_headers[i]
            amount = r[col]

            if pd.isna(account) or pd.isna(amount) or amount == 0:
                continue

            account = int(account)

            # Debit line
            all_rows.append({
                "Journal Code": "JV",
                "Sequence": 1,
                "Account": account,
                "Sub Account": "",
                "Department": branch,
                "Document Date": doc_date,
                "Debit": amount,
                "Credit": "",
                "Supplier Id": "",
                "Customer Id": "",
                "SAC/HSN": "",
                "Reference": ref,
                "Branch Id": branch,
                "Invoice Num": "",
                "Comments": ref
            })

            # Credit line
            all_rows.append({
                "Journal Code": "JV",
                "Sequence": 2,
                "Account": 411202,
                "Sub Account": "",
                "Department": branch,
                "Document Date": doc_date,
                "Debit": "",
                "Credit": amount,
                "Supplier Id": "",
                "Customer Id": "",
                "SAC/HSN": "",
                "Reference": ref,
                "Branch Id": branch,
                "Invoice Num": "",
                "Comments": ref
            })

    # ============================
    # 2️⃣ HO001 LOGIC
    # ============================

    ho_rows = []
    seq = 1

    for _, r in df_ho.iterrows():

        account = r[COL_ACCOUNT]
        sub_acc = r[COL_SUB_ACCOUNT]
        debit_input = r[COL_DEBIT]
        credit_input = r[COL_CREDIT]

        if pd.isna(account):
            continue

        account = int(account)

        debit_amt = "" if pd.isna(debit_input) else debit_input
        credit_amt = "" if pd.isna(credit_input) else credit_input

        ho_rows.append({
            "Journal Code": "JV",
            "Sequence": seq,
            "Account": account,
            "Sub Account": "" if pd.isna(sub_acc) else sub_acc,
            "Department": "HO001",
            "Document Date": doc_date,
            "Debit": debit_amt,
            "Credit": credit_amt,
            "Supplier Id": "",
            "Customer Id": "",
            "SAC/HSN": "",
            "Reference": ref,
            "Branch Id": "HO001",
            "Invoice Num": "",
            "Comments": ref
        })

        seq += 1

    # ============================
    # 3️⃣ HO001 ADJUSTMENT SHEET
    # ============================

    adj_rows = []
    seq = 1

    # Auto-detect AK column (AJ=35 or AK=36)
    possible_ak_cols = [35, 36]

    for col in possible_ak_cols:
        tmp = (
            df[col]
            .astype(str)
            .str.strip()
            .replace("", "0")
            .replace("nan", "0")
        )
        tmp = pd.to_numeric(tmp, errors="coerce").fillna(0)

        if tmp.abs().sum() > 0:
            df["_AK_NUM"] = tmp
            break
    else:
        df["_AK_NUM"] = (
            df[36]
            .astype(str)
            .str.strip()
            .replace("", "0")
            .replace("nan", "0")
        )
        df["_AK_NUM"] = pd.to_numeric(df["_AK_NUM"], errors="coerce").fillna(0)

    # -------- Line 1: Total Debit Line --------
    total_amount = df["_AK_NUM"].sum()

    adj_rows.append({
        "Journal Code": "JV",
        "Sequence": seq,
        "Account": 413201,
        "Sub Account": "HO0005",
        "Department": "HO001",
        "Document Date": doc_date,
        "Debit": total_amount,
        "Credit": "",
        "Supplier Id": "",
        "Customer Id": "",
        "SAC/HSN": "",
        "Reference": ref,
        "Branch Id": "HO001",
        "Invoice Num": "",
        "Comments": ref
    })

    seq += 1

    # -------- Supplier-wise lines (only where AK has data) --------
    df_valid = df[df["_AK_NUM"] != 0].copy()

    for _, r in df_valid.iterrows():

        amount_ak = r["_AK_NUM"]
        supplier = r[34] if not pd.isna(r[34]) else ""

        debit_amt = ""
        credit_amt = ""

        if amount_ak > 0:
            credit_amt = amount_ak
        elif amount_ak < 0:
            debit_amt = abs(amount_ak)

        adj_rows.append({
            "Journal Code": "JV",
            "Sequence": seq,
            "Account": 316301,
            "Sub Account": "",
            "Department": "HO001",
            "Document Date": doc_date,
            "Debit": debit_amt,
            "Credit": credit_amt,
            "Supplier Id": supplier,
            "Customer Id": "",
            "SAC/HSN": "",
            "Reference": ref,
            "Branch Id": "HO001",
            "Invoice Num": "",
            "Comments": ref
        })

        seq += 1

    if "_AK_NUM" in df.columns:
        df.drop(columns=["_AK_NUM"], inplace=True)

    # ============================
    # CREATE OUTPUT FILE
    # ============================

    output_path = "/tmp/Salary_JV_Upload.xlsx"

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        pd.DataFrame(all_rows, columns=columns).to_excel(
            writer, sheet_name="All_Branches_JV", index=False
        )
        pd.DataFrame(ho_rows, columns=columns).to_excel(
            writer, sheet_name="HO001_JV", index=False
        )
        pd.DataFrame(adj_rows, columns=columns).to_excel(
            writer, sheet_name="HO001_Adjustment_JV", index=False
        )

    st.success("✅ File generated successfully!")

    with open(output_path, "rb") as f:
        st.download_button(
            label="📥 Download Salary_JV_Upload.xlsx",
            data=f,
            file_name="Salary_JV_Upload.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.write("Rows created:")
    st.write(f"All Branches: {len(all_rows)}")
    st.write(f"HO001: {len(ho_rows)}")
    st.write(f"Adjustment: {len(adj_rows)}")
