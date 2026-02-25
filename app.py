import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# -------------------- PAGE CONFIG --------------------
st.set_page_config(
    page_title="Invoice Reconciliation Dashboard",
    layout="wide"
)

st.title("📊 Invoice Reconciliation Dashboard")
st.markdown("Upload Seller and Vendor SOA files to perform reconciliation.")
st.divider()

# -------------------- FILE UPLOAD --------------------
col1, col2 = st.columns(2)

with col1:
    seller_file = st.file_uploader("📘 Upload Seller SOA", type=["xlsx"])

with col2:
    vendor_file = st.file_uploader("📕 Upload Vendor SOA", type=["xlsx"])

threshold = st.number_input(
    "💰 Acceptable Difference Threshold (₹)",
    min_value=0,
    value=1000,
    step=100
)

if not seller_file or not vendor_file:
    st.info("⬆️ Please upload both files to continue.")
    st.stop()

# -------------------- READ FILES --------------------
seller_df = pd.read_excel(seller_file)
vendor_raw_df = pd.read_excel(vendor_file)

# =====================================================
# 1️⃣ PROCESS SELLER FILE (NORMAL INVOICE FORMAT)
# =====================================================

def detect_seller_columns(df):
    df.columns = df.columns.str.strip().str.lower()

    invoice_keywords = ["invoice", "document", "number", "no."]
    amount_keywords = ["amount", "amt", "total", "net"]

    invoice_col = None
    amount_col = None

    for col in df.columns:
        for key in invoice_keywords:
            if key in col:
                invoice_col = col
                break
        if invoice_col:
            break

    for col in df.columns:
        for key in amount_keywords:
            if key in col:
                amount_col = col
                break
        if amount_col:
            break

    if not invoice_col or not amount_col:
        st.error("❌ Could not detect required columns in Seller file.")
        st.stop()

    df = df.rename(columns={
        invoice_col: "Seller_Invoice_No",
        amount_col: "Seller_Invoice_Amount"
    })

    df["Seller_Invoice_Amount"] = pd.to_numeric(
        df["Seller_Invoice_Amount"], errors="coerce"
    ).fillna(0)

    return df[["Seller_Invoice_No", "Seller_Invoice_Amount"]]

seller_df = detect_seller_columns(seller_df)

# =====================================================
# 2️⃣ PROCESS VENDOR FILE (LEDGER FORMAT)
# =====================================================

def process_vendor_ledger(df):

    df.columns = df.columns.str.strip()

    required_cols = ["Vch Type", "Vch No.", "Debit", "Credit"]

    for col in required_cols:
        if col not in df.columns:
            st.error(f"❌ Required column '{col}' not found in Vendor file.")
            st.stop()

    # Keep only Purchase vouchers
    df = df[df["Vch Type"].str.lower() == "purchase"]

    # Convert Debit & Credit to numeric
    df["Debit"] = pd.to_numeric(df["Debit"], errors="coerce").fillna(0)
    df["Credit"] = pd.to_numeric(df["Credit"], errors="coerce").fillna(0)

    # Group by voucher number
    summary = df.groupby("Vch No.", as_index=False).agg({
        "Debit": "sum",
        "Credit": "sum"
    })

    # Net amount
    summary["Vendor_Invoice_Amount"] = summary["Debit"] - summary["Credit"]

    summary = summary.rename(columns={
        "Vch No.": "Vendor_Invoice_No"
    })

    return summary[["Vendor_Invoice_No", "Vendor_Invoice_Amount"]]

vendor_df = process_vendor_ledger(vendor_raw_df)

# =====================================================
# 3️⃣ MERGE
# =====================================================

recon_df = pd.merge(
    seller_df,
    vendor_df,
    left_on="Seller_Invoice_No",
    right_on="Vendor_Invoice_No",
    how="outer"
)

# -------------------- AMOUNT DIFFERENCE --------------------
recon_df["Amount_Difference"] = abs(
    recon_df["Seller_Invoice_Amount"] - recon_df["Vendor_Invoice_Amount"]
)

# -------------------- STATUS LOGIC --------------------
def get_status(row):
    if pd.isna(row["Seller_Invoice_No"]):
        return "Missing in Seller Books"
    elif pd.isna(row["Vendor_Invoice_No"]):
        return "Missing in Vendor Books"
    elif row["Amount_Difference"] == 0:
        return "Matched"
    elif row["Amount_Difference"] <= threshold:
        return "Within Threshold"
    else:
        return "Amount Mismatch"

recon_df["Status"] = recon_df.apply(get_status, axis=1)

# -------------------- FINAL TABLE --------------------
final_df = recon_df[[
    "Seller_Invoice_No",
    "Seller_Invoice_Amount",
    "Vendor_Invoice_No",
    "Vendor_Invoice_Amount",
    "Amount_Difference",
    "Status"
]]

# ----------- RESET INDEX TO START FROM 1 ------------
final_df = final_df.reset_index(drop=True)
final_df.index = final_df.index + 1
final_df.index.name = "S.No"

# -------------------- SUMMARY METRICS --------------------
st.divider()
st.subheader("📌 Reconciliation Summary")

total_invoices = len(final_df)
matched = (final_df["Status"] == "Matched").sum()
within_threshold = (final_df["Status"] == "Within Threshold").sum()
mismatch = (final_df["Status"] == "Amount Mismatch").sum()
missing_seller = (final_df["Status"] == "Missing in Seller Books").sum()
missing_vendor = (final_df["Status"] == "Missing in Vendor Books").sum()
total_difference = final_df["Amount_Difference"].sum()

m1, m2, m3, m4, m5, m6 = st.columns(6)

m1.metric("Total Invoices", total_invoices)
m2.metric("Matched", matched)
m3.metric("Within Threshold", within_threshold)
m4.metric("Mismatch (>Threshold)", mismatch)
m5.metric("Missing in Seller", missing_seller)
m6.metric("Missing in Vendor", missing_vendor)

st.metric("💰 Total Difference (₹)", f"{total_difference:,.2f}")

# -------------------- DISPLAY TABLE --------------------
st.divider()
st.subheader("📋 Detailed Reconciliation Result")
st.dataframe(final_df, use_container_width=True)

# -------------------- DOWNLOAD BUTTON --------------------
output = BytesIO()
final_df.to_excel(output, index=True, engine="openpyxl")
output.seek(0)

st.download_button(
    label="⬇️ Download Reconciliation Report",
    data=output,
    file_name="Invoice_Reconciliation_Report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
