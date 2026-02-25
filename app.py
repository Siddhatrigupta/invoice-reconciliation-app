import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# ================== PAGE CONFIG ==================
st.set_page_config(
    page_title="Invoice Reconciliation Dashboard",
    layout="wide"
)

st.title("📊 Invoice Reconciliation Dashboard")
st.markdown("Upload Seller and Vendor SOA files to perform reconciliation.")
st.divider()

# ================== FILE UPLOAD ==================
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

# =========================================================
# FUNCTION: FIND HEADER ROW
# =========================================================
def find_header_row(file, required_keywords):
    temp_df = pd.read_excel(file, header=None)

    for i in range(len(temp_df)):
        row = temp_df.iloc[i].astype(str).str.lower()
        if all(any(keyword in cell for cell in row) for keyword in required_keywords):
            return i
    return None


# =========================================================
# PROCESS SELLER FILE
# =========================================================
seller_header = find_header_row(seller_file, ["no", "amount"])

if seller_header is None:
    st.error("❌ Could not detect header row in Seller file.")
    st.stop()

seller_df = pd.read_excel(seller_file, header=seller_header)
seller_df.columns = seller_df.columns.str.strip()

# Ensure required columns exist
if "No." not in seller_df.columns or "Amount" not in seller_df.columns:
    st.error("❌ Seller file must contain 'No.' and 'Amount' columns.")
    st.stop()

seller_df = seller_df.rename(columns={
    "No.": "Seller_Invoice_No",
    "Amount": "Seller_Invoice_Amount"
})

seller_df["Seller_Invoice_Amount"] = pd.to_numeric(
    seller_df["Seller_Invoice_Amount"], errors="coerce"
).fillna(0)

seller_df = seller_df[[
    "Seller_Invoice_No",
    "Seller_Invoice_Amount"
]]

# =========================================================
# PROCESS VENDOR FILE (HANDLES BOTH FORMATS)
# =========================================================

# First detect header row based on voucher keywords
vendor_header = find_header_row(
    vendor_file,
    ["voucher", "type"]
)

if vendor_header is None:
    # Try alternate pattern
    vendor_header = find_header_row(
        vendor_file,
        ["vch", "type"]
    )

if vendor_header is None:
    st.error("❌ Could not detect header row in Vendor file.")
    st.stop()

vendor_df = pd.read_excel(vendor_file, header=vendor_header)
vendor_df.columns = vendor_df.columns.str.strip()

# ---------------------------------------------------------
# FORMAT 1: Voucher Type / Voucher No. / Gross Total
# ---------------------------------------------------------
if (
    "Voucher Type" in vendor_df.columns and
    "Voucher No." in vendor_df.columns and
    "Gross Total" in vendor_df.columns
):

    vendor_df = vendor_df[
        vendor_df["Voucher Type"].astype(str).str.lower() == "purchase"
    ]

    vendor_df["Vendor_Invoice_Amount"] = pd.to_numeric(
        vendor_df["Gross Total"], errors="coerce"
    ).fillna(0)

    vendor_df = vendor_df.rename(columns={
        "Voucher No.": "Vendor_Invoice_No"
    })

    vendor_summary = vendor_df[[
        "Vendor_Invoice_No",
        "Vendor_Invoice_Amount"
    ]]

# ---------------------------------------------------------
# FORMAT 2: Vch Type / Vch No. / Debit / Credit
# ---------------------------------------------------------
elif (
    "Vch Type" in vendor_df.columns and
    "Vch No." in vendor_df.columns and
    "Debit" in vendor_df.columns and
    "Credit" in vendor_df.columns
):

    vendor_df = vendor_df[
        vendor_df["Vch Type"].astype(str).str.lower() == "purchase"
    ]

    vendor_df["Debit"] = pd.to_numeric(
        vendor_df["Debit"], errors="coerce"
    ).fillna(0)

    vendor_df["Credit"] = pd.to_numeric(
        vendor_df["Credit"], errors="coerce"
    ).fillna(0)

    vendor_summary = vendor_df.groupby("Vch No.", as_index=False).agg({
        "Debit": "sum",
        "Credit": "sum"
    })

    vendor_summary["Vendor_Invoice_Amount"] = (
        vendor_summary["Debit"] - vendor_summary["Credit"]
    )

    vendor_summary = vendor_summary.rename(columns={
        "Vch No.": "Vendor_Invoice_No"
    })

    vendor_summary = vendor_summary[[
        "Vendor_Invoice_No",
        "Vendor_Invoice_Amount"
    ]]

else:
    st.error("❌ Vendor file format not recognized.")
    st.write("Detected columns:", vendor_df.columns.tolist())
    st.stop()


# =========================================================
# MERGE
# =========================================================
recon_df = pd.merge(
    seller_df,
    vendor_summary,
    left_on="Seller_Invoice_No",
    right_on="Vendor_Invoice_No",
    how="outer"
)

recon_df["Amount_Difference"] = abs(
    recon_df["Seller_Invoice_Amount"] -
    recon_df["Vendor_Invoice_Amount"]
)

# =========================================================
# STATUS LOGIC
# =========================================================
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

# =========================================================
# FINAL TABLE
# =========================================================
final_df = recon_df[[
    "Seller_Invoice_No",
    "Seller_Invoice_Amount",
    "Vendor_Invoice_No",
    "Vendor_Invoice_Amount",
    "Amount_Difference",
    "Status"
]]

final_df = final_df.reset_index(drop=True)
final_df.index = final_df.index + 1
final_df.index.name = "S.No"

# =========================================================
# SUMMARY
# =========================================================
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

# =========================================================
# DISPLAY TABLE
# =========================================================
st.divider()
st.subheader("📋 Detailed Reconciliation Result")
st.dataframe(final_df, use_container_width=True)

# =========================================================
# DOWNLOAD BUTTON
# =========================================================
output = BytesIO()
final_df.to_excel(output, index=True, engine="openpyxl")
output.seek(0)

st.download_button(
    label="⬇️ Download Reconciliation Report",
    data=output,
    file_name="Invoice_Reconciliation_Report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
