import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(
    page_title="Invoice Reconciliation Dashboard",
    layout="wide"
)

st.title("📊 Invoice Reconciliation Dashboard")
st.markdown("Upload Seller and Vendor SOA files to perform reconciliation.")
st.divider()

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
# COMMON LEDGER PROCESSOR
# =========================================================

def process_ledger(file, side_name):

    # Detect header row
    raw = pd.read_excel(file, header=None)

    header_row = None
    for i in range(len(raw)):
        row = raw.iloc[i].astype(str).str.lower()
        if "voucher type" in row.values or "vch type" in row.values:
            header_row = i
            break

    if header_row is None:
        st.error(f"❌ Could not detect header row in {side_name} file.")
        st.stop()

    df = pd.read_excel(file, header=header_row)
    df.columns = df.columns.str.strip()

    # Keep Purchase only
    if "Voucher Type" in df.columns:
        df = df[df["Voucher Type"].astype(str).str.lower() == "purchase"]
        invoice_col = "Voucher No."
    elif "Vch Type" in df.columns:
        df = df[df["Vch Type"].astype(str).str.lower() == "purchase"]
        invoice_col = "Vch No."
    else:
        st.error(f"❌ Voucher Type column missing in {side_name} file.")
        st.stop()

    # Format 1: Gross Total present
    if "Gross Total" in df.columns:
        df["Invoice_Amount"] = pd.to_numeric(
            df["Gross Total"], errors="coerce"
        ).fillna(0)

        summary = df.groupby(invoice_col, as_index=False)["Invoice_Amount"].sum()

    # Format 2: Debit/Credit present
    elif "Debit" in df.columns and "Credit" in df.columns:

        df["Debit"] = pd.to_numeric(df["Debit"], errors="coerce").fillna(0)
        df["Credit"] = pd.to_numeric(df["Credit"], errors="coerce").fillna(0)

        summary = df.groupby(invoice_col, as_index=False).agg({
            "Debit": "sum",
            "Credit": "sum"
        })

        summary["Invoice_Amount"] = summary["Debit"] - summary["Credit"]
        summary = summary[[invoice_col, "Invoice_Amount"]]

    else:
        st.error(f"❌ Amount columns not found in {side_name} file.")
        st.stop()

    summary = summary.rename(columns={
        invoice_col: "Invoice_No",
        "Invoice_Amount": f"{side_name}_Invoice_Amount"
    })

    return summary


# Process both files
seller_df = process_ledger(seller_file, "Seller")
vendor_df = process_ledger(vendor_file, "Vendor")

# Rename invoice column for merge clarity
seller_df = seller_df.rename(columns={"Invoice_No": "Seller_Invoice_No"})
vendor_df = vendor_df.rename(columns={"Invoice_No": "Vendor_Invoice_No"})

# =========================================================
# MERGE
# =========================================================

recon_df = pd.merge(
    seller_df,
    vendor_df,
    left_on="Seller_Invoice_No",
    right_on="Vendor_Invoice_No",
    how="outer"
)

recon_df["Amount_Difference"] = abs(
    recon_df["Seller_Invoice_Amount"] +
    recon_df["Vendor_Invoice_Amount"]
)

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

final_df = recon_df[[
    "Seller_Invoice_No",
    "Seller_Invoice_Amount",
    "Vendor_Invoice_No",
    "Vendor_Invoice_Amount",
    "Amount_Difference",
    "Status"
]]

final_df = final_df.reset_index(drop=True)
final_df.index += 1
final_df.index.name = "S.No"

# =========================================================
# SUMMARY
# =========================================================

st.divider()
st.subheader("📌 Reconciliation Summary")

m1, m2, m3, m4, m5, m6 = st.columns(6)

m1.metric("Total Invoices", len(final_df))
m2.metric("Matched", (final_df["Status"] == "Matched").sum())
m3.metric("Within Threshold", (final_df["Status"] == "Within Threshold").sum())
m4.metric("Mismatch (>Threshold)", (final_df["Status"] == "Amount Mismatch").sum())
m5.metric("Missing in Seller", (final_df["Status"] == "Missing in Seller Books").sum())
m6.metric("Missing in Vendor", (final_df["Status"] == "Missing in Vendor Books").sum())

st.metric("💰 Total Difference (₹)", f"{final_df['Amount_Difference'].sum():,.2f}")

st.divider()
st.subheader("📋 Detailed Reconciliation Result")
st.dataframe(final_df, use_container_width=True)

output = BytesIO()
final_df.to_excel(output, index=True, engine="openpyxl")
output.seek(0)

st.download_button(
    label="⬇️ Download Reconciliation Report",
    data=output,
    file_name="Invoice_Reconciliation_Report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
