import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Reconciliation Dashboard", layout="wide")

st.title("📊 Reconciliation Dashboard")
st.markdown("Upload both ledger files (any format) to perform voucher-wise reconciliation.")
st.divider()

col1, col2 = st.columns(2)
with col1:
    seller_file = st.file_uploader("📘 File 1 — Seller / Our Books", type=["xlsx", "xls", "csv"])
with col2:
    vendor_file = st.file_uploader("📕 File 2 — Vendor / Party Books", type=["xlsx", "xls", "csv"])

threshold = st.number_input(
    "💰 Acceptable Difference Threshold (₹)",
    min_value=0, value=1000, step=100
)

if not seller_file or not vendor_file:
    st.info("⬆️ Please upload both files to continue.")
    st.stop()


# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────

def read_file(file) -> pd.DataFrame:
    """Read xlsx, xls, or csv into a raw DataFrame (no header)."""
    name = file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(file, header=None, dtype=str)
    else:
        return pd.read_excel(file, header=None, dtype=str)


def detect_header_row(raw: pd.DataFrame) -> int:
    """
    Score each of the first 20 rows by how many financial keywords it contains.
    Return the row index with the highest score (minimum score of 2 required).
    """
    KEYWORDS = [
        "type", "voucher", "vch", "invoice", "bill",
        "debit", "credit", "amount", "total", "date", "number", "no"
    ]
    best_row, best_score = 0, 0
    for i in range(min(20, len(raw))):
        row_str = " ".join(str(c).lower() for c in raw.iloc[i].values)
        score = sum(1 for kw in KEYWORDS if kw in row_str)
        if score > best_score:
            best_score, best_row = score, i
    if best_score < 2:
        raise ValueError(
            f"Could not reliably detect a header row (best score={best_score}). "
            "Please check your file format."
        )
    return best_row


def find_col(columns: list[str], priority_keywords: list[str]) -> str | None:
    """
    Return the first column whose name contains any keyword,
    in priority order (earlier keywords win).
    """
    col_lower = {c: c.lower() for c in columns}
    for kw in priority_keywords:
        for col, low in col_lower.items():
            if kw in low:
                return col
    return None


# ─────────────────────────────────────────────
# LEDGER PROCESSOR
# ─────────────────────────────────────────────

VOUCHER_KEYWORDS  = ["voucher type", "vch type", "document type", "type"]
INVOICE_KEYWORDS  = ["voucher no", "vch no", "document no", "invoice no",
                     "bill no", "invoice number", "number", "no.", "no"]
DEBIT_KEYWORDS    = ["debit"]
CREDIT_KEYWORDS   = ["credit"]
AMOUNT_KEYWORDS   = ["gross total", "gross amount", "invoice amount",
                     "taxable amount", "amount", "net amount", "total", "value"]


def process_ledger(file, side_name: str) -> pd.DataFrame:
    raw = read_file(file)

    try:
        header_row = detect_header_row(raw)
    except ValueError as e:
        st.error(f"❌ {side_name}: {e}")
        st.stop()

    df = pd.read_excel(file, header=header_row, dtype=str) \
        if not file.name.lower().endswith(".csv") \
        else pd.read_csv(file, header=header_row, dtype=str)

    df.columns = df.columns.astype(str).str.strip()
    cols = df.columns.tolist()

    # ── Column detection ──────────────────────────────────
    voucher_col = find_col(cols, VOUCHER_KEYWORDS)
    invoice_col = find_col(cols, INVOICE_KEYWORDS)
    debit_col   = find_col(cols, DEBIT_KEYWORDS)
    credit_col  = find_col(cols, CREDIT_KEYWORDS)
    amount_col  = find_col(cols, AMOUNT_KEYWORDS) if not (debit_col and credit_col) else None

    if not voucher_col:
        st.error(f"❌ {side_name}: Voucher Type column not found. Detected columns: {cols}")
        st.stop()
    if not invoice_col:
        st.error(f"❌ {side_name}: Invoice Number column not found. Detected columns: {cols}")
        st.stop()
    if not debit_col and not credit_col and not amount_col:
        st.error(f"❌ {side_name}: No amount column found. Detected columns: {cols}")
        st.stop()

    # ── Clean & coerce ────────────────────────────────────
    df["_invoice"]      = df[invoice_col].astype(str).str.strip()
    df["_voucher_type"] = df[voucher_col].astype(str).str.strip()

    # Drop blank / nan invoice rows
    df = df[~df["_invoice"].str.lower().isin(["", "nan", "none", "null"])]

    if debit_col and credit_col:
        df["_debit"]  = pd.to_numeric(df[debit_col],  errors="coerce").fillna(0)
        df["_credit"] = pd.to_numeric(df[credit_col], errors="coerce").fillna(0)
        grouped = (
            df.groupby(["_invoice", "_voucher_type"], as_index=False)
              .agg(_debit=("_debit", "sum"), _credit=("_credit", "sum"))
        )
        grouped["Invoice_Amount"] = grouped["_debit"] - grouped["_credit"]
    else:
        df["_amount"] = pd.to_numeric(df[amount_col], errors="coerce").fillna(0)
        grouped = (
            df.groupby(["_invoice", "_voucher_type"], as_index=False)
              .agg(Invoice_Amount=("_amount", "sum"))
        )

    # Seller amounts are credits → make negative after aggregation
    if side_name == "Seller":
        grouped["Invoice_Amount"] = -grouped["Invoice_Amount"].abs()

    grouped = grouped.rename(columns={
        "_invoice":      f"{side_name}_Invoice_No",
        "_voucher_type": f"{side_name}_Voucher_Type",
        "Invoice_Amount": f"{side_name}_Invoice_Amount",
    })

    return grouped[[
        f"{side_name}_Invoice_No",
        f"{side_name}_Voucher_Type",
        f"{side_name}_Invoice_Amount",
    ]]


# ─────────────────────────────────────────────
# PROCESS & MERGE
# ─────────────────────────────────────────────

seller_df = process_ledger(seller_file, "Seller")
vendor_df = process_ledger(vendor_file, "Vendor")

recon_df = pd.merge(
    seller_df,
    vendor_df,
    left_on  =["Seller_Invoice_No", "Seller_Voucher_Type"],
    right_on =["Vendor_Invoice_No", "Vendor_Voucher_Type"],
    how="outer",
)

# Accounting difference: seller (negative) + vendor (positive) should net to 0
recon_df["Amount_Difference"] = (
    recon_df["Seller_Invoice_Amount"].fillna(0)
    + recon_df["Vendor_Invoice_Amount"].fillna(0)
).abs()


def get_status(row) -> str:
    if pd.isna(row["Seller_Invoice_No"]):
        return "Missing in Seller Books"
    if pd.isna(row["Vendor_Invoice_No"]):
        return "Missing in Vendor Books"
    if row["Amount_Difference"] == 0:
        return "Matched"
    if row["Amount_Difference"] <= threshold:
        return "Within Threshold"
    return "Amount Mismatch"


recon_df["Status"] = recon_df.apply(get_status, axis=1)

# Unified invoice / voucher columns for display
recon_df["Invoice_No"]    = recon_df["Seller_Invoice_No"].fillna(recon_df["Vendor_Invoice_No"])
recon_df["Voucher_Type"]  = recon_df["Seller_Voucher_Type"].fillna(recon_df["Vendor_Voucher_Type"])

final_df = recon_df[[
    "Invoice_No",
    "Voucher_Type",
    "Seller_Invoice_Amount",
    "Vendor_Invoice_Amount",
    "Amount_Difference",
    "Status",
]].reset_index(drop=True)

final_df.index     += 1
final_df.index.name = "S.No"


# ─────────────────────────────────────────────
# UI — SUMMARY METRICS
# ─────────────────────────────────────────────

st.divider()
st.subheader("📌 Reconciliation Summary")

m1, m2, m3, m4, m5, m6 = st.columns(6)
m1.metric("Total Invoices",        len(final_df))
m2.metric("✅ Matched",             (final_df["Status"] == "Matched").sum())
m3.metric("🟡 Within Threshold",    (final_df["Status"] == "Within Threshold").sum())
m4.metric("🔴 Amount Mismatch",     (final_df["Status"] == "Amount Mismatch").sum())
m5.metric("⚠️ Missing in Seller",   (final_df["Status"] == "Missing in Seller Books").sum())
m6.metric("⚠️ Missing in Vendor",   (final_df["Status"] == "Missing in Vendor Books").sum())

st.metric(
    "💰 Total Unreconciled Difference (₹)",
    f"₹{final_df['Amount_Difference'].sum():,.2f}"
)

# ─────────────────────────────────────────────
# UI — FILTER + TABLE
# ─────────────────────────────────────────────

st.divider()
st.subheader("📋 Detailed Reconciliation Results")

status_options = ["All"] + sorted(final_df["Status"].unique().tolist())
selected_status = st.selectbox("Filter by status", status_options)

search_query = st.text_input("🔍 Search by invoice number, voucher type or status", "")

display_df = final_df.copy()

if selected_status != "All":
    display_df = display_df[display_df["Status"] == selected_status]

if search_query:
    q = search_query.lower()
    mask = (
        display_df["Invoice_No"].astype(str).str.lower().str.contains(q, na=False)
        | display_df["Voucher_Type"].astype(str).str.lower().str.contains(q, na=False)
        | display_df["Status"].astype(str).str.lower().str.contains(q, na=False)
    )
    display_df = display_df[mask]

st.caption(f"Showing {len(display_df)} of {len(final_df)} records")
st.dataframe(display_df, use_container_width=True)

# ─────────────────────────────────────────────
# DOWNLOAD
# ─────────────────────────────────────────────

output = BytesIO()
with pd.ExcelWriter(output, engine="openpyxl") as writer:
    final_df.to_excel(writer, sheet_name="Reconciliation", index=True)

    # Summary sheet
    summary_data = {
        "Metric": [
            "Total Invoices", "Matched", "Within Threshold",
            "Amount Mismatch", "Missing in Seller", "Missing in Vendor",
            "Total Difference (₹)"
        ],
        "Value": [
            len(final_df),
            (final_df["Status"] == "Matched").sum(),
            (final_df["Status"] == "Within Threshold").sum(),
            (final_df["Status"] == "Amount Mismatch").sum(),
            (final_df["Status"] == "Missing in Seller Books").sum(),
            (final_df["Status"] == "Missing in Vendor Books").sum(),
            round(final_df["Amount_Difference"].sum(), 2),
        ]
    }
    pd.DataFrame(summary_data).to_excel(writer, sheet_name="Summary", index=False)

output.seek(0)

st.download_button(
    label="⬇️ Download Full Report (.xlsx)",
    data=output,
    file_name="Reconciliation_Report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
