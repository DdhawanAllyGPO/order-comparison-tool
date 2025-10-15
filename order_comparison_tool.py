import streamlit as st
import pandas as pd
import re

st.title("📊 Unified Order Comparison & Forecast Alignment Tool")

st.markdown("""
Upload your three Excel workbooks below:
1. **Draft Order**  
2. **Submitted Order**  
3. **Forecast Report**  

The tool will:
- Identify all mismatches (**Added**, **Removed**, **Quantity Changed**)  
- Combine them into **one unified report**  
- Attach relevant **forecast details** based on `StationName` + `NDC`
""")

# --- File uploaders ---
draft_file = st.file_uploader("📄 Upload Draft Order Excel", type=["xlsx", "xls"])
submitted_file = st.file_uploader("📄 Upload Submitted Order Excel", type=["xlsx", "xls"])
forecast_file = st.file_uploader("📈 Upload Forecast Report Excel", type=["xlsx", "xls"])

if draft_file and submitted_file and forecast_file:
    try:
        # --- Read all Excel files ---
        draft_df = pd.read_excel(draft_file, dtype=str).fillna("")
        submitted_df = pd.read_excel(submitted_file, dtype=str).fillna("")
        forecast_df = pd.read_excel(forecast_file, dtype=str).fillna("")

        # --- Normalize NDC ---
        def normalize_ndc(ndc):
            if pd.isna(ndc) or ndc == "":
                return ""
            ndc_str = str(ndc)
            ndc_str = ndc_str.replace("\u00A0", "").strip()  # remove spaces & non-breaking spaces
            ndc_str = re.sub(r"\D", "", ndc_str)  # keep only digits
            ndc_str = ndc_str.zfill(11)  # pad to 11 digits
            return ndc_str

        draft_df["NDC"] = draft_df["NDC"].apply(normalize_ndc)
        submitted_df["NDC"] = submitted_df["NDC"].apply(normalize_ndc)
        forecast_df["NDC"] = forecast_df["NDC"].apply(normalize_ndc)

    except Exception as e:
        st.error(f"❌ Error reading Excel files: {e}")
        st.stop()

    # --- Prepare numeric columns ---
    for df in [draft_df, submitted_df]:
        if "Quantity" in df.columns:
            df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").fillna(0)

    # --- Unique key for matching ---
    draft_df["Key"] = (
        draft_df["Notes"].str.strip().str.lower()
        + "|"
        + draft_df["Name"].str.strip().str.lower()
        + "|"
        + draft_df["NDC"].astype(str).str.strip()
    )

    submitted_df["Key"] = (
        submitted_df["Notes"].str.strip().str.lower()
        + "|"
        + submitted_df["Name"].str.strip().str.lower()
        + "|"
        + submitted_df["NDC"].astype(str).str.strip()
    )

    # --- Build lookup dictionaries ---
    draft_dict = draft_df.set_index("Key")["Quantity"].to_dict()
    submitted_dict = submitted_df.set_index("Key")["Quantity"].to_dict()

    # --- Find mismatches ---
    diff_qty = [k for k in draft_dict if k in submitted_dict and draft_dict[k] != submitted_dict[k]]
    added_keys = [k for k in submitted_dict if k not in draft_dict]
    removed_keys = [k for k in draft_dict if k not in submitted_dict]

    # --- Build dataframes ---
    diff_qty_df = draft_df[draft_df["Key"].isin(diff_qty)].copy()
    diff_qty_df["Submitted Quantity"] = diff_qty_df["Key"].map(submitted_dict)
    diff_qty_df["ChangeType"] = "Quantity Changed"

    added_df = submitted_df[submitted_df["Key"].isin(added_keys)].copy()
    added_df["ChangeType"] = "Added"

    removed_df = draft_df[draft_df["Key"].isin(removed_keys)].copy()
    removed_df["ChangeType"] = "Removed"

    # --- Helper: join with forecast report ---
    def attach_forecast(order_df):
        order_df["StationNameKey"] = order_df["Notes"].str.strip().str.lower()
        order_df["NDCKey"] = order_df["NDC"].astype(str).str.strip()

        forecast_df["StationNameKey"] = forecast_df["StationName"].str.strip().str.lower()
        forecast_df["NDCKey"] = forecast_df["NDC"].astype(str).str.strip()

        merged = order_df.merge(
            forecast_df,
            how="left",
            left_on=["StationNameKey", "NDCKey"],
            right_on=["StationNameKey", "NDCKey"],
            suffixes=("", "_Forecast"),
        )
        return merged

    # --- Merge all three types ---
    unified_df = pd.concat(
        [attach_forecast(diff_qty_df), attach_forecast(added_df), attach_forecast(removed_df)],
        ignore_index=True,
    )

    # --- Forecast columns to show ---
    forecast_cols = [
        "Required Qty",
        "On Hand Qty",
        "Pending Qty",
        "Pending Treatment Qty",
        "Patient Qty",
        "Transfer In",
        "Transfer Out",
        "Net Qty",
        "PAR Min",
        "PAR Max",
        "Order Qty with PAR (in Inventory Units)",
    ]

    # --- Select display columns ---
    display_cols = [
        "ChangeType",
        "POReferenceNumber",
        "Notes",
        "Name",
        "DrugName",
        "NDC",
        "Quantity",
        "Submitted Quantity",
        "Product Description",
    ] + [col for col in forecast_cols if col in unified_df.columns]

    st.subheader("📋 Unified Comparison Report")
    st.dataframe(unified_df[display_cols])

    # --- Download single CSV ---
    st.download_button(
        "⬇️ Download Unified Report (CSV)",
        unified_df[display_cols].to_csv(index=False),
        "unified_comparison_report.csv",
    )