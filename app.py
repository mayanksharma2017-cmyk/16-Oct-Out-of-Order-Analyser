import streamlit as st
import pandas as pd
from io import BytesIO

# --- Streamlit page setup ---
st.set_page_config(page_title="Shipment Summary Analyzer", layout="wide")

st.title("📦 Shipment Summary Analyzer")
st.write("""
Upload the Excel file generated from your **MS Shipment Analyzer** app.  
This tool will:
- Exclude shipments with 2 stops  
- Exclude shipments with *no milestone received*  
- Generate summary pivot tables  
- Provide a downloadable Excel report
""")

# --- File uploader ---
uploaded_file = st.file_uploader("📤 Upload your Analyzer output (.xlsx)", type=["xlsx"])

if uploaded_file:
    # Read the uploaded Excel
    df = pd.read_excel(uploaded_file)
    st.subheader("✅ Preview of Uploaded Data")
    st.dataframe(df.head(10), use_container_width=True)

    # --- Normalize column names to avoid KeyError ---
    df.columns = df.columns.str.strip().str.lower().str.replace(" ", "_")

    # --- Validate required columns ---
    required_cols = ["num_stops", "milestone_status", "out_of_order", "origin_name", "carrier"]
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        st.error(f"⚠️ Missing required columns in uploaded file: {', '.join(missing_cols)}")
        st.stop()

    # --- Apply filters safely ---
    filtered_df = df[
        (df["num_stops"] != 2) &
        (~df["milestone_status"].str.lower().eq("no_milestone_received")) &
        (~df["milestone_status"].str.lower().eq("no milestone received"))
    ]

    st.subheader("📋 Filtered Dataset Summary")
    st.write(f"Total shipments before filtering: {len(df):,}")
    st.write(f"Total shipments after filtering: {len(filtered_df):,}")

    # --- 1️⃣ Out-of-Order Summary ---
    out_of_order_summary = (
        filtered_df.groupby(["out_of_order", "milestone_status"])
        .size()
        .unstack(fill_value=0)
    )
    out_of_order_summary.loc["Grand Total"] = out_of_order_summary.sum()
    out_of_order_summary["Grand Total"] = out_of_order_summary.sum(axis=1)

    # --- 2️⃣ Origin Summary ---
    origin_summary = (
        filtered_df.groupby(["origin_name", "out_of_order"])
        .size()
        .unstack(fill_value=0)
        .rename(columns={"No": "No", "Yes": "Yes"})
    )
    origin_summary["Grand Total"] = origin_summary.sum(axis=1)
    origin_summary["% Out of Order"] = (
        (origin_summary.get("Yes", 0) / origin_summary["Grand Total"])
        .fillna(0)
        .apply(lambda x: f"{x:.0%}")
    )
    origin_summary.loc["Grand Total"] = origin_summary.sum(numeric_only=True)
    origin_summary.at["Grand Total", "% Out of Order"] = (
        f"{origin_summary['Yes'].sum() / origin_summary['Grand Total'].sum():.0%}"
    )

    # --- 3️⃣ Carrier Summary ---
    carrier_summary = (
        filtered_df.groupby(["carrier", "origin_name", "out_of_order"])
        .size()
        .unstack(fill_value=0)
    )
    carrier_summary["Grand Total"] = carrier_summary.sum(axis=1)
    carrier_summary["% Out of Order"] = (
        (carrier_summary.get("Yes", 0) / carrier_summary["Grand Total"])
        .fillna(0)
        .apply(lambda x: f"{x:.0%}")
    )
    carrier_summary.loc["Grand Total"] = carrier_summary.sum(numeric_only=True)
    carrier_summary.at["Grand Total", "% Out of Order"] = (
        f"{carrier_summary['Yes'].sum() / carrier_summary['Grand Total'].sum():.0%}"
    )

    # --- Display results ---
    st.subheader("📊 Out-of-Order Summary")
    st.dataframe(out_of_order_summary, use_container_width=True)

    st.subheader("📍 Origin Summary")
    st.dataframe(origin_summary, use_container_width=True)

    st.subheader("🚛 Carrier Summary")
    st.dataframe(carrier_summary, use_container_width=True)

    # --- Export to Excel ---
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        filtered_df.to_excel(writer, index=False, sheet_name="Filtered_Data")

        # Write pivot summaries on one sheet
        out_of_order_summary.to_excel(writer, sheet_name="Summary_Analysis", startrow=0)
        origin_summary.to_excel(
            writer,
            sheet_name="Summary_Analysis",
            startrow=len(out_of_order_summary) + 4
        )
        carrier_summary.to_excel(
            writer,
            sheet_name="Summary_Analysis",
            startrow=len(out_of_order_summary) + len(origin_summary) + 8
        )

    st.download_button(
        label="⬇️ Download Excel Report",
        data=output.getvalue(),
        file_name="shipment_summary_analysis.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("👆 Please upload your shipment analyzer output Excel file to begin.")
