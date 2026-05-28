import streamlit as st
import pandas as pd
from io import BytesIO

# =========================================
# PAGE CONFIG
# =========================================
st.set_page_config(page_title="SAC Comparison Tool", layout="wide")

# =========================================
# CUSTOM CSS
# =========================================
st.markdown("""
<style>
    /* Dark theme background */
    .main {
        background-color: #1e1e2f;
        color: #f0f0f0;
    }
    /* Sidebar styling */
    [data-testid="stSidebar"] {
        background-color: #12121c;
    }
    /* Upload boxes */
    .stFileUploader {
        border: 1px solid #2e2e3e;
        border-radius: 8px;
        padding: 10px;
        background-color: #1a1a28;
    }
    /* Metric cards */
    div[data-testid="metric-container"] {
        background-color: #2a2a3a;
        border-radius: 8px;
        padding: 10px;
    }
    /* Buttons */
    .stDownloadButton button {
        background-color: #0078d4;
        color: white;
        border-radius: 6px;
    }
</style>
""", unsafe_allow_html=True)

# =========================================
# HEADER
# =========================================
st.markdown("<h1 style='color:#00aaff;'>SAP SAC Comparison Tool</h1>", unsafe_allow_html=True)
st.caption("Compare SAC Story / Model Excel Exports")

st.markdown("---")

# =========================================
# SIDEBAR
# =========================================
st.sidebar.header("📂 Upload SAC Excel Files")

file_a = st.sidebar.file_uploader("Upload Excel File A", type=["xlsx"])
file_b = st.sidebar.file_uploader("Upload Excel File B", type=["xlsx"])

# =========================================
# HELPER FUNCTION
# =========================================
def extract_all_values(workbook):
    values = []
    for sheet_name, df in workbook.items():
        try:
            flat_values = df.astype(str).fillna("").values.flatten()
            for item in flat_values:
                clean_item = str(item).strip()
                if clean_item and clean_item.lower() != "nan":
                    values.append({"Sheet": sheet_name, "Value": clean_item})
        except:
            pass
    return pd.DataFrame(values)

# =========================================
# MAIN LOGIC
# =========================================
if file_a and file_b:
    workbook_a = pd.read_excel(file_a, sheet_name=None)
    workbook_b = pd.read_excel(file_b, sheet_name=None)

    df_a = extract_all_values(workbook_a)
    df_b = extract_all_values(workbook_b)

    values_a = set(df_a["Value"].unique())
    values_b = set(df_b["Value"].unique())

    all_values = sorted(list(values_a.union(values_b)))
    comparison_rows = []

    for value in all_values:
        in_a = value in values_a
        in_b = value in values_b
        status = "Same" if in_a and in_b else "Missing in B" if in_a else "Missing in A"
        comparison_rows.append({
            "Field": value,
            "Exists in A": "Yes" if in_a else "No",
            "Exists in B": "Yes" if in_b else "No",
            "Status": status
        })

    result_df = pd.DataFrame(comparison_rows)

    # Metrics
    total = len(result_df)
    same_count = len(result_df[result_df["Status"] == "Same"])
    diff_count = len(result_df[result_df["Status"] != "Same"])

    col1, col2, col3 = st.columns(3)
    col1.metric("Total Fields", total)
    col2.metric("Matched", same_count)
    col3.metric("Differences", diff_count)

    st.markdown("---")

    # Filter
    filter_option = st.selectbox("Filter Results", ["All", "Same", "Missing in A", "Missing in B"])
    filtered_df = result_df if filter_option == "All" else result_df[result_df["Status"] == filter_option]

    st.subheader("📋 Comparison Result")
    st.dataframe(filtered_df, use_container_width=True, height=600)

    st.markdown("---")
    st.subheader("⚠ Difference Report")

    diff_df = result_df[result_df["Status"] != "Same"]
    if not diff_df.empty:
        st.dataframe(diff_df, use_container_width=True)
    else:
        st.success("✅ No Differences Found")

    # Export
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        result_df.to_excel(writer, index=False, sheet_name="Comparison")
        diff_df.to_excel(writer, index=False, sheet_name="Differences")
    output.seek(0)

    st.download_button(
        label="⬇ Download Comparison Report",
        data=output,
        file_name="sac_comparison_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("⬅ Upload both Excel files to start comparison")

st.markdown("---")
st.caption("SAC Story / Model Comparison Dashboard")
