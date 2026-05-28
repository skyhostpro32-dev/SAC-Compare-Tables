import streamlit as st
import pandas as pd
from io import BytesIO

# =========================================
# PAGE CONFIG
# =========================================

st.set_page_config(
    page_title="SAC Comparison Tool",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =========================================
# LOAD CSS
# =========================================

def load_css():

    with open("styles.css") as f:

        st.markdown(
            f"<style>{f.read()}</style>",
            unsafe_allow_html=True
        )

load_css()

# =========================================
# HEADER
# =========================================

st.markdown(
    """
    <div class="top-header">

        <div class="header-content">

            <img
                src="https://upload.wikimedia.org/wikipedia/commons/5/59/SAP_2011_logo.svg"
                class="sap-logo"
            >

            <div class="header-text">

                <div class="main-title">
                    SAC Comparison Tool
                </div>

                <div class="sub-title">
                    Compare SAC Story / Model Excel Exports
                </div>

            </div>

        </div>

    </div>
    """,
    unsafe_allow_html=True
)

# =========================================
# SIDEBAR
# =========================================

st.sidebar.markdown(
    """
    <div class="sidebar-title">
        📂 Upload SAC Excel Files
    </div>
    """,
    unsafe_allow_html=True
)

file_a = st.sidebar.file_uploader(
    "Upload Excel File A",
    type=["xlsx"]
)

file_b = st.sidebar.file_uploader(
    "Upload Excel File B",
    type=["xlsx"]
)

# =========================================
# EXTRACT VALUES
# =========================================

def extract_all_values(workbook):

    values = []

    for sheet_name, df in workbook.items():

        try:

            flat_values = (
                df.astype(str)
                .fillna("")
                .values
                .flatten()
            )

            for item in flat_values:

                clean_item = str(item).strip()

                if clean_item != "" and clean_item.lower() != "nan":

                    values.append({
                        "Sheet": sheet_name,
                        "Field": clean_item
                    })

        except:
            pass

    return pd.DataFrame(values)

# =========================================
# MAIN
# =========================================

st.markdown('<div class="main-page">', unsafe_allow_html=True)

if file_a and file_b:

    try:

        workbook_a = pd.read_excel(
            file_a,
            sheet_name=None
        )

        workbook_b = pd.read_excel(
            file_b,
            sheet_name=None
        )

        df_a = extract_all_values(workbook_a)
        df_b = extract_all_values(workbook_b)

        values_a = set(df_a["Field"].unique())
        values_b = set(df_b["Field"].unique())

        all_values = sorted(
            list(values_a.union(values_b))
        )

        comparison_rows = []

        for value in all_values:

            in_a = value in values_a
            in_b = value in values_b

            if in_a and in_b:
                status = "Same"

            elif in_a and not in_b:
                status = "Missing in B"

            else:
                status = "Missing in A"

            comparison_rows.append({
                "Field": value,
                "Exists in A": "✅" if in_a else "❌",
                "Exists in B": "✅" if in_b else "❌",
                "Status": status
            })

        result_df = pd.DataFrame(comparison_rows)

        st.markdown(
            """
            <div class="section-title">
                📋 Comparison Result
            </div>
            """,
            unsafe_allow_html=True
        )

        st.dataframe(
            result_df,
            use_container_width=True,
            height=500
        )

        output = BytesIO()

        with pd.ExcelWriter(
            output,
            engine="openpyxl"
        ) as writer:

            result_df.to_excel(
                writer,
                index=False
            )

        output.seek(0)

        st.download_button(
            label="⬇ Download Comparison Report",
            data=output,
            file_name="sac_comparison_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:

        st.error(f"Error: {e}")

else:

    st.markdown(
        """
        <div class="upload-message">
            ⬅ Upload both SAC Excel files to begin comparison
        </div>
        """,
        unsafe_allow_html=True
    )

st.markdown("</div>", unsafe_allow_html=True)

# =========================================
# FOOTER
# =========================================

st.markdown(
    """
    <div class="footer">
        SAC Story / Model Comparison Dashboard
    </div>
    """,
    unsafe_allow_html=True
)
