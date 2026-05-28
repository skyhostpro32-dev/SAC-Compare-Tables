import streamlit as st
import pandas as pd
from io import BytesIO

# =========================================
# PAGE CONFIG
# =========================================
st.set_page_config(
    page_title="SAP SAC Comparison Tool",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# =========================================
# LOAD CSS
# =========================================
def load_css():

    try:

        with open("styles.css") as f:

            st.markdown(
                f"<style>{f.read()}</style>",
                unsafe_allow_html=True
            )

    except:
        pass

load_css()

# =========================================
# TOP SAP HEADER
# =========================================
st.markdown("""
<div class="sap-header">

    <div class="sap-left">

        <div class="sap-menu">
            ☰
        </div>

        <div class="sap-logo">
            <img src="https://upload.wikimedia.org/wikipedia/commons/5/59/SAP_2011_logo.svg">
        </div>

        <div class="sap-title">
            Stories
        </div>

        <div class="sap-subtitle">
            New Story ▼
        </div>

    </div>

    <div class="sap-right">
        ☆
    </div>

</div>
""", unsafe_allow_html=True)

# =========================================
# LEFT NAVIGATION
# =========================================
st.markdown("""
<div class="left-nav">

    <div class="nav-icon active">
        🏠
    </div>

    <div class="nav-icon">
        📂
    </div>

</div>
""", unsafe_allow_html=True)

# =========================================
# MAIN CONTAINER
# =========================================
st.markdown(
    '<div class="main-container">',
    unsafe_allow_html=True
)

# =========================================
# PAGE TITLE
# =========================================
st.markdown("""
<div class="page-title">
📊 SAC Story / Model Comparison Tool
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="page-desc">
Compare Measures, Dimensions, Widgets and Missing Fields between two SAC Excel exports.
</div>
""", unsafe_allow_html=True)

# =========================================
# UPLOAD SECTION
# =========================================
col1, col2 = st.columns(2)

with col1:

    st.markdown("""
    <div class="upload-title">
    Upload Excel File A
    </div>
    """, unsafe_allow_html=True)

    file_a = st.file_uploader(
        "",
        type=["xlsx"],
        key="a"
    )

with col2:

    st.markdown("""
    <div class="upload-title">
    Upload Excel File B
    </div>
    """, unsafe_allow_html=True)

    file_b = st.file_uploader(
        "",
        type=["xlsx"],
        key="b"
    )

# =========================================
# EXTRACT VALUES
# =========================================
def extract_all_values(workbook):

    values = []

    for sheet_name, df in workbook.items():

        try:

            df = df.fillna("")

            flat_values = (
                df.astype(str)
                .values
                .flatten()
            )

            for item in flat_values:

                clean_item = str(item).strip()

                if (
                    clean_item != ""
                    and clean_item.lower() != "nan"
                ):

                    value_type = "Other"

                    lower_item = clean_item.lower()

                    # =================================
                    # DETECT TYPE
                    # =================================
                    if (
                        "measure" in lower_item
                        or "revenue" in lower_item
                        or "sales" in lower_item
                        or "profit" in lower_item
                        or "cost" in lower_item
                    ):

                        value_type = "Measure"

                    elif (
                        "dimension" in lower_item
                        or "country" in lower_item
                        or "region" in lower_item
                        or "date" in lower_item
                        or "product" in lower_item
                    ):

                        value_type = "Dimension"

                    elif (
                        "chart" in lower_item
                        or "widget" in lower_item
                        or "table" in lower_item
                    ):

                        value_type = "Widget"

                    values.append({

                        "Sheet": sheet_name,
                        "Field": clean_item,
                        "Type": value_type

                    })

        except:
            pass

    return pd.DataFrame(values)

# =========================================
# PROCESS FILES
# =========================================
if file_a and file_b:

    try:

        # =========================================
        # READ EXCEL
        # =========================================
        workbook_a = pd.read_excel(
            file_a,
            sheet_name=None
        )

        workbook_b = pd.read_excel(
            file_b,
            sheet_name=None
        )

        # =========================================
        # EXTRACT DATA
        # =========================================
        df_a = extract_all_values(workbook_a)
        df_b = extract_all_values(workbook_b)

        # =========================================
        # UNIQUE VALUES
        # =========================================
        values_a = set(df_a["Field"].unique())
        values_b = set(df_b["Field"].unique())

        all_values = sorted(
            list(values_a.union(values_b))
        )

        # =========================================
        # COMPARE
        # =========================================
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

            lower_value = value.lower()

            value_type = "Other"

            if (
                "measure" in lower_value
                or "revenue" in lower_value
                or "sales" in lower_value
                or "profit" in lower_value
                or "cost" in lower_value
            ):

                value_type = "Measure"

            elif (
                "dimension" in lower_value
                or "country" in lower_value
                or "region" in lower_value
                or "date" in lower_value
                or "product" in lower_value
            ):

                value_type = "Dimension"

            elif (
                "chart" in lower_value
                or "widget" in lower_value
                or "table" in lower_value
            ):

                value_type = "Widget"

            comparison_rows.append({

                "Field": value,
                "Type": value_type,
                "Exists in A": "✅" if in_a else "❌",
                "Exists in B": "✅" if in_b else "❌",
                "Status": status

            })

        # =========================================
        # RESULT DATAFRAME
        # =========================================
        result_df = pd.DataFrame(comparison_rows)

        # =========================================
        # METRICS
        # =========================================
        total = len(result_df)

        same_count = len(
            result_df[
                result_df["Status"] == "Same"
            ]
        )

        diff_count = len(
            result_df[
                result_df["Status"] != "Same"
            ]
        )

        # =========================================
        # METRIC CARDS
        # =========================================
        c1, c2, c3 = st.columns(3)

        with c1:

            st.markdown(f"""
            <div class="metric-card">
                <h2>{total}</h2>
                <p>Total Fields</p>
            </div>
            """, unsafe_allow_html=True)

        with c2:

            st.markdown(f"""
            <div class="metric-card green">
                <h2>{same_count}</h2>
                <p>Matched</p>
            </div>
            """, unsafe_allow_html=True)

        with c3:

            st.markdown(f"""
            <div class="metric-card red">
                <h2>{diff_count}</h2>
                <p>Differences</p>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # =========================================
        # FILTER
        # =========================================
        filter_option = st.selectbox(
            "Filter Result",
            [
                "All",
                "Same",
                "Missing in A",
                "Missing in B"
            ]
        )

        if filter_option == "All":

            filtered_df = result_df

        else:

            filtered_df = result_df[
                result_df["Status"] == filter_option
            ]

        # =========================================
        # RESULT TABLE
        # =========================================
        st.markdown("""
        <div class="section-title">
        Comparison Result
        </div>
        """, unsafe_allow_html=True)

        st.dataframe(
            filtered_df,
            use_container_width=True,
            height=500
        )

        # =========================================
        # DIFFERENCE TABLE
        # =========================================
        st.markdown("""
        <div class="section-title">
        Differences
        </div>
        """, unsafe_allow_html=True)

        diff_df = result_df[
            result_df["Status"] != "Same"
        ]

        if not diff_df.empty:

            st.dataframe(
                diff_df,
                use_container_width=True,
                height=300
            )

        else:

            st.success("✅ No Differences Found")

        # =========================================
        # EXPORT EXCEL
        # =========================================
        output = BytesIO()

        with pd.ExcelWriter(
            output,
            engine="openpyxl"
        ) as writer:

            result_df.to_excel(
                writer,
                sheet_name="Comparison",
                index=False
            )

            diff_df.to_excel(
                writer,
                sheet_name="Differences",
                index=False
            )

        output.seek(0)

        st.download_button(
            label="⬇ Download Excel Report",
            data=output,
            file_name="sac_compare_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:

        st.error(f"Application Error: {e}")

# =========================================
# NO FILES
# =========================================
else:

    st.info(
        "⬅ Upload both Excel files to start comparison"
    )

# =========================================
# FOOTER
# =========================================
st.markdown("</div>", unsafe_allow_html=True)

st.markdown("""
<div class="footer">
SAP SAC Story Comparison Dashboard
</div>
""", unsafe_allow_html=True)
