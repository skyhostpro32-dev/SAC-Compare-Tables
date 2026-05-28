import streamlit as st
import pandas as pd
from io import BytesIO

# =========================================
# PAGE CONFIG
# =========================================
st.set_page_config(
    page_title="SAP SAC Compare",
    layout="wide",
    initial_sidebar_state="expanded"
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

        <div class="sap-logo">
            SAP
        </div>

        <div class="sap-menu">
            ☰
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
# SIDEBAR NAVIGATION
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
st.markdown('<div class="main-container">', unsafe_allow_html=True)

# =========================================
# TITLE
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

            flat_values = (
                df.astype(str)
                .fillna("")
                .values
                .flatten()
            )

            for item in flat_values:

                clean_item = str(item).strip()

                if clean_item != "" and clean_item.lower() != "nan":

                    value_type = "Other"

                    lower_item = clean_item.lower()

                    if "measure" in lower_item:
                        value_type = "Measure"

                    elif "dimension" in lower_item:
                        value_type = "Dimension"

                    elif "chart" in lower_item or "widget" in lower_item:
                        value_type = "Widget"

                    values.append({
                        "Field": clean_item,
                        "Type": value_type
                    })

        except:
            pass

    return pd.DataFrame(values)

# =========================================
# PROCESS
# =========================================
if file_a and file_b:

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

        lower_value = value.lower()

        value_type = "Other"

        if "measure" in lower_value:
            value_type = "Measure"

        elif "dimension" in lower_value:
            value_type = "Dimension"

        elif "chart" in lower_value or "widget" in lower_value:
            value_type = "Widget"

        comparison_rows.append({
            "Field": value,
            "Type": value_type,
            "Exists in A": "✅" if in_a else "❌",
            "Exists in B": "✅" if in_b else "❌",
            "Status": status
        })

    result_df = pd.DataFrame(comparison_rows)

    # =========================================
    # METRICS
    # =========================================
    total = len(result_df)

    same_count = len(
        result_df[result_df["Status"] == "Same"]
    )

    diff_count = len(
        result_df[result_df["Status"] != "Same"]
    )

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
    # TABLE
    # =========================================
    st.markdown("""
    <div class="section-title">
    Comparison Result
    </div>
    """, unsafe_allow_html=True)

    st.dataframe(
        result_df,
        use_container_width=True,
        height=500
    )

    # =========================================
    # EXPORT
    # =========================================
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
        label="⬇ Download Excel Report",
        data=output,
        file_name="sac_compare_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.markdown("</div>", unsafe_allow_html=True)
