import streamlit as st
import pandas as pd
from io import BytesIO

# ======================================================
# PAGE CONFIG
# ======================================================
st.set_page_config(
    page_title="SAP SAC Comparison Tool",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ======================================================
# LOAD CSS
# ======================================================
def load_css():

    css = """
    <style>

    /* ======================================================
    HIDE STREAMLIT DEFAULT
    ====================================================== */

    #MainMenu {
        visibility: hidden;
    }

    footer {
        visibility: hidden;
    }

    header {
        visibility: hidden;
    }

    .block-container {
        padding-top: 0rem;
        padding-left: 0rem;
        padding-right: 0rem;
        max-width: 100%;
    }

    /* ======================================================
    BODY
    ====================================================== */

    body {
        background-color: #f5f7fa;
    }

    /* ======================================================
    SAP HEADER
    ====================================================== */

    .sap-header {

        position: fixed;
        top: 0;
        left: 0;
        right: 0;

        height: 68px;

        background: white;

        border-bottom: 1px solid #d9d9d9;

        display: flex;
        justify-content: space-between;
        align-items: center;

        padding-left: 90px;
        padding-right: 30px;

        z-index: 999;
    }

    .sap-left {
        display: flex;
        align-items: center;
        gap: 20px;
    }

    .sap-menu {
        font-size: 24px;
        cursor: pointer;
        color: #1f2937;
    }

    .sap-logo img {
        height: 42px;
    }

    .sap-title {
        font-size: 24px;
        font-weight: 600;
        color: #111827;
    }

    .sap-subtitle {
        font-size: 20px;
        color: #4b5563;
    }

    .sap-right {
        font-size: 22px;
    }

    /* ======================================================
    LEFT NAVIGATION
    ====================================================== */

    .left-nav {

        position: fixed;

        top: 68px;
        left: 0;
        bottom: 0;

        width: 72px;

        background: white;

        border-right: 1px solid #d9d9d9;

        display: flex;
        flex-direction: column;
        align-items: center;

        padding-top: 20px;

        gap: 16px;

        z-index: 998;
    }

    .nav-icon {

        width: 48px;
        height: 48px;

        display: flex;
        align-items: center;
        justify-content: center;

        border-radius: 12px;

        font-size: 22px;

        cursor: pointer;
    }

    .nav-icon:hover {
        background: #e8f0fe;
    }

    .nav-icon.active {
        background: #0070f2;
        color: white;
    }

    /* ======================================================
    MAIN
    ====================================================== */

    .main-container {

        margin-left: 90px;
        margin-top: 90px;

        padding: 30px;
    }

    .page-title {

        font-size: 34px;
        font-weight: bold;

        color: #111827;
    }

    .page-desc {

        font-size: 18px;
        color: #6b7280;

        margin-top: 10px;
        margin-bottom: 30px;
    }

    /* ======================================================
    UPLOAD BOX
    ====================================================== */

    .upload-title {

        font-size: 22px;
        font-weight: 600;

        margin-bottom: 10px;
    }

    /* ======================================================
    METRIC CARDS
    ====================================================== */

    .metric-card {

        background: white;

        border-radius: 18px;

        padding: 25px;

        border: 1px solid #e5e7eb;

        text-align: center;
    }

    .metric-card h2 {

        margin: 0;

        font-size: 42px;

        color: #0070f2;
    }

    .metric-card.green h2 {
        color: #16a34a;
    }

    .metric-card.red h2 {
        color: #dc2626;
    }

    .metric-card p {

        margin-top: 10px;

        font-size: 16px;

        color: #6b7280;
    }

    /* ======================================================
    SECTION TITLE
    ====================================================== */

    .section-title {

        font-size: 28px;
        font-weight: bold;

        margin-top: 30px;
        margin-bottom: 15px;

        color: #111827;
    }

    /* ======================================================
    DATAFRAME
    ====================================================== */

    [data-testid="stDataFrame"] {

        border-radius: 18px;

        overflow: hidden;

        border: 1px solid #d9d9d9;
    }

    /* ======================================================
    BUTTON
    ====================================================== */

    .stDownloadButton button {

        width: 100%;

        height: 52px;

        background: #0070f2;

        color: white;

        border-radius: 12px;

        border: none;

        font-size: 18px;
        font-weight: bold;
    }

    .stDownloadButton button:hover {

        background: #0057d2;
    }

    /* ======================================================
    FOOTER
    ====================================================== */

    .footer {

        margin-left: 90px;

        padding: 20px;

        text-align: center;

        color: #6b7280;

        font-size: 15px;
    }

    </style>
    """

    st.markdown(css, unsafe_allow_html=True)

load_css()

# ======================================================
# SAP HEADER
# ======================================================
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

# ======================================================
# LEFT NAVIGATION
# ======================================================
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

# ======================================================
# MAIN CONTAINER
# ======================================================
st.markdown(
    '<div class="main-container">',
    unsafe_allow_html=True
)

# ======================================================
# TITLE
# ======================================================
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

# ======================================================
# UPLOADS
# ======================================================
col1, col2 = st.columns(2)

with col1:

    st.markdown("""
    <div class="upload-title">
    Upload Excel File A
    </div>
    """, unsafe_allow_html=True)

    file_a = st.file_uploader(
        " ",
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
        " ",
        type=["xlsx"],
        key="b"
    )

# ======================================================
# FUNCTION
# ======================================================
def extract_values(workbook):

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

                item = str(item).strip()

                if item != "" and item.lower() != "nan":

                    field_type = "Other"

                    lower_item = item.lower()

                    if (
                        "measure" in lower_item
                        or "sales" in lower_item
                        or "revenue" in lower_item
                        or "profit" in lower_item
                    ):

                        field_type = "Measure"

                    elif (
                        "dimension" in lower_item
                        or "country" in lower_item
                        or "region" in lower_item
                        or "product" in lower_item
                    ):

                        field_type = "Dimension"

                    elif (
                        "widget" in lower_item
                        or "chart" in lower_item
                        or "table" in lower_item
                    ):

                        field_type = "Widget"

                    values.append({

                        "Sheet": sheet_name,
                        "Field": item,
                        "Type": field_type

                    })

        except:
            pass

    return pd.DataFrame(values)

# ======================================================
# PROCESS
# ======================================================
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

        df_a = extract_values(workbook_a)
        df_b = extract_values(workbook_b)

        values_a = set(df_a["Field"].unique())
        values_b = set(df_b["Field"].unique())

        all_values = sorted(
            list(values_a.union(values_b))
        )

        rows = []

        for value in all_values:

            in_a = value in values_a
            in_b = value in values_b

            if in_a and in_b:

                status = "Same"

            elif in_a and not in_b:

                status = "Missing in B"

            else:

                status = "Missing in A"

            field_type = "Other"

            lv = value.lower()

            if (
                "measure" in lv
                or "sales" in lv
                or "revenue" in lv
            ):

                field_type = "Measure"

            elif (
                "dimension" in lv
                or "country" in lv
                or "region" in lv
            ):

                field_type = "Dimension"

            elif (
                "widget" in lv
                or "chart" in lv
            ):

                field_type = "Widget"

            rows.append({

                "Field": value,
                "Type": field_type,
                "Exists in A": "✅" if in_a else "❌",
                "Exists in B": "✅" if in_b else "❌",
                "Status": status

            })

        result_df = pd.DataFrame(rows)

        # ======================================================
        # METRICS
        # ======================================================
        total = len(result_df)

        matched = len(
            result_df[
                result_df["Status"] == "Same"
            ]
        )

        diff = len(
            result_df[
                result_df["Status"] != "Same"
            ]
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
                <h2>{matched}</h2>
                <p>Matched</p>
            </div>
            """, unsafe_allow_html=True)

        with c3:

            st.markdown(f"""
            <div class="metric-card red">
                <h2>{diff}</h2>
                <p>Differences</p>
            </div>
            """, unsafe_allow_html=True)

        # ======================================================
        # TABLE
        # ======================================================
        st.markdown("""
        <div class="section-title">
        Comparison Result
        </div>
        """, unsafe_allow_html=True)

        st.dataframe(
            result_df,
            use_container_width=True,
            height=550
        )

        # ======================================================
        # EXPORT
        # ======================================================
        output = BytesIO()

        with pd.ExcelWriter(
            output,
            engine="openpyxl"
        ) as writer:

            result_df.to_excel(
                writer,
                index=False,
                sheet_name="Comparison"
            )

        output.seek(0)

        st.download_button(
            label="⬇ Download Comparison Report",
            data=output,
            file_name="sac_compare_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:

        st.error(f"Application Error: {e}")

else:

    st.info(
        "⬅ Upload both Excel files to start comparison"
    )

# ======================================================
# CLOSE MAIN
# ======================================================
st.markdown("</div>", unsafe_allow_html=True)

# ======================================================
# FOOTER
# ======================================================
st.markdown("""
<div class="footer">
SAP SAC Story Comparison Dashboard
</div>
""", unsafe_allow_html=True)
