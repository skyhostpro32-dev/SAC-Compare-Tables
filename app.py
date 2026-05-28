import streamlit as st
import pandas as pd

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

```
with open("styles.css") as f:

    st.markdown(
        f"<style>{f.read()}</style>",
        unsafe_allow_html=True
    )
```

load_css()

# =========================================

# HEADER CONTAINER

# =========================================

st.markdown(
""" <div class="header-wrapper">

```
    <div class="sap-container">

        <img
            src="https://upload.wikimedia.org/wikipedia/commons/5/59/SAP_2011_logo.svg"
            class="sap-logo"
        >

        <div class="sap-text-area">

            <div class="sap-title">
                SAC Comparison Tool
            </div>

            <div class="sap-subtitle">
                Compare SAC Story / Model Excel Exports
            </div>

        </div>

    </div>

</div>
""",
unsafe_allow_html=True
```

)

# =========================================

# SIDEBAR

# =========================================

st.sidebar.markdown(
""" <div class="sidebar-title">
📂 Upload SAC Excel Files </div>
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

# MAIN CONTENT

# =========================================

st.markdown('<div class="main-page">', unsafe_allow_html=True)

if not file_a or not file_b:

```
st.markdown(
    """
    <div class="upload-message">
        ⬅ Upload both SAC Excel files to begin comparison
    </div>
    """,
    unsafe_allow_html=True
)
```

else:

```
st.success("Files Uploaded Successfully")
```

st.markdown("</div>", unsafe_allow_html=True)

# =========================================

# FOOTER

# =========================================

st.markdown(
""" <div class="footer">
SAC Story / Model Comparison Dashboard </div>
""",
unsafe_allow_html=True
)
