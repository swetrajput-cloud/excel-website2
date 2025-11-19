"""
Streamlit app: builds a small web UI from an uploaded Excel file.

File used (from your upload): /mnt/data/swet slab.xlsx

How to run:
1. (Optional) create a venv: python -m venv venv && source venv/bin/activate
2. Install requirements: pip install streamlit pandas openpyxl
3. Run: streamlit run streamlit_app.py

This single-file app reads the Excel workbook, lists its sheets, and for each sheet provides:
- Data preview and full table download
- Column filtering and simple queries
- Summary stats for numeric columns
- Simple plotting (line, bar, area) via streamlit's chart helpers

Modify the code to add custom pages, styling, or to export to a production-ready framework (Flask/React/Django) if you want a multi-page site.

Note: The app directly reads the path /mnt/data/swet slab.xlsx which should already exist in the environment (from your uploaded file).

"""

import streamlit as st
import pandas as pd
import io
from pathlib import Path

EXCEL_PATH = "swet slab.xlsx"


st.set_page_config(page_title="Excel → Website", layout="wide")
st.title("Excel → Website: interactive viewer")
st.markdown("Upload / preview / explore the workbook stored at: `{}`".format(EXCEL_PATH))

# Load workbook
@st.cache_data
def load_workbook(path):
    xl = pd.ExcelFile(path)
    sheets = xl.sheet_names
    dfs = {s: xl.parse(s) for s in sheets}
    return dfs

try:
    dfs = load_workbook(EXCEL_PATH)
except Exception as e:
    st.error(f"Failed to load workbook at {EXCEL_PATH}: {e}")
    st.stop()

sheet = st.sidebar.selectbox("Choose sheet", list(dfs.keys()))
df = dfs[sheet]

st.header(f"Sheet: {sheet}")

# Basic info
c1, c2, c3 = st.columns([1,1,1])
with c1:
    st.write("**Rows**")
    st.write(df.shape[0])
with c2:
    st.write("**Columns**")
    st.write(df.shape[1])
with c3:
    st.write("**Preview**")
    st.write(df.head(3))

st.subheader("Data table")
# Column selection
cols = st.multiselect("Select columns to display", options=df.columns.tolist(), default=df.columns.tolist())
filtered = df[cols]

# Filtering: allow simple equals filter on up to 3 columns
st.subheader("Filters")
filter_cols = st.multiselect("Pick columns to filter (optional)", options=df.columns.tolist())
query_df = filtered.copy()
for fc in filter_cols:
    unique_vals = query_df[fc].dropna().unique()
    if query_df[fc].dtype == 'object' or len(unique_vals) <= 30:
        sel = st.multiselect(f"Values for {fc}", options=sorted(map(str, unique_vals)), key=f"f_{fc}")
        if sel:
            query_df = query_df[query_df[fc].astype(str).isin(sel)]
    else:
        min_v = float(query_df[fc].min())
        max_v = float(query_df[fc].max())
        r = st.slider(f"Range for {fc}", min_value=min_v, max_value=max_v, value=(min_v, max_v), key=f"r_{fc}")
        query_df = query_df[(query_df[fc] >= r[0]) & (query_df[fc] <= r[1])]

st.write(f"Showing {query_df.shape[0]} rows")
st.dataframe(query_df, use_container_width=True)

# Download filtered data
@st.cache_data
def to_excel_bytes(df_in):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_in.to_excel(writer, index=False, sheet_name='export')
    return buffer.getvalue()

if st.download_button("Download filtered data as Excel", data=to_excel_bytes(query_df), file_name=f"{sheet}_filtered.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"):
    st.success("Download prepared")

# Summary stats
st.subheader("Summary statistics (numeric)")
num = query_df.select_dtypes(include='number')
if not num.empty:
    st.dataframe(num.describe().T)
else:
    st.info("No numeric columns on this sheet to summarize")

# Simple plotting
st.subheader("Quick plots")
plot_cols = st.multiselect("Choose up to 3 numeric columns to plot", options=num.columns.tolist(), max_selections=3)
if plot_cols:
    plot_type = st.selectbox("Plot type", ['line','area','bar'])
    try:
        if plot_type == 'line':
            st.line_chart(query_df[plot_cols])
        elif plot_type == 'area':
            st.area_chart(query_df[plot_cols])
        else:
            st.bar_chart(query_df[plot_cols])
    except Exception as e:
        st.error(f"Could not draw chart: {e}")

# Show whole workbook navigation
st.sidebar.markdown("---")
if st.sidebar.button("List sheets & row counts"):
    for sname, sdata in dfs.items():
        st.sidebar.write(f"{sname}: {sdata.shape[0]} rows × {sdata.shape[1]} cols")

# Footer / next steps
st.markdown("---")
st.markdown("**Next steps / customization ideas**:\n\n- Turn this into a Flask or React front-end if you need custom styling and auth.\n- Add server-side processing if the Excel is large.\n- Add charts and dashboards per sheet for domain-specific views.")

st.caption("App generated from your uploaded Excel file. Ask me to convert this into a Flask app, a multi-page React + API site, or deploy it to Streamlit Cloud / Heroku.")
