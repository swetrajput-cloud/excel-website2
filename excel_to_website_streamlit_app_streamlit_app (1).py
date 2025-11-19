"""
Interactive Streamlit app that reads your uploaded Excel workbook and provides a dynamic form
so you can input data and immediately get results.

Features:
- Loads Excel from repo-relative path `swet slab.xlsx` or fallback `/mnt/data/swet slab.xlsx` (your uploaded file).
- Lets you choose a sheet; auto-generates an input form from the sheet's columns (numeric, text, bool).
- On submit: shows the new input as a row, finds the top-3 closest existing rows (by numeric columns),
  displays them, shows updated summary statistics, and allows downloading the input + matched rows.
- Includes a file uploader fallback if the workbook isn't present in the repo.

Run:
1. pip install streamlit pandas scikit-learn openpyxl
2. streamlit run interactive_app.py

"""

import streamlit as st
import pandas as pd
import numpy as np
import io
import os
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import pairwise_distances

# Try relative path first, then fallback to the uploaded path you used earlier
DEFAULT_FILENAMES = ["swet slab.xlsx", "/mnt/data/swet slab.xlsx"]

st.set_page_config(page_title="Interactive Excel -> Results", layout="wide")
st.title("Interactive Excel-powered website")
st.markdown("Enter values in the form and get instant results (closest matches, summary stats, downloads).")

# Helper: load workbook
@st.cache_data
def load_workbook(path):
    xl = pd.ExcelFile(path)
    sheets = xl.sheet_names
    dfs = {s: xl.parse(s) for s in sheets}
    return dfs

# Choose how to load workbook
excel_path = None
for p in DEFAULT_FILENAMES:
    if os.path.exists(p):
        excel_path = p
        break

if excel_path is None:
    st.warning("Workbook not found in repo. Upload it now (or place `swet slab.xlsx` in the repo root).")
    uploaded = st.file_uploader("Upload Excel workbook", type=["xlsx"])
    if uploaded:
        excel_path = uploaded

if excel_path is None:
    st.info("No workbook available yet. Upload `swet slab.xlsx` or push it to your repo and refresh.")
    st.stop()

# Try loading
try:
    dfs = load_workbook(excel_path)
except Exception as e:
    st.error(f"Failed to load workbook: {e}")
    st.stop()

# Choose sheet
sheet = st.sidebar.selectbox("Choose sheet", list(dfs.keys()))
df = dfs[sheet].copy()

st.sidebar.write(f"Rows: {df.shape[0]}, Columns: {df.shape[1]}")

# Identify column types
numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
bool_cols = df.select_dtypes(include=[bool]).columns.tolist()
object_cols = [c for c in df.columns if c not in numeric_cols + bool_cols]

st.header(f"Sheet: {sheet}")
with st.expander("Preview data (first 5 rows)"):
    st.dataframe(df.head())

st.subheader("Input form")
st.markdown("Fill the form below and press Submit. The app will show nearest rows (based on numeric columns) and updated summaries.")

# Build dynamic form
with st.form(key="input_form"):
    inputs = {}
    # Numeric inputs
    for c in numeric_cols:
        min_v = float(np.nanmin(df[c]) if not df[c].isna().all() else 0)
        max_v = float(np.nanmax(df[c]) if not df[c].isna().all() else 100)
        mean_v = float(np.nanmean(df[c]) if not df[c].isna().all() else 0)
        inputs[c] = st.number_input(label=f"{c} (numeric)", value=mean_v, format="%.5f")
    # Boolean inputs
    for c in bool_cols:
        default = bool(df[c].mode()[0]) if not df[c].isna().all() else False
        inputs[c] = st.checkbox(label=f"{c} (bool)", value=default)
    # Text/object inputs
    for c in object_cols:
        sample = str(df[c].dropna().astype(str).iloc[0]) if not df[c].dropna().empty else ""
        inputs[c] = st.text_input(label=f"{c} (text)", value=sample)

    submitted = st.form_submit_button("Submit")

if submitted:
    st.success("Input received — processing...")

    # Build input row as DataFrame with same columns and dtypes
    input_row = pd.DataFrame(columns=df.columns)
    # Fill numeric
    for c in numeric_cols:
        input_row.loc[0, c] = inputs[c]
    for c in bool_cols:
        input_row.loc[0, c] = inputs[c]
    for c in object_cols:
        input_row.loc[0, c] = inputs[c]

    # Cast numeric columns
    for c in numeric_cols:
        input_row[c] = pd.to_numeric(input_row[c], errors='coerce')

    st.subheader("Your input (as row)")
    st.dataframe(input_row)

    # If there are numeric columns, find nearest neighbors
    if numeric_cols:
        st.subheader("Top 3 closest rows (by numeric columns)")
        # Prepare data for distance: standardize
        X_existing = df[numeric_cols].fillna(df[numeric_cols].mean())
        X_input = input_row[numeric_cols].fillna(X_existing.mean())
        scaler = StandardScaler()
        Xs = scaler.fit_transform(X_existing)
        Xi = scaler.transform(X_input)
        dists = pairwise_distances(Xs, Xi, metric='euclidean').reshape(-1)
        closest_idx = np.argsort(dists)[:3]
        matches = df.iloc[closest_idx].copy()
        matches['distance'] = dists[closest_idx]
        st.dataframe(matches)

        # Show a simple comparison chart for the first numeric column
        first_num = numeric_cols[0]
        st.line_chart(pd.concat([X_existing.reset_index(drop=True)[[first_num]].assign(type='existing'),
                                 pd.DataFrame({first_num: X_input.iloc[0].values}).assign(type='input')], ignore_index=True)[first_num])
    else:
        st.info("No numeric columns to compute closest rows. Showing exact text matches instead.")
        # Try to find rows that match text columns
        if object_cols:
            mask = pd.Series([True]*len(df))
            for c in object_cols:
                val = str(inputs[c]).strip()
                if val:
                    mask = mask & df[c].astype(str).str.contains(val, case=False, na=False)
            res = df[mask]
            st.write(f"Found {len(res)} matching rows")
            st.dataframe(res.head(10))

    # Append input to top of results and offer download
    out_df = pd.concat([input_row, df], ignore_index=True)
    st.subheader("Updated summary (numeric)")
    if not out_df.select_dtypes(include=[np.number]).empty:
        st.dataframe(out_df.select_dtypes(include=[np.number]).describe().T)

    # Prepare download
    def to_excel_bytes(df_in):
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_in.to_excel(writer, index=False, sheet_name='results')
        return buffer.getvalue()

    download_bytes = to_excel_bytes(pd.concat([input_row, matches] if numeric_cols else out_df.head(10)))
    st.download_button("Download input + matches as Excel", data=download_bytes, file_name="input_and_matches.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# Extra utilities
st.sidebar.markdown("---")
if st.sidebar.checkbox("Show full sheet data"):
    st.header("Full sheet data")
    st.dataframe(df, use_container_width=True)

st.sidebar.markdown("\nDeploy notes:\n- Add requirements.txt with: streamlit, pandas, scikit-learn, openpyxl\n- Place the Excel file `swet slab.xlsx` in repo root or upload it via the uploader in the app.")

st.caption("Interactive app created for you. Ask me to convert this to Flask + React or to tweak specific behavior (e.g., change matching method, prediction model, or form layout).")
