
import os
import math
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
from openpyxl import load_workbook

st.set_page_config(page_title="EPL Workbook Browser", layout="wide")

# Hide these sheets from the UI (case-insensitive names)
HIDDEN_SHEETS = {"index"}  # add more names here if needed
DEFAULT_WORKBOOK = "EPL macro.xlsm"

st.title("EPL Macro Workbook → Web App")
st.caption("Browse sheets, filter, compute quick stats, chart, or render Excel colors & fonts.")

@st.cache_data(show_spinner=False)
def load_excel_all_sheets(path_or_file) -> dict:
    dfs = pd.read_excel(path_or_file, sheet_name=None, engine="openpyxl")
    cleaned = {}
    for name, df in dfs.items():
        df = df.copy()
        df.columns = [str(c).strip() for c in df.columns]
        cleaned[name] = df
    return cleaned

def detect_numeric_cols(df: pd.DataFrame):
    return [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]

def detect_datetime_cols(df: pd.DataFrame):
    return [c for c in df.columns if pd.api.types.is_datetime64_any_dtype(df[c])]

def apply_text_search(df: pd.DataFrame, query: str):
    if not query:
        return df
    obj_cols = [c for c in df.columns if pd.api.types.is_object_dtype(df[c])]
    if not obj_cols:
        return df
    mask = pd.Series(False, index=df.index)
    q = query.strip().lower()
    for c in obj_cols:
        mask = mask | df[c].astype(str).str.lower().str_contains(q, na=False)
    return df[mask]

def to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8")

def _argb_to_hex(argb: str) -> str:
    if not argb:
        return ""
    argb = str(argb)
    if len(argb) == 8:
        return "#" + argb[-6:]
    if len(argb) == 6:
        return "#" + argb
    return ""

def read_sheet_styles(xlsm_path: str, sheet_name: str):
    wb = load_workbook(xlsm_path, data_only=True)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' not found.")
    ws = wb[sheet_name]

    vals, css = [], []
    max_row, max_col = ws.max_row, ws.max_column

    for r in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
        row_vals, row_css = [], []
        for cell in r:
            row_vals.append(cell.value)
            bg = ""
            fill = getattr(cell, "fill", None)
            if fill:
                rgb = getattr(fill.fgColor, "rgb", None) or getattr(fill.start_color, "rgb", None)
                if rgb and rgb != "00000000":
                    bg = _argb_to_hex(rgb)
            f = getattr(cell, "font", None)
            css_parts = []
            if bg:
                css_parts.append(f"background-color:{bg}")
            if f:
                if f.color and getattr(f.color, "rgb", None):
                    css_parts.append(f"color:{_argb_to_hex(f.color.rgb)}")
                if f.bold:
                    css_parts.append("font-weight:bold")
                if f.italic:
                    css_parts.append("font-style:italic")
                if f.underline:
                    css_parts.append("text-decoration: underline")
            row_css.append("; ".join(css_parts))
        vals.append(row_vals)
        css.append(row_css)

    values_df = pd.DataFrame(vals)
    header_row = None
    if len(vals) >= 1:
        first = vals[0]
        non_empty = any(v is not None and str(v).strip() != "" for v in first)
        unique = len(set(first)) == len(first)
        if non_empty and unique:
            header_row = 0
    if header_row == 0:
        values_df.columns = [str(c) if c is not None else "" for c in values_df.iloc[0]]
        values_df = values_df.iloc[1:].reset_index(drop=True)
        css_df = pd.DataFrame(css).iloc[1:].reset_index(drop=True)
        css_df.columns = values_df.columns
    else:
        css_df = pd.DataFrame(css)
    return values_df, css_df

def style_dataframe(values_df, css_df):
    css_df = css_df.reindex_like(values_df).fillna("")
    styler = values_df.style
    def _apply(_df):
        return css_df
    styler = styler.apply(_apply, axis=None)
    styler = styler.set_properties(**{"white-space": "nowrap"})
    return styler

# Sidebar: choose file & sheet
with st.sidebar:
    st.header("Workbook")
    uploaded = st.file_uploader("Upload Excel (.xlsx / .xlsm)", type=["xlsx", "xlsm"])

    if uploaded is not None:
        dfs = load_excel_all_sheets(uploaded)
        source_label = "Uploaded file"
        local_path_for_styles = None
    else:
        if os.path.exists(DEFAULT_WORKBOOK):
            dfs = load_excel_all_sheets(DEFAULT_WORKBOOK)
            source_label = f"Local file: {DEFAULT_WORKBOOK}"
            local_path_for_styles = DEFAULT_WORKBOOK
        else:
            st.info("Upload an Excel file to get started.")
            st.stop()
    st.success(f"Loaded: {source_label}")

    sheet_names = [s for s in dfs.keys() if s and s.lower() not in HIDDEN_SHEETS]
    sheet = st.selectbox("Sheet", sorted(sheet_names))
    show_styles = st.checkbox("Show Excel cell colors & fonts (slower)", value=False)
    if show_styles:
        st.caption("Styled view is a static snapshot (best for viewing formatting).")

df = dfs[sheet].copy()
st.subheader(f"Sheet: {sheet}")
st.write(f"Rows: **{len(df):,}** | Columns: **{len(df.columns)}**")

with st.expander("Column overview", expanded=False):
    st.dataframe(pd.DataFrame({"Column": df.columns, "dtype": [str(df[c].dtype) for c in df.columns]}))

if show_styles:
    if uploaded is not None:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsm") as tmp:
            tmp.write(uploaded.getbuffer())
            path_for_styles = tmp.name
    else:
        path_for_styles = local_path_for_styles
    try:
        if len(df) * max(1, len(df.columns)) > 150_000:
            st.info("Sheet too large to render with styles; showing plain table instead.")
            st.dataframe(df, use_container_width=True, hide_index=True)
        else:
            vals_df, css_df = read_sheet_styles(path_for_styles, sheet)
            styler = style_dataframe(vals_df, css_df)
            st.dataframe(styler, use_container_width=True)
            st.caption("Rendered with Excel formatting (solid fills, font color/bold/italic/underline).")
    except Exception as e:
        st.warning(f"Could not render Excel styles for '{sheet}': {e}")
        st.dataframe(df, use_container_width=True, hide_index=True)
    st.stop()

# Filterable view
q = st.text_input("Quick text search (matches any text column; case-insensitive)").strip()
if q:
    # pandas .str.contains
    obj_cols = [c for c in df.columns if pd.api.types.is_object_dtype(df[c])]
    if obj_cols:
        mask = pd.Series(False, index=df.index)
        for c in obj_cols:
            mask = mask | df[c].astype(str).str.lower().str.contains(q, na=False)
        df = df[mask]

st.markdown("**Filters**")
cols_left, cols_right = st.columns(2)

num_cols = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
dt_cols  = [c for c in df.columns if pd.api.types.is_datetime64_any_dtype(df[c])]
cat_cols = [c for c in df.columns if c not in num_cols and c not in dt_cols]

with cols_left:
    if cat_cols:
        col_cat = st.selectbox("Categorical column", ["—"] + cat_cols, index=0)
        if col_cat != "—":
            uniques = df[col_cat].dropna().astype(str).unique().tolist()
            uniques = sorted(uniques)
            if len(uniques) > 3000:
                st.info("Too many unique values; try text search instead.")
            else:
                selected_values = st.multiselect("Values", options=uniques)
                if selected_values:
                    df = df[df[col_cat].astype(str).isin(selected_values)]

with cols_right:
    if num_cols:
        col_num = st.selectbox("Numeric column", ["—"] + num_cols, index=0)
        if col_num != "—":
            series = pd.to_numeric(df[col_num], errors="coerce")
            min_v, max_v = float(series.min()), float(series.max())
            if math.isfinite(min_v) and math.isfinite(max_v):
                rng = st.slider("Range", min_value=min_v, max_value=max_v, value=(min_v, max_v))
                df = df[series.between(rng[0], rng[1], inclusive="both")]
    if dt_cols:
        col_dt = st.selectbox("Datetime column", ["—"] + dt_cols, index=0)
        if col_dt != "—":
            df[col_dt] = pd.to_datetime(df[col_dt], errors="coerce")
            min_d = pd.to_datetime(df[col_dt].min())
            max_d = pd.to_datetime(df[col_dt].max())
            if pd.notna(min_d) and pd.notna(max_d):
                d1, d2 = st.date_input("Date range", value=(min_d.date(), max_d.date()))
                if isinstance(d1, tuple):
                    d1, d2 = d1
                df = df[(df[col_dt] >= pd.to_datetime(d1)) & (df[col_dt] <= pd.to_datetime(d2))]

st.divider()
st.dataframe(df, use_container_width=True, hide_index=True)

with st.expander("Quick stats (numeric columns)", expanded=False):
    if num_cols:
        st.dataframe(df[num_cols].describe().T)
    else:
        st.write("No numeric columns.")

st.download_button("Download filtered CSV", data=df.to_csv(index=False).encode("utf-8"), file_name=f"{sheet}_filtered.csv")

st.divider()
st.subheader("Quick chart")
if num_cols:
    left, right = st.columns(2)
    x_col = left.selectbox("X axis (categorical or numeric)", ["—"] + list(df.columns), index=0)
    y_col = right.selectbox("Y axis (numeric)", ["—"] + num_cols, index=0)
    if x_col != "—" and y_col != "—":
        agg = st.selectbox("Aggregation", ["sum", "mean", "count", "min", "max"], index=1)
        temp = df[[x_col, y_col]].copy()
        if agg == "count":
            temp["_ones"] = 1
            grouped = temp.groupby(x_col)["_ones"].count().reset_index(name="count")
            y_series = grouped["count"]
            x_series = grouped[x_col]
        else:
            grouped = temp.groupby(x_col)[y_col].agg(agg).reset_index()
            y_series = grouped[y_col]
            x_series = grouped[x_col]
        fig, ax = plt.subplots()
        ax.plot(range(len(y_series)), y_series)
        ax.set_xticks(range(len(x_series)))
        ax.set_xticklabels([str(v) for v in x_series], rotation=45, ha="right")
        ax.set_xlabel(str(x_col))
        ax.set_ylabel(f"{agg}({y_col})")
        ax.set_title(f"{agg} of {y_col} by {x_col}")
        st.pyplot(fig, clear_figure=True)
else:
    st.info("Select a sheet with numeric columns to enable charting.")
