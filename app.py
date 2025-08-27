# app.py — EPL Excel → Streamlit (with Excel colors/fonts + CF approximation)
import os
import re
import math
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

st.set_page_config(page_title="EPL Workbook Browser", layout="wide")

HIDDEN_SHEETS = {"index"}   # hide these sheets
DEFAULT_WORKBOOK = "EPL macro.xlsm"

st.title("EPL Macro Workbook → Web App")
st.caption("Filter/search, quick stats, charts, static Excel colors/fonts, and approximate conditional formatting (color scales + simple numeric rules).")

# ---------------------- Data Loading ----------------------
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

def to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8")

# ---------------------- Static Excel formatting helpers ----------------------
def _argb_to_hex(argb: str) -> str:
    if not argb: return ""
    argb = str(argb)
    if len(argb) == 8: return "#" + argb[-6:]
    if len(argb) == 6: return "#" + argb
    return ""

def read_sheet_styles(xlsm_path: str, sheet_name: str):
    wb = load_workbook(xlsm_path, data_only=True)
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

# ---------------------- Conditional Formatting (approximation) ----------------------
def a1_to_col_indices(a1_range):
    m = re.match(r"\$?([A-Z]+)\$?(\d+)?:\$?([A-Z]+)\$?(\d+)?", a1_range)
    if m:
        c1, _, c2, _ = m.groups()
        i1 = column_index_from_string(c1) - 1
        i2 = column_index_from_string(c2) - 1
        return list(range(min(i1, i2), max(i1, i2)+1))
    m2 = re.match(r"\$?([A-Z]+):\$?([A-Z]+)", a1_range)
    if m2:
        c1, c2 = m2.groups()
        i1 = column_index_from_string(c1) - 1
        i2 = column_index_from_string(c2) - 1
        return list(range(min(i1, i2), max(i1, i2)+1))
    m3 = re.match(r"\$?([A-Z]+)\$?\d+", a1_range)
    if m3:
        return [column_index_from_string(m3.group(1)) - 1]
    return []

def parse_cf(ws):
    out = []
    cf = ws.conditional_formatting
    for rng in getattr(cf, "_cf_rules", {}):
        for rule in cf._cf_rules[rng]:
            kind = getattr(rule, "type", None) or getattr(rule, "rule_type", None)
            if kind == "colorScale":
                cs = getattr(rule, "colorScale", None)
                if not cs: continue
                thresholds = []
                for cfvo, color in zip(cs.cfvo, cs.color):
                    thresholds.append({
                        "type": cfvo.type,
                        "val": cfvo.val,
                        "color": _argb_to_hex(getattr(color, "rgb", "")),
                    })
                out.append({"kind":"colorScale","range":rng,"thresholds":thresholds})
            elif kind == "cellIs":
                op = getattr(rule, "operator", None)
                dxf = getattr(rule, "dxf", None)
                bg = None
                if dxf and getattr(dxf, "fill", None):
                    col = getattr(dxf.fill.fgColor, "rgb", None) or getattr(dxf.fill.start_color, "rgb", None)
                    if col: bg = _argb_to_hex(col)
                f_list = getattr(rule, "formula", None) or []
                out.append({"kind":"cellIs","range":rng,"operator":op,"formula":f_list,"bg":bg})
    return out

def apply_cf_colors(df, ws):
    rules = parse_cf(ws)
    if not rules: return None
    headers = [cell.value for cell in ws[1]]
    styler = df.style
    def affected_cols(a1):
        ws_cols = a1_to_col_indices(a1)
        cols = []
        if len(headers) == len(df.columns):
            for c in ws_cols:
                if 0 <= c < len(df.columns):
                    cols.append(df.columns[c])
        else:
            for c in ws_cols:
                if 0 <= c < len(headers):
                    hv = headers[c]
                    if hv is not None and str(hv).strip() in df.columns:
                        cols.append(str(hv).strip())
        seen, ordered = set(), []
        for x in cols:
            if x in df.columns and x not in seen:
                seen.add(x); ordered.append(x)
        return ordered
    for r in rules:
        if r["kind"]=="colorScale":
            cols = []
            for part in str(r["range"]).split():
                cols += affected_cols(part)
            for col in cols:
                if pd.api.types.is_numeric_dtype(df[col]):
                    styler = styler.background_gradient(axis=None, subset=[col])
    def mask_series(series, operator, formula_vals):
        s = pd.to_numeric(series, errors="coerce")
        try:
            if operator in ("lessThan","lt"): return s < float(formula_vals[0])
            if operator in ("lessThanOrEqual","le"): return s <= float(formula_vals[0])
            if operator in ("greaterThan","gt"): return s > float(formula_vals[0])
            if operator in ("greaterThanOrEqual","ge"): return s >= float(formula_vals[0])
            if operator in ("equal","eq"): return s == float(formula_vals[0])
            if operator == "between": return s.between(float(formula_vals[0]), float(formula_vals[1]), inclusive="both")
        except: return pd.Series(False, index=series.index)
        return pd.Series(False, index=series.index)
    for r in rules:
        if r["kind"]=="cellIs" and r.get("bg"):
            cols = []
            for part in str(r["range"]).split():
                cols += affected_cols(part)
            for col in cols:
                if col in df.columns:
                    m = mask_series(df[col], r.get("operator"), r.get("formula", []))
                    css_col = pd.Series([""]*len(df), index=df.index, dtype="object")
                    css_col[m.fillna(False)] = f"background-color:{r['bg']}"
                    styler = styler.apply(lambda s: css_col, subset=[col])
    return styler

with st.sidebar:
    st.header("Workbook")
    uploaded = st.file_uploader("Upload Excel (.xlsx / .xlsm)", type=["xlsx","xlsm"])
    if uploaded is not None:
        dfs = load_excel_all_sheets(uploaded)
        source_label="Uploaded file"; local_path=None
    else:
        if os.path.exists(DEFAULT_WORKBOOK):
            dfs = load_excel_all_sheets(DEFAULT_WORKBOOK)
            source_label=f"Local file: {DEFAULT_WORKBOOK}"; local_path=DEFAULT_WORKBOOK
        else:
            st.info("Upload an Excel file to get started."); st.stop()
    st.success(f"Loaded: {source_label}")
    sheet_names=[s for s in dfs.keys() if s and s.lower() not in HIDDEN_SHEETS]
    sheet=st.selectbox("Sheet", sorted(sheet_names))
    mode=st.radio("View mode",["Filterable","Static Excel formatting","Conditional formatting (approx)"],index=0)

df=dfs[sheet].copy()
st.subheader(f"Sheet: {sheet}")
st.write(f"Rows: **{len(df):,}** | Columns: **{len(df.columns)}**")

if mode=="Static Excel formatting":
    if uploaded is not None:
        with tempfile.NamedTemporaryFile(delete=False,suffix=".xlsm") as tmp:
            tmp.write(uploaded.getbuffer()); path=tmp.name
    else: path=local_path
    try:
        vals_df,css_df=read_sheet_styles(path,sheet)
        styler=style_dataframe(vals_df,css_df)
        st.dataframe(styler,use_container_width=True)
    except Exception as e:
        st.warning(f"Static formatting failed: {e}"); st.dataframe(df)
    st.stop()

if mode=="Conditional formatting (approx)":
    if uploaded is not None:
        with tempfile.NamedTemporaryFile(delete=False,suffix=".xlsm") as tmp:
            tmp.write(uploaded.getbuffer()); path=tmp.name
    else: path=local_path
    try:
        wb=load_workbook(path,data_only=True); ws=wb[sheet]
        styler=apply_cf_colors(df,ws)
        if styler is None: st.info("No supported CF found"); st.dataframe(df)
        else: st.dataframe(styler,use_container_width=True)
    except Exception as e:
        st.warning(f"CF failed: {e}"); st.dataframe(df)
    st.stop()

# Default filterable view
st.dataframe(df,use_container_width=True)
