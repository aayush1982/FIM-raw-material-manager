# app.py — FIM Raw Material Manager (Rows = PO lines with attached Received_Qty per receiver)
# Minimal UI: KPIs + one table (no weight aggregation in the table, only Qty summed).
# Install: pip install streamlit pandas plotly xlsxwriter

import numpy as np
import pandas as pd
import streamlit as st
import plotly.graph_objects as go  # ⬅️ Plotly

st.set_page_config(page_title="FIM Raw Material Manager", layout="wide")
st.title("📦 FIM Raw Material Dashboard")

# ----------------------------
# Helpers
# ----------------------------
def _lower_map(df: pd.DataFrame) -> dict:
    return {str(c).strip().lower(): c for c in df.columns}

def _pick(cols_map: dict, *cands):
    for c in cands:
        key = str(c).strip().lower()
        if key in cols_map:
            return cols_map[key]
    return None

def _to_int_qty(s):
    return pd.to_numeric(s, errors="coerce").fillna(0).round(0).astype("Int64")

def _to_mt2(s):
    return pd.to_numeric(s, errors="coerce").fillna(0).round(2)

# ----------------------------
# Parsers for 3 sheets (1=PO, 2=Core_Fab, 3=Vrinda)
# ----------------------------
def parse_po(df: pd.DataFrame) -> pd.DataFrame:
    m = _lower_map(df)
    out = pd.DataFrame()

    # Supplier must come from PO column SUPPLIER
    c_supplier = _pick(m, "supplier")
    out["supplier"] = df[c_supplier].astype(str).str.strip() if c_supplier else ""

    c_t = _pick(m, "t (mm)", "t")
    c_l = _pick(m, "l (mm)", "l")
    c_w = _pick(m, "w (mm)", "w")
    c_qty = _pick(m, "qty (nos)", "qty")
    c_wkg = _pick(m, "weight (kg)", "weight")

    out["T_mm"] = pd.to_numeric(df[c_t], errors="coerce") if c_t else np.nan
    out["L_mm"] = pd.to_numeric(df[c_l], errors="coerce") if c_l else np.nan
    out["W_mm"] = pd.to_numeric(df[c_w], errors="coerce") if c_w else np.nan
    out["PO_Qty"] = _to_int_qty(df[c_qty]) if c_qty else pd.Series([0]*len(df), dtype="Int64")

    # Convert KG → MT (2 decimals)
    out["PO_Weight_MT"] = _to_mt2(df[c_wkg] / 1000.0) if c_wkg else _to_mt2(0)

    out["record_type"] = "PO"
    return out.dropna(how="all")

def parse_core_fab(df: pd.DataFrame) -> pd.DataFrame:
    m = _lower_map(df)
    out = pd.DataFrame()

    c_t = _pick(m, "t (mm)", "t")
    c_l = _pick(m, "l (mm)", "l")
    c_w = _pick(m, "w (mm)", "w")
    c_qty = _pick(m, "qty (nos)", "qty")
    c_wtmt = _pick(m, "inv wt (mt)", "inv wt")

    out["T_mm"] = pd.to_numeric(df[c_t], errors="coerce") if c_t else np.nan
    out["L_mm"] = pd.to_numeric(df[c_l], errors="coerce") if c_l else np.nan
    out["W_mm"] = pd.to_numeric(df[c_w], errors="coerce") if c_w else np.nan
    out["Received_Qty"] = _to_int_qty(df[c_qty]) if c_qty else pd.Series([0]*len(df), dtype="Int64")
    out["Weight_MT"] = _to_mt2(df[c_wtmt]) if c_wtmt else _to_mt2(0)
    out["Receiver"] = "Core_Fab"
    out["record_type"] = "RCV"
    return out.dropna(how="all")

def parse_vrinda(df: pd.DataFrame) -> pd.DataFrame:
    m = _lower_map(df)
    out = pd.DataFrame()

    c_t = _pick(m, "t (mm)", "t")
    c_l = _pick(m, "l (mm)", "l")
    c_w = _pick(m, "w (mm)", "w")
    c_qty = _pick(m, "qty (nos)", "qty")
    c_wtmt = _pick(m, "inv wt (mt)", "inv wt")

    out["T_mm"] = pd.to_numeric(df[c_t], errors="coerce") if c_t else np.nan
    out["L_mm"] = pd.to_numeric(df[c_l], errors="coerce") if c_l else np.nan
    out["W_mm"] = pd.to_numeric(df[c_w], errors="coerce") if c_w else np.nan
    out["Received_Qty"] = _to_int_qty(df[c_qty]) if c_qty else pd.Series([0]*len(df), dtype="Int64")
    out["Weight_MT"] = _to_mt2(df[c_wtmt]) if c_wtmt else _to_mt2(0)
    out["Receiver"] = "Vrinda"
    out["record_type"] = "RCV"
    return out.dropna(how="all")

def parse_excel(file) -> pd.DataFrame:
    xls = pd.ExcelFile(file)
    frames = []
    for i, sheet in enumerate(xls.sheet_names):
        df = xls.parse(sheet)
        if i == 0:
            frames.append(parse_po(df))
        elif i == 1:
            frames.append(parse_core_fab(df))
        elif i == 2:
            frames.append(parse_vrinda(df))
        # ignore extra sheets
    if not frames:
        return pd.DataFrame()
    # Return a combined frame; PO uses PO_* fields; RCV rows use Received fields
    return pd.concat(frames, ignore_index=True)

# ----------------------------
# Sidebar upload
# ----------------------------
with st.sidebar:
    uploads = st.file_uploader(
        "Upload Excel file(s) (Sheet1=PO, Sheet2=Core_Fab, Sheet3=Vrinda)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
    )

# ----------------------------
# Parse all & build datasets
# ----------------------------
all_df = [parse_excel(u) for u in uploads] if uploads else []
data = pd.concat(all_df, ignore_index=True) if all_df else pd.DataFrame()

# ----------------------------
# Supplier Filter (default: All)
# ----------------------------
if data.empty:
    st.info("Upload file(s) to see KPIs, table, and charts.")
    st.stop()

# Build size → supplier mapping from PO; mark duplicates as "Multiple"
po_only = data.loc[data["record_type"] == "PO", ["supplier", "T_mm", "L_mm", "W_mm"]].copy()
po_only["supplier"] = po_only["supplier"].astype(str).str.strip()

dup_sizes = (
    po_only.groupby(["T_mm", "L_mm", "W_mm"])["supplier"]
           .nunique()
           .reset_index(name="n_sup")
)
multi_keys = set(tuple(x) for x in dup_sizes.loc[dup_sizes["n_sup"] > 1, ["T_mm","L_mm","W_mm"]].to_numpy())

size_to_supplier = {}
for _, r in po_only.iterrows():
    key = (r["T_mm"], r["L_mm"], r["W_mm"])
    size_to_supplier[key] = "Multiple" if key in multi_keys else r["supplier"]

supplier_options = sorted([s for s in po_only["supplier"].dropna().unique() if str(s).strip() != ""])
if not supplier_options:
    supplier_options = ["Unknown"]

selected_suppliers = st.multiselect(
    "Filter by Supplier",
    options=supplier_options,
    default=supplier_options,
    help="Pick supplier(s) to filter KPIs, table, and charts. Defaults to all.",
)

# Tag supplier for RCV rows using size map
data = data.copy()
data["supplier_rcv"] = None
rcv_mask = data["record_type"] == "RCV"
if rcv_mask.any():
    keys = list(zip(data.loc[rcv_mask, "T_mm"], data.loc[rcv_mask, "L_mm"], data.loc[rcv_mask, "W_mm"]))
    data.loc[rcv_mask, "supplier_rcv"] = [size_to_supplier.get(k, "Unknown") for k in keys]

# Apply filter to PO and RCV rows, then combine
po_f  = data.loc[(data["record_type"] == "PO")  & (data["supplier"].isin(selected_suppliers))].copy()
rcv_f = data.loc[(data["record_type"] == "RCV") & (data["supplier_rcv"].isin(selected_suppliers))].copy()
data  = pd.concat([po_f, rcv_f], ignore_index=True)

# ----------------------------
# PO vs Dispatch Summary — Matrix Card + Plotly % Dispatch bar
# ----------------------------

st.markdown("---")
st.subheader("📊 PO vs Dispatch Summary")

# --- compute totals on filtered data ---
po_qty = int(_to_int_qty(data.get("PO_Qty", 0)).sum())
rcv_qty = int(_to_int_qty(data.get("Received_Qty", 0)).sum())
bal_qty = po_qty - rcv_qty if po_qty > 0 else 0

po_mt = float(_to_mt2(data.get("PO_Weight_MT", 0)).sum())
rcv_mt = float(_to_mt2(data.get("Weight_MT", 0)).sum())
bal_mt = po_mt - rcv_mt if po_mt > 0 else 0.0

# --- compute DC / Invoice counts (respecting supplier filter) ---
def _find_col(cols, *cands):
    m = {str(c).strip().lower(): c for c in cols}
    for name in cands:
        key = str(name).strip().lower()
        if key in m:
            return m[key]
    return None

dc_set, inv_set = set(), set()
if uploads:
    for upl in uploads:
        try:
            xls = pd.ExcelFile(upl)
            sheets = xls.sheet_names
            for idx in [1, 2]:
                if len(sheets) <= idx:
                    continue
                df = xls.parse(sheets[idx])

                # identify columns
                c_dc  = _find_col(df.columns, "DC NO", "DC NO.", "DC Number", "DC Number.")
                c_inv = _find_col(df.columns, "INVOICE NO", "INV NO", "Invoice No", "Inv No", "INVOICE NO.", "INV NO.")

                # map supplier via size (to respect selected_suppliers)
                c_t = _find_col(df.columns, "T (mm)", "t (mm)", "t")
                c_l = _find_col(df.columns, "L (mm)", "l (mm)", "l")
                c_w = _find_col(df.columns, "W (mm)", "w (mm)", "w")
                if c_t and c_l and c_w:
                    sz = df[[c_t, c_l, c_w]].copy()
                    sz.columns = ["T_mm", "L_mm", "W_mm"]
                    keys = list(zip(pd.to_numeric(sz["T_mm"], errors="coerce"),
                                    pd.to_numeric(sz["L_mm"], errors="coerce"),
                                    pd.to_numeric(sz["W_mm"], errors="coerce")))
                    sup_col = pd.Series([size_to_supplier.get(k, "Unknown") for k in keys], index=df.index)
                else:
                    sup_col = pd.Series(["Unknown"] * len(df), index=df.index)

                mask = sup_col.isin(selected_suppliers)

                if c_dc:
                    vals = (df.loc[mask, c_dc]
                              .astype(str).str.strip()
                              .replace({"nan": "", "None": ""}))
                    dc_set.update(v for v in vals if v)
                if c_inv:
                    vals = (df.loc[mask, c_inv]
                              .astype(str).str.strip()
                              .replace({"nan": "", "None": ""}))
                    inv_set.update(v for v in vals if v)
        except Exception as e:
            st.warning(f"Could not read DC/Invoice counts from {upl.name}: {e}")

dc_count  = len(dc_set)
inv_count = len(inv_set)

# --- four matrix cards in one line (HD style) ---
c1, c2, c3, c4 = st.columns(4)

card_css = """
<style>
.matrix-card {
  border-radius: 14px; padding: 18px;
  background: linear-gradient(135deg, #e0f2fe, #ede9fe); /* 4K HD gradient */
  text-align: center; box-shadow: 0 4px 12px rgba(0,0,0,0.08);
}
.matrix-card .title {
  font-weight: 700; color:#1e293b; font-size:20px;
}
.matrix-card .value {
  font-size: 26px; font-weight:600; color:#111827; margin-top:10px;
}
</style>
"""
st.markdown(card_css, unsafe_allow_html=True)

with c1:
    st.markdown(
        f"""
        <div class="matrix-card">
            <div class="title">Dispatched / PO (Qty)</div>
            <div class="value">{rcv_qty:,} / {po_qty:,}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with c2:
    st.markdown(
        f"""
        <div class="matrix-card">
            <div class="title">Dispatched / PO (MT)</div>
            <div class="value">{rcv_mt:,.2f} / {po_mt:,.2f}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with c3:
    qty_style = "color:#065f46;" if bal_qty == 0 else ("color:#991b1b;" if bal_qty < 0 else "color:#111827;")
    mt_style  = "color:#065f46;" if bal_mt == 0 else ("color:#991b1b;" if bal_mt < 0 else "color:#111827;")
    st.markdown(
        f"""
        <div class="matrix-card">
            <div class="title">Balance Qty / MT</div>
            <div class="value">
                <span style="{qty_style}">{bal_qty:,}</span> /
                <span style="{mt_style}">{bal_mt:,.2f}</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with c4:
    st.markdown(
        f"""
        <div class="matrix-card">
            <div class="title">DC Completed / Invoices (Qty)</div>
            <div class="value">{dc_count:,} / {inv_count:,}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

# --- add vertical gap between cards and graph ---
st.markdown("<div style='margin-top:25px;'></div>", unsafe_allow_html=True)

# --- Plotly % Dispatch bar (by Weight MT) ---
pct_mt = (rcv_mt / po_mt * 100.0) if po_mt > 0 else 0.0
remaining = max(0.0, 100.0 - pct_mt)

fig_pct = go.Figure()
fig_pct.add_trace(go.Bar(x=[pct_mt], y=[""], orientation="h",
                         marker=dict(color="#2e7d32"),
                         hovertemplate="Dispatched: %{x:.1f}%<extra></extra>", showlegend=False))
fig_pct.add_trace(go.Bar(x=[remaining], y=[""], orientation="h",
                         marker=dict(color="#d9d9d9"),
                         hovertemplate="Remaining: %{x:.1f}%<extra></extra>", showlegend=False))
fig_pct.update_layout(
    barmode="stack",
    xaxis=dict(range=[0, 100], ticksuffix="%", showgrid=False, zeroline=False),
    yaxis=dict(showticklabels=False),
    height=120,
    margin=dict(l=10, r=10, t=30, b=10),
    title=f"{pct_mt:.1f}% Dispatched (by Weight MT)",
    plot_bgcolor="white",
)
st.plotly_chart(fig_pct, use_container_width=True)

# ----------------------------
# Build the consolidated table
# Rows are based on PO groups (size + supplier).
# Each receiver gets its own row with Received_Qty and Received_Weight_MT,
# and the Balance_Qty is size-level (PO_Qty − total Received_Qty across receivers).
# ----------------------------
po = data[data["record_type"] == "PO"].copy()
po["PO_Qty"] = _to_int_qty(po["PO_Qty"])
po["PO_Weight_MT"] = _to_mt2(po["PO_Weight_MT"])

po_g = (
    po.groupby(["T_mm", "L_mm", "W_mm", "supplier"], dropna=False)[["PO_Qty", "PO_Weight_MT"]]
      .sum()
      .reset_index()
)

rcv = data[data["record_type"] == "RCV"].copy()
rcv["Received_Qty"] = _to_int_qty(rcv["Received_Qty"])
rcv["Received_Weight_MT"] = _to_mt2(rcv["Weight_MT"])

rcv_g = (
    rcv.groupby(["T_mm", "L_mm", "W_mm", "Receiver"], dropna=False)[["Received_Qty", "Received_Weight_MT"]]
       .sum()
       .reset_index()
)

rcv_sum = (
    rcv.groupby(["T_mm", "L_mm", "W_mm"], dropna=False)["Received_Qty"]
       .sum()
       .reset_index()
       .rename(columns={"Received_Qty": "Total_Received_Qty"})
)
rcv_sum_idx = rcv_sum.set_index(["T_mm", "L_mm", "W_mm"])

rows = []
rcv_idx = rcv_g.set_index(["T_mm", "L_mm", "W_mm"])
for _, r in po_g.iterrows():
    size_key = (r["T_mm"], r["L_mm"], r["W_mm"])
    total_rcv = int(rcv_sum_idx.loc[size_key]["Total_Received_Qty"]) if size_key in rcv_sum_idx.index else 0
    balance_qty = int(max((int(r["PO_Qty"]) if pd.notna(r["PO_Qty"]) else 0) - total_rcv, 0))

    if size_key in rcv_idx.index:
        rcv_rows = rcv_g[
            (rcv_g["T_mm"] == size_key[0]) &
            (rcv_g["L_mm"] == size_key[1]) &
            (rcv_g["W_mm"] == size_key[2])
        ]
        for _, rr in rcv_rows.iterrows():
            rows.append({
                "T_mm": r["T_mm"],
                "L_mm": r["L_mm"],
                "W_mm": r["W_mm"],
                "PO_Qty": int(r["PO_Qty"]) if pd.notna(r["PO_Qty"]) else 0,
                "PO_Weight_MT": float(r["PO_Weight_MT"]) if pd.notna(r["PO_Weight_MT"]) else 0.0,
                "supplier": r.get("supplier", ""),
                "Receiver": rr.get("Receiver", ""),
                "Received_Qty": int(rr.get("Received_Qty", 0)) if pd.notna(rr.get("Received_Qty", 0)) else 0,
                "Received_Weight_MT": float(rr.get("Received_Weight_MT", 0.0)) if pd.notna(rr.get("Received_Weight_MT", 0.0)) else 0.0,
                "Total_Received_Qty": total_rcv,
                "Balance_Qty": balance_qty,
            })
    else:
        rows.append({
            "T_mm": r["T_mm"],
            "L_mm": r["L_mm"],
            "W_mm": r["W_mm"],
            "PO_Qty": int(r["PO_Qty"]) if pd.notna(r["PO_Qty"]) else 0,
            "PO_Weight_MT": float(r["PO_Weight_MT"]) if pd.notna(r["PO_Weight_MT"]) else 0.0,
            "supplier": r.get("supplier", ""),
            "Receiver": "",
            "Received_Qty": 0,
            "Received_Weight_MT": 0.0,
            "Total_Received_Qty": total_rcv,
            "Balance_Qty": balance_qty,
        })

final_tbl = pd.DataFrame(rows)

# Format types
final_tbl["PO_Qty"] = _to_int_qty(final_tbl["PO_Qty"])
final_tbl["Received_Qty"] = _to_int_qty(final_tbl["Received_Qty"])
final_tbl["Total_Received_Qty"] = _to_int_qty(final_tbl["Total_Received_Qty"])
final_tbl["Balance_Qty"] = _to_int_qty(final_tbl["Balance_Qty"])
final_tbl["PO_Weight_MT"] = _to_mt2(final_tbl["PO_Weight_MT"])
final_tbl["Received_Weight_MT"] = _to_mt2(final_tbl["Received_Weight_MT"])

# ----------------------------
# Option 1 visual: show size-level figures only once per (T,L,W,supplier)
# ----------------------------
final_tbl = final_tbl.sort_values(
    ["T_mm", "L_mm", "W_mm", "supplier", "Receiver"], kind="mergesort"
).reset_index(drop=True)

final_tbl["PO_Qty_display"]             = final_tbl["PO_Qty"].astype("Int64")
final_tbl["PO_Weight_MT_display"]       = final_tbl["PO_Weight_MT"]
final_tbl["Total_Received_Qty_display"] = final_tbl["Total_Received_Qty"].astype("Int64")
final_tbl["Balance_Qty_display"]        = final_tbl["Balance_Qty"].astype("Int64")

dupe_mask = final_tbl.duplicated(subset=["T_mm", "L_mm", "W_mm", "supplier"], keep="first")
final_tbl.loc[dupe_mask, [
    "PO_Qty_display",
    "PO_Weight_MT_display",
    "Total_Received_Qty_display",
    "Balance_Qty_display",
]] = [pd.NA, np.nan, pd.NA, pd.NA]

# ----------------------------
# Display table (web) with Balance coloring
# ----------------------------
st.markdown("---")
st.subheader("📑 Consolidated")

view_tbl = final_tbl[
    ["supplier","T_mm", "L_mm", "W_mm",
     "PO_Qty_display", "PO_Weight_MT_display", 
     "Receiver", "Received_Qty", "Received_Weight_MT",
     "Total_Received_Qty_display", "Balance_Qty_display"]
].rename(columns={
    "PO_Qty_display": "PO_Qty",
    "PO_Weight_MT_display": "PO_Weight_MT",
    "Total_Received_Qty_display": "Total_Received_Qty",
    "Balance_Qty_display": "Balance_Qty",
})

def color_balance(val):
    if pd.isna(val): return ""
    if val == 0:     return "background-color: lightgreen; color: black;"
    if val < 0:      return "background-color: salmon; color: white;"
    return ""

st.dataframe(view_tbl.style.applymap(color_balance, subset=["Balance_Qty"]),
             use_container_width=True)

# ----------------------------
# Excel export (pretty + traffic-light rules)
# ----------------------------
import io
with io.BytesIO() as buf:
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        view_tbl.to_excel(writer, index=False, sheet_name="Consolidated")
        wb = writer.book
        ws = writer.sheets["Consolidated"]

        header_fmt = wb.add_format({"bold": True, "bg_color": "#F2F2F2", "border": 1})
        int_fmt    = wb.add_format({"num_format": "0"})
        mt_fmt     = wb.add_format({"num_format": "#,##0.00"})
        text_fmt   = wb.add_format({})
        green_fmt  = wb.add_format({"bg_color": "#C6EFCE", "font_color": "#006100"})
        red_fmt    = wb.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})

        for col_idx, col_name in enumerate(view_tbl.columns):
            ws.write(0, col_idx, col_name, header_fmt)

        cols = list(view_tbl.columns)
        idx = {name: cols.index(name) for name in cols}

        ws.set_column(idx["T_mm"],                   idx["T_mm"],                   8,  int_fmt)
        ws.set_column(idx["L_mm"],                   idx["L_mm"],                   12, int_fmt)
        ws.set_column(idx["W_mm"],                   idx["W_mm"],                   12, int_fmt)
        ws.set_column(idx["PO_Qty"],                 idx["PO_Qty"],                 12, int_fmt)
        ws.set_column(idx["PO_Weight_MT"],           idx["PO_Weight_MT"],           16, mt_fmt)
        ws.set_column(idx["supplier"],               idx["supplier"],               18, text_fmt)
        ws.set_column(idx["Receiver"],               idx["Receiver"],               14, text_fmt)
        ws.set_column(idx["Received_Qty"],           idx["Received_Qty"],           16, int_fmt)
        ws.set_column(idx["Received_Weight_MT"],     idx["Received_Weight_MT"],     20, mt_fmt)
        ws.set_column(idx["Total_Received_Qty"],     idx["Total_Received_Qty"],     20, int_fmt)
        ws.set_column(idx["Balance_Qty"],            idx["Balance_Qty"],            14, int_fmt)

        last_row = len(view_tbl)
        last_col = len(cols) - 1
        ws.freeze_panes(1, 0)
        ws.autofilter(0, 0, last_row, last_col)

        bal_col = idx["Balance_Qty"]
        ws.conditional_format(1, bal_col, last_row, bal_col,
                              {"type": "cell", "criteria": "==", "value": 0, "format": green_fmt})
        ws.conditional_format(1, bal_col, last_row, bal_col,
                              {"type": "cell", "criteria": "<", "value": 0, "format": red_fmt})

    excel_bytes = buf.getvalue()

st.download_button(
    label="⬇️ Download Excel",
    data=excel_bytes,
    file_name="FIM_consolidated.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

# ============================
# Advanced Chart: Cumulative Dispatch Weight vs Date (Combined)
# with Supplier-labeled Delivery Date markers
# ============================
st.markdown("---")
st.subheader("📈 Cumulative Dispatch Weight vs Date")

def _find_col(cols, *cands):
    m = {str(c).strip().lower(): c for c in cols}
    for name in cands:
        key = str(name).strip().lower()
        if key in m:
            return m[key]
    return None

dispatch_rows = []
supplier_delivery_dates = {}

if uploads:
    for upl in uploads:
        try:
            xls = pd.ExcelFile(upl)
            sheets = xls.sheet_names

            # --- PO sheet (0) → Supplier + Delivery Dates (respect filter) ---
            if len(sheets) >= 1:
                po_df = xls.parse(sheets[0])
                c_sup = _find_col(po_df.columns, "SUPPLIER", "Supplier")
                c_del = _find_col(po_df.columns, "Delivery Date", "DELIVERY DATE", "delivery_date", "Delv Date", "Delv_Date")
                if c_sup and c_del:
                    tmp = po_df[[c_sup, c_del]].copy()
                    tmp.columns = ["supplier", "deliv"]
                    tmp["supplier"] = tmp["supplier"].astype(str).str.strip()
                    tmp = tmp[tmp["supplier"].isin(selected_suppliers)]  # filter
                    tmp["deliv"] = pd.to_datetime(tmp["deliv"], errors="coerce").dt.normalize()
                    tmp = tmp.dropna()
                    for sup, g in tmp.groupby("supplier"):
                        supplier_delivery_dates.setdefault(sup, set()).update(set(g["deliv"].unique()))

            # --- Receiver sheets (1 & 2) → INV DATE + INV WT (MT), map to supplier & filter ---
            for idx in [1, 2]:
                if len(sheets) <= idx:
                    continue
                rcv_df = xls.parse(sheets[idx])

                c_date = _find_col(rcv_df.columns, "INV DATE", "INVOICE DATE", "Inv Date", "Invoice Date", "inv_date", "invoice_date")
                c_wt   = _find_col(rcv_df.columns, "INV WT (MT)", "INV WT", "Weight (MT)", "Net wt of Inv (MT)", "NET wt of Inv (MT)")
                c_t    = _find_col(rcv_df.columns, "T (mm)", "t (mm)", "t")
                c_l    = _find_col(rcv_df.columns, "L (mm)", "l (mm)", "l")
                c_w    = _find_col(rcv_df.columns, "W (mm)", "w (mm)", "w")
                if not (c_date and c_wt):
                    continue

                # Map each row to supplier via size; Unknown if not matched
                if all([c_t, c_l, c_w]):
                    sz = rcv_df[[c_t, c_l, c_w]].copy()
                    sz.columns = ["T_mm", "L_mm", "W_mm"]
                    keys = list(zip(pd.to_numeric(sz["T_mm"], errors="coerce"),
                                    pd.to_numeric(sz["L_mm"], errors="coerce"),
                                    pd.to_numeric(sz["W_mm"], errors="coerce")))
                    sup_col = [size_to_supplier.get(k, "Unknown") for k in keys]
                else:
                    sup_col = ["Unknown"] * len(rcv_df)

                tmp = rcv_df[[c_date, c_wt]].copy()
                tmp.columns = ["date", "wt_mt"]
                tmp["date"] = pd.to_datetime(tmp["date"], errors="coerce").dt.normalize()
                tmp["wt_mt"] = pd.to_numeric(tmp["wt_mt"], errors="coerce")
                tmp["supplier"] = sup_col
                tmp = tmp[(~tmp["date"].isna()) & (~tmp["wt_mt"].isna())]
                tmp = tmp[tmp["supplier"].isin(selected_suppliers)]  # filter
                if not tmp.empty:
                    g = tmp.groupby("date", as_index=False)["wt_mt"].sum()
                    dispatch_rows.append(g)
        except Exception as e:
            st.warning(f"Could not read data from {upl.name}: {e}")

# Build daily + cumulative (combined, filtered)
if dispatch_rows:
    daily = (
        pd.concat(dispatch_rows, ignore_index=True)
          .groupby("date", as_index=False)["wt_mt"].sum()
          .sort_values("date")
          .reset_index(drop=True)
    )
else:
    daily = pd.DataFrame(columns=["date", "wt_mt"])

if daily.empty and not supplier_delivery_dates:
    st.info("No dispatch dates or delivery-date data found in the uploaded files for the selected supplier(s).")
else:
    daily["cum_wt_mt"] = daily["wt_mt"].cumsum()

    fig = go.Figure()
    if not daily.empty:
        fig.add_trace(go.Scatter(
            x=daily["date"], y=daily["cum_wt_mt"],
            mode="lines+markers+text",
            name="Cumulative Dispatched (MT)",
            line=dict(color="#2e7d32", width=3),
            marker=dict(size=8, color="#2e7d32"),
            text=[f"{v:.2f}" for v in daily["cum_wt_mt"]],
            textposition="top center",
            hovertemplate="Date: %{x|%Y-%m-%d}<br>Cumulative: %{y:.2f} MT<extra></extra>",
        ))

    # --- Dashed red Delivery Date lines + supplier annotations (staggered & always visible) ---
    all_delivery_dates = []
    for sup, dates in supplier_delivery_dates.items():
     dd_sorted = sorted(pd.to_datetime(list(dates)).unique())
     all_delivery_dates.extend(dd_sorted)
     for i, d in enumerate(dd_sorted):
        fig.add_vline(x=d, line_width=2, line_dash="dash", line_color="red")

        # Two rows of labels at the very top of the plotting area
        yshift = -10 if (i % 2 == 0) else 6
        fig.add_annotation(
            x=d, y=1, xref="x", yref="paper",
            text=str(sup),
            showarrow=False, yanchor="bottom",
            font=dict(color="black", size=14, family="Arial"),
            textangle=-90,
            bgcolor="rgba(255,255,255,0.6)",
            bordercolor="red", borderwidth=0,
            yshift=yshift  # <-- push label above the plot so it isn't clipped
        )


    # --- X-axis ticks: every dispatch date; include delivery-only dates and pad range ---
    if not daily.empty:
        dispatch_dates = pd.to_datetime(daily["date"]).sort_values().unique()
    else:
        dispatch_dates = np.array([], dtype="datetime64[ns]")

    all_tick_dates = sorted(pd.to_datetime(np.unique(np.concatenate([
        dispatch_dates,
        np.array(all_delivery_dates, dtype="datetime64[ns]") if len(all_delivery_dates) else np.array([], dtype="datetime64[ns]")
    ]))))

    if all_tick_dates:
        x_min = (pd.Timestamp(all_tick_dates[0]) - pd.Timedelta(days=1))
        x_max = (pd.Timestamp(all_tick_dates[-1]) + pd.Timedelta(days=1))
        fig.update_xaxes(
            range=[x_min, x_max],
            tickmode="array",
            tickvals=all_tick_dates,
            tickformat="%d-%b",
            tickangle=-45,
            showgrid=True,
            gridcolor="#eaeaea"
        )

    fig.update_layout(
        xaxis_title="Date",
        yaxis_title="Cumulative Dispatch (MT)",
        margin=dict(l=10, r=10, t=30, b=10),
        plot_bgcolor="white",
        hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0)
    )
    fig.update_yaxes(showgrid=True, gridcolor="#f0f0f0")

    st.plotly_chart(fig, use_container_width=True)
