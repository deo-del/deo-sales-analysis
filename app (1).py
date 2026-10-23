"""
Sales Order vs Invoice – Analytics Dashboard
============================================
Reads three Google Sheets (Sale order, Sale Invoice, Client_master), links them on
SODOCNO + ITEM NAME and Party Name, and shows month-wise order / invoice / shortfall
reports with a Party Group -> Party -> Order-line drill-down.

Month is always taken from SODATE (the sales-order date), as requested.
Run locally:  streamlit run app.py
"""
import io
import re
from datetime import datetime

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Order vs Invoice Report", page_icon="📊", layout="wide")

COLORS = {"order": "#2F4B7C", "invoice": "#3A9E8F", "short": "#D1495B", "neutral": "#8C8C8C"}


# =============================================================================
# Small helpers
# =============================================================================
def secret(key, default=None):
    """Read st.secrets safely (works even when no secrets file exists)."""
    try:
        return st.secrets[key] if key in st.secrets else default
    except Exception:
        return default


def norm_col(c):
    return re.sub(r"[^A-Z0-9]", "", str(c).upper())


def make_unique(cols):
    seen, out = {}, []
    for c in cols:
        c = str(c).strip() or "BLANK"
        if c in seen:
            seen[c] += 1
            out.append(f"{c}_{seen[c]}")
        else:
            seen[c] = 0
            out.append(c)
    return out


def find_col(df, *candidates, required=True, prefix=False):
    """Find a column by name, ignoring case/spaces/punctuation (e.g. 'BILL NO.' == 'billno')."""
    lookup = {norm_col(c): c for c in df.columns}
    for cand in candidates:
        if norm_col(cand) in lookup:
            return lookup[norm_col(cand)]
    if prefix:
        for cand in candidates:
            for k, orig in lookup.items():
                if k.startswith(norm_col(cand)):
                    return orig
    if required:
        raise ValueError(f"Column {candidates[0]!r} not found. Columns available: {list(df.columns)}")
    return None


def norm_text(s):
    return s.fillna("").astype(str).str.upper().str.replace(r"\s+", " ", regex=True).str.strip()


def norm_key(s):
    """Join key: upper-case, letters & digits only, so 'New  Bhiwandi )' == 'NEW BHIWANDI'."""
    return s.fillna("").astype(str).str.upper().str.replace(r"[^A-Z0-9]", "", regex=True)


def to_id(s):
    return s.fillna("").astype(str).str.strip().str.replace(r"\.0$", "", regex=True)


def to_num(s):
    s = s.fillna("").astype(str).str.replace(",", "", regex=False)
    s = s.str.extract(r"(-?\d+(?:\.\d+)?)", expand=False)  # handles '90 KGS', '₹1,200'
    return pd.to_numeric(s, errors="coerce").fillna(0.0)


DATE_FORMATS = ["%Y%m%d", "%d.%m.%Y", "%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d", "%m/%d/%Y %H:%M:%S"]


def to_date(s, formats=DATE_FORMATS):
    s = s.fillna("").astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    out = pd.Series(pd.NaT, index=s.index, dtype="datetime64[ns]")
    for f in formats:
        miss = out.isna()
        if not miss.any():
            break
        out[miss] = pd.to_datetime(s[miss], format=f, errors="coerce")
    return out


def drop_blank_rows(df):
    return df.replace(r"^\s*$", np.nan, regex=True).dropna(how="all").fillna("")


def inr(x):
    """Indian number grouping: 1234567 -> ₹12,34,567"""
    if pd.isna(x):
        return "–"
    neg, s = x < 0, str(int(round(abs(x))))
    if len(s) > 3:
        head = re.sub(r"(\d)(?=(\d{2})+$)", r"\1,", s[:-3])
        s = f"{head},{s[-3:]}"
    return f"{'-' if neg else ''}₹{s}"


def short_inr(x):
    if abs(x) >= 1e7:
        return f"₹{x / 1e7:.2f} Cr"
    if abs(x) >= 1e5:
        return f"₹{x / 1e5:.2f} L"
    return inr(x)


def pct(a, b):
    return np.where(np.asarray(b) > 0, np.asarray(a) / np.where(np.asarray(b) > 0, b, 1) * 100, np.nan)


# =============================================================================
# Loading
# =============================================================================
@st.cache_data(ttl=600, show_spinner=False)
def load_google_sheet(sheet_id, tab, gid, use_service_account):
    if use_service_account:
        import gspread

        gc = gspread.service_account_from_dict(dict(st.secrets["gcp_service_account"]))
        sh = gc.open_by_key(sheet_id)
        ws = sh.worksheet(tab) if tab else sh.get_worksheet(0)
        values = ws.get_all_values()
        if not values:
            return pd.DataFrame()
        return pd.DataFrame(values[1:], columns=make_unique(values[0]))
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
    df = pd.read_csv(url, dtype=str, keep_default_na=False)
    df.columns = make_unique(df.columns)
    return df


@st.cache_data(show_spinner=False)
def load_sample():
    read = lambda p: pd.read_csv(p, dtype=str, keep_default_na=False)
    return (read("sample_data/sale_order.csv"), read("sample_data/sale_invoice.csv"),
            read("sample_data/client_master.csv"))


def read_upload(f):
    if f.name.lower().endswith((".xlsx", ".xls")):
        df = pd.read_excel(f, dtype=str).fillna("")
    else:
        df = pd.read_csv(f, dtype=str, keep_default_na=False)
    df.columns = make_unique(df.columns)
    return df


# =============================================================================
# Cleaning – one function per sheet
# =============================================================================
def dedupe(df):
    cols = [c for c in df.columns if norm_col(c) not in ("TIMESTAMP", "ID")]
    dup = df.duplicated(subset=cols)
    return df[~dup], int(dup.sum())


def prep_orders(raw, remove_dups):
    df = drop_blank_rows(raw)
    n_dup = 0
    if remove_dups:
        df, n_dup = dedupe(df)
    col = lambda *c, **k: find_col(df, *c, **k)
    o = pd.DataFrame(index=df.index)
    o["SODOCNO"] = to_id(df[col("SODOCNO", "SO DOC NO", "SO NO")])
    o["SODATE"] = to_date(df[col("SODATE", "SO DATE")])
    o["PARTY_NAME"] = norm_text(df[col("PARTY NAME", "PARTY")])
    c = col("ORDNO", "ORDER NO", required=False)
    o["ORDNO"] = norm_text(df[c]) if c else ""
    c = col("ORDDATE", "ORDER DATE", required=False)
    o["ORDDATE"] = to_date(df[c]) if c else pd.NaT
    o["ITEM_NAME"] = norm_text(df[col("ITEM NAME", "ITEM")])
    o["ORDER_QTY"] = to_num(df[col("QTY", "ORDER QTY")])
    o["ORDER_RATE"] = to_num(df[col("RATE")])
    c = col("ITEM MRP", "MRP", required=False)
    o["MRP"] = to_num(df[c]) if c else 0.0
    o["ORDER_AMT"] = to_num(df[col("AMOUNT", "AMT")])
    o["ORDER_AMT"] = np.where(o["ORDER_AMT"].eq(0), o["ORDER_QTY"] * o["ORDER_RATE"], o["ORDER_AMT"])
    c = col("GSTNO", "GST NO", "GSTIN", required=False)
    o["GSTNO"] = norm_key(df[c]) if c else ""
    c = col("PARTIAL STATUS", required=False)
    o["PARTIAL_STATUS"] = norm_text(df[c]) if c else ""
    o["DISCONTINUED"] = o["PARTIAL_STATUS"].str.contains("DISCONTINU", na=False)
    o["MRP_VALUE"] = o["ORDER_QTY"] * o["MRP"]
    o = o[o["SODOCNO"].ne("")]
    return o, n_dup


def prep_invoices(raw, remove_dups):
    df = drop_blank_rows(raw)
    n_dup = 0
    if remove_dups:
        df, n_dup = dedupe(df)
    col = lambda *c, **k: find_col(df, *c, **k)
    i = pd.DataFrame(index=df.index)
    i["BILL_NO"] = to_id(df[col("BILL NO", "INVOICE NO", "BILLNO")])
    i["BILL_DATE"] = to_date(df[col("BILL DATE", "INVOICE DATE")])
    i["INV_PARTY_NAME"] = norm_text(df[col("PARTY NAME", "PARTY")])
    i["ITEM_NAME"] = norm_text(df[col("ITEM NAME", "ITEM")])
    i["SODOCNO"] = to_id(df[col("SODOCNO", "SO DOC NO", "SO NO")])
    optional = {"TAXABLE": ("TAXABLE",), "TRANSPORT": ("TRANSPORT",), "DELIVERY": ("DELIVERY",),
                "GROSS_WT": ("GROSS WT", "GROSS WEIGHT"), "PACKAGE": ("PACKAGE", "PACKAGES", "PKG"),
                "DISC_PCT": ("DISC%", "DISC PCT"), "DISC_AMT": ("DISC. AMOUNT", "DISCOUNT AMOUNT")}
    text_cols = {"TRANSPORT", "DELIVERY"}
    for name, cands in optional.items():
        c = col(*cands, required=False, prefix=(name == "TAXABLE"))
        if name in text_cols:
            i[name] = norm_text(df[c]) if c else "NOT SET"
        else:
            i[name] = to_num(df[c]) if c else 0.0
    i["INV_QTY"] = to_num(df[col("QTY", "INVOICE QTY")])
    i["INV_RATE"] = to_num(df[col("RATE")])
    i["INV_AMT"] = to_num(df[col("AMOUNT", "AMT")])
    i["INV_AMT"] = np.where(i["INV_AMT"].eq(0), i["INV_QTY"] * i["INV_RATE"] - i["DISC_AMT"], i["INV_AMT"])
    i = i[i["BILL_NO"].ne("") | i["SODOCNO"].ne("")]
    return i, n_dup


CLIENT_ATTRS = {
    "CLIENT_CODE": ("CLIENT CODE",), "PARTY_GROUP": ("PARTY GROUP", "GROUP"), "STATUS": ("STATUS",),
    "TYPE": ("TYPE",), "STATE": ("STATE",), "DISTRICT": ("DISTRICT",), "CITY": ("CITY",),
    "ZONE": ("ZONE",), "SALESMAN": ("SALESMAN", "SALES MAN"), "PAYMENT_TERM": ("PAYMENT TERM",),
    "TRANSPORT_NAME": ("TRANSPORT NAME",), "PLATFORM": ("PLATEFORM", "PLATFORM"), "CL_GSTNO": ("GSTNO", "GST NO", "GSTIN"),
}


def prep_clients(raw):
    df = drop_blank_rows(raw)
    c = pd.DataFrame(index=df.index)
    c["CL_PARTY_NAME"] = norm_text(df[find_col(df, "PARTY NAME", "PARTY")])
    c["PARTY_KEY"] = norm_key(c["CL_PARTY_NAME"])
    for name, cands in CLIENT_ATTRS.items():
        col = find_col(df, *cands, required=False)
        c[name] = norm_text(df[col]) if col else ""
    c["CL_GSTNO"] = norm_key(c["CL_GSTNO"])
    c = c[c["PARTY_KEY"].ne("")].drop_duplicates("PARTY_KEY")
    return c


# =============================================================================
# Data model
# =============================================================================
@st.cache_data(show_spinner="Building report…")
def build_model(order_raw, inv_raw, client_raw, remove_dups):
    orders, dup_o = prep_orders(order_raw, remove_dups)
    inv, dup_i = prep_invoices(inv_raw, remove_dups)
    client = prep_clients(client_raw)

    orders["PARTY_KEY"] = norm_key(orders["PARTY_NAME"])
    orders["ITEM_KEY"] = norm_key(orders["ITEM_NAME"])
    inv["ITEM_KEY"] = norm_key(inv["ITEM_NAME"])

    # ---- SO header (one row per SODOCNO) -----------------------------------
    so = orders.groupby("SODOCNO").agg(
        SODATE=("SODATE", "min"), PARTY_NAME=("PARTY_NAME", "first"), PARTY_KEY=("PARTY_KEY", "first"),
        GSTNO=("GSTNO", "first"), ORDNO=("ORDNO", "first"), ORDDATE=("ORDDATE", "min"),
        N_PARTIES=("PARTY_KEY", "nunique")).reset_index()

    # map party -> client master (by name, then GST no. as fallback)
    key = so["PARTY_KEY"].where(so["PARTY_KEY"].isin(client["PARTY_KEY"]))
    gst = client[client["CL_GSTNO"].ne("")].drop_duplicates("CL_GSTNO").set_index("CL_GSTNO")["PARTY_KEY"]
    if len(gst):
        key = key.fillna(so["GSTNO"].map(gst))
    so["CLIENT_KEY"] = key
    so = so.merge(client.rename(columns={"PARTY_KEY": "CLIENT_KEY"}), on="CLIENT_KEY", how="left")
    so["MAPPED"] = so["CLIENT_KEY"].notna()
    for name in CLIENT_ATTRS:
        so[name] = so[name].replace("", np.nan).fillna("NOT SET")
    so["PARTY_GROUP"] = np.where(so["MAPPED"], so["PARTY_GROUP"].replace("NOT SET", np.nan).fillna(so["PARTY_NAME"]),
                                 "UNMAPPED PARTY")
    so["MONTH"] = so["SODATE"].dt.strftime("%Y-%m").fillna("NO DATE")
    so["MONTH_NAME"] = so["SODATE"].dt.strftime("%b-%Y").fillna("No date")

    # ---- item-level lines --------------------------------------------------
    ol = orders.groupby(["SODOCNO", "ITEM_KEY"]).agg(
        ITEM_NAME=("ITEM_NAME", "first"), ORDER_QTY=("ORDER_QTY", "sum"), ORDER_AMT=("ORDER_AMT", "sum"),
        MRP=("MRP", "max"), MRP_VALUE=("MRP_VALUE", "sum"), DISCONTINUED=("DISCONTINUED", "max"),
        PARTIAL_STATUS=("PARTIAL_STATUS", "max")).reset_index()

    known = set(so["SODOCNO"])
    orphan_inv = inv[~inv["SODOCNO"].isin(known)]
    inv_linked = inv[inv["SODOCNO"].isin(known)]
    il = inv_linked.groupby(["SODOCNO", "ITEM_KEY"]).agg(
        INV_ITEM_NAME=("ITEM_NAME", "first"), INV_QTY=("INV_QTY", "sum"), INV_AMT=("INV_AMT", "sum"),
        BILLS=("BILL_NO", "nunique"), FIRST_BILL=("BILL_DATE", "min"), LAST_BILL=("BILL_DATE", "max")).reset_index()

    fact = ol.merge(il, on=["SODOCNO", "ITEM_KEY"], how="outer")
    fact["ITEM_NAME"] = fact["ITEM_NAME"].fillna(fact["INV_ITEM_NAME"])
    for c in ["ORDER_QTY", "ORDER_AMT", "MRP", "MRP_VALUE", "INV_QTY", "INV_AMT", "BILLS"]:
        fact[c] = fact[c].fillna(0)
    fact["DISCONTINUED"] = fact["DISCONTINUED"].fillna(False).astype(bool)
    fact["PARTIAL_STATUS"] = fact["PARTIAL_STATUS"].fillna("")
    fact = fact.merge(so.drop(columns=["PARTY_KEY", "GSTNO", "CLIENT_KEY", "CL_PARTY_NAME", "CL_GSTNO"]),
                      on="SODOCNO", how="left")

    fact["SHORT_QTY"] = (fact["ORDER_QTY"] - fact["INV_QTY"]).clip(lower=0)
    avg_rate = np.where(fact["ORDER_QTY"] > 0, fact["ORDER_AMT"] / fact["ORDER_QTY"].where(fact["ORDER_QTY"] > 0, 1), 0)
    fact["ORDER_AVG_RATE"] = avg_rate
    fact["SHORT_AMT"] = fact["SHORT_QTY"] * avg_rate
    fact["EXCESS_QTY"] = (fact["INV_QTY"] - fact["ORDER_QTY"]).clip(lower=0)
    fact["SHORT_AMT_DISC"] = np.where(fact["DISCONTINUED"], fact["SHORT_AMT"], 0)
    fact["SHORT_AMT_PENDING"] = fact["SHORT_AMT"] - fact["SHORT_AMT_DISC"]
    fact["LINE_STATUS"] = np.select(
        [(fact["ORDER_QTY"] == 0) & (fact["INV_QTY"] > 0), fact["INV_QTY"] == 0, fact["SHORT_QTY"] > 0],
        ["Extra (not in order)", "Not invoiced", "Partially invoiced"], "Fully invoiced")

    # invoices enriched with SO attributes (for dispatch / discount reports)
    inv_linked = inv_linked.merge(so[["SODOCNO", "PARTY_NAME", "PARTY_GROUP", "MONTH", "MONTH_NAME", "ZONE", "STATE",
                                      "SALESMAN", "PLATFORM", "TYPE", "CITY"]], on="SODOCNO", how="left")

    dq = {"dup_orders": dup_o, "dup_invoices": dup_i, "orphan_inv": orphan_inv,
          "bad_sodate": int(orders["SODATE"].isna().sum()), "bad_billdate": int(inv["BILL_DATE"].isna().sum()),
          "multi_party_so": so[so["N_PARTIES"] > 1][["SODOCNO", "PARTY_NAME", "N_PARTIES"]],
          "client_rows": len(client)}
    return fact, inv_linked, so, dq


# =============================================================================
# Report helpers
# =============================================================================
METRICS = {
    "Order Qty": ("ORDER_QTY", "sum"), "Order Amount": ("ORDER_AMT", "sum"),
    "Invoice Qty": ("INV_QTY", "sum"), "Invoice Amount": ("INV_AMT", "sum"),
    "Shortfall Qty": ("SHORT_QTY", "sum"), "Shortfall Amount": ("SHORT_AMT", "sum"),
}
LABELS = {"PARTY_GROUP": "Party Group", "PARTY_NAME": "Party", "MONTH_NAME": "Month", "ITEM_NAME": "Item",
          "ZONE": "Zone", "STATE": "State", "CITY": "City", "SALESMAN": "Salesman", "PLATFORM": "Platform",
          "TYPE": "Type", "PAYMENT_TERM": "Payment Term", "CLIENT_CODE": "Client Code", "SODOCNO": "SO No",
          "LINE_STATUS": "Line Status", "PARTIAL_STATUS": "Partial Status", "TRANSPORT": "Transport",
          "DELIVERY": "Delivery Location", "BILL_NO": "Bill No", "DISTRICT": "District"}


def summarize(df, by, counts=True):
    agg = dict(METRICS)
    if counts:
        agg = {"SOs": ("SODOCNO", "nunique"), **agg}
    g = df.groupby(by, dropna=False).agg(**agg).reset_index()
    g["Qty Fill %"] = pct(g["Invoice Qty"], g["Order Qty"])
    g["Value Fill %"] = pct(g["Invoice Amount"], g["Order Amount"])
    if "MONTH" in by:
        g = g.sort_values("MONTH")
    return g


def present(df):
    """Rename internal columns to friendly names and drop helper columns."""
    return df.drop(columns=[c for c in ["MONTH"] if c in df.columns]).rename(columns=LABELS)


def fill_color(v):
    if pd.isna(v):
        return ""
    if v < 80:
        return f"color: {COLORS['short']}; font-weight: 600"
    if v < 95:
        return "color: #B7791F; font-weight: 600"
    return f"color: {COLORS['invoice']}; font-weight: 600"


def styled(df):
    fmt = {}
    for c in df.columns:
        if not pd.api.types.is_numeric_dtype(df[c]) or pd.api.types.is_bool_dtype(df[c]):
            continue
        lc = c.lower()
        if "%" in c:
            fmt[c] = "{:.1f}%"
        elif "rate" in lc or "mrp" in lc or "per kg" in lc:
            fmt[c] = "{:,.2f}"
        elif "amount" in lc or "value" in lc:
            fmt[c] = inr
        else:
            fmt[c] = "{:,.0f}"
    sty = df.style.format(fmt, na_rep="–")
    fill_cols = [c for c in df.columns if "Fill %" in c]
    if fill_cols:
        sty = sty.map(fill_color, subset=fill_cols)
    return sty


def show_table(df, key, selectable=False, height="auto"):
    data = styled(df) if df.size <= 200_000 else df
    if selectable:
        ev = st.dataframe(data, hide_index=True, key=key, height=height,
                          on_select="rerun", selection_mode="single-row")
        rows = ev.selection.rows
        return df.iloc[rows[0]] if rows else None
    st.dataframe(data, hide_index=True, key=key, height=height)
    return None


def order_vs_invoice_chart(df, x, title, horizontal=False):
    fig = go.Figure()
    for col, name, color in [("Order Amount", "Order", COLORS["order"]),
                             ("Invoice Amount", "Invoice", COLORS["invoice"]),
                             ("Shortfall Amount", "Shortfall", COLORS["short"])]:
        if horizontal:
            fig.add_bar(y=df[x], x=df[col], name=name, marker_color=color, orientation="h")
        else:
            fig.add_bar(x=df[x], y=df[col], name=name, marker_color=color)
    fig.update_layout(barmode="group", title=title, height=420, legend=dict(orientation="h", y=1.1),
                      margin=dict(l=10, r=10, t=60, b=10))
    if horizontal:
        fig.update_yaxes(autorange="reversed")
    return fig


def kpi_row(df):
    oa, ia, sa = df["ORDER_AMT"].sum(), df["INV_AMT"].sum(), df["SHORT_AMT"].sum()
    oq, iq, sq = df["ORDER_QTY"].sum(), df["INV_QTY"].sum(), df["SHORT_QTY"].sum()
    c = st.columns(4)
    c[0].metric("Order amount", short_inr(oa), f"{df['SODOCNO'].nunique():,} sales orders", delta_color="off")
    c[1].metric("Invoice amount", short_inr(ia), f"{(ia / oa * 100 if oa else 0):.1f}% of order value", delta_color="off")
    c[2].metric("Shortfall amount", short_inr(sa), f"{(sa / oa * 100 if oa else 0):.1f}% of order value", delta_color="off")
    c[3].metric("Qty fill rate", f"{(iq / oq * 100 if oq else 0):.1f}%", f"{iq:,.0f} of {oq:,.0f} units", delta_color="off")
    c = st.columns(4)
    c[0].metric("Order qty", f"{oq:,.0f}")
    c[1].metric("Invoice qty", f"{iq:,.0f}")
    c[2].metric("Shortfall qty", f"{sq:,.0f}")
    c[3].metric("Parties ordering", f"{df['PARTY_NAME'].nunique():,}", f"{df['ITEM_NAME'].nunique():,} items",
                delta_color="off")


def to_excel(sheets):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as xw:
        for name, d in sheets.items():
            d.to_excel(xw, sheet_name=name[:31], index=False)
            ws = xw.sheets[name[:31]]
            for i, c in enumerate(d.columns):
                ws.set_column(i, i, min(max(len(str(c)), 10) + 2, 45))
            ws.freeze_panes(1, 0)
    return buf.getvalue()


# =============================================================================
# Login (optional – set app_password in secrets)
# =============================================================================
def check_password():
    pwd = secret("app_password")
    if not pwd or st.session_state.get("auth_ok"):
        return True
    st.title("📊 Order vs Invoice Report")
    with st.form("login"):
        p = st.text_input("Password", type="password")
        if st.form_submit_button("Open report"):
            if p == pwd:
                st.session_state["auth_ok"] = True
                st.rerun()
            st.error("That password is not correct.")
    return False


if not check_password():
    st.stop()

# =============================================================================
# Sidebar – data source & filters
# =============================================================================
sheets_cfg = secret("sheets")
with st.sidebar:
    st.header("Data")
    sources = (["Google Sheets"] if sheets_cfg else []) + ["Sample data", "Upload files"]
    source = st.radio("Source", sources, label_visibility="collapsed")
    remove_dups = st.toggle("Ignore duplicate imported rows", value=True,
                            help="Rows identical in every column except Timestamp/id are counted once.")
    if st.button("🔄 Refresh data", width="stretch"):
        st.cache_data.clear()
        st.rerun()

try:
    if source == "Google Sheets":
        use_sa = secret("gcp_service_account") is not None
        with st.spinner("Reading Google Sheets…"):
            order_raw = load_google_sheet(sheets_cfg["order_id"], sheets_cfg.get("order_tab"), sheets_cfg.get("order_gid", 0), use_sa)
            inv_raw = load_google_sheet(sheets_cfg["invoice_id"], sheets_cfg.get("invoice_tab"), sheets_cfg.get("invoice_gid", 0), use_sa)
            client_raw = load_google_sheet(sheets_cfg["client_id"], sheets_cfg.get("client_tab"), sheets_cfg.get("client_gid", 0), use_sa)
    elif source == "Sample data":
        order_raw, inv_raw, client_raw = load_sample()
    else:
        with st.sidebar:
            fo = st.file_uploader("Sale order (CSV/XLSX)", type=["csv", "xlsx"])
            fi = st.file_uploader("Sale invoice (CSV/XLSX)", type=["csv", "xlsx"])
            fc = st.file_uploader("Client master (CSV/XLSX)", type=["csv", "xlsx"])
        if not (fo and fi and fc):
            st.info("Upload the Sale order, Sale invoice and Client master files in the sidebar to build the report.")
            st.stop()
        order_raw, inv_raw, client_raw = read_upload(fo), read_upload(fi), read_upload(fc)

    fact, inv_all, so_all, dq = build_model(order_raw, inv_raw, client_raw, remove_dups)
except ValueError as e:
    st.error(f"A sheet is missing a column this report needs. {e}")
    st.stop()
except Exception as e:  # network / permission problems
    st.error("The data could not be loaded. If you use Google Sheets without a service account, "
             "each sheet must be shared as 'Anyone with the link – Viewer'.")
    st.exception(e)
    st.stop()

if fact.empty:
    st.warning("No sales-order rows were found.")
    st.stop()

with st.sidebar:
    st.header("Filters")
    st.caption("Leave a filter empty to include everything.")
    month_map = fact.drop_duplicates("MONTH").sort_values("MONTH").set_index("MONTH")["MONTH_NAME"]
    f_month = st.multiselect("Month (from SO date)", list(month_map.index), format_func=lambda m: month_map[m])
    filters = {"MONTH": f_month}
    for col, label in [("PARTY_GROUP", "Party group"), ("PARTY_NAME", "Party"), ("ZONE", "Zone"), ("STATE", "State"),
                       ("SALESMAN", "Salesman"), ("PLATFORM", "Platform"), ("TYPE", "Type")]:
        filters[col] = st.multiselect(label, sorted(fact[col].dropna().unique()))
    excl_disc = st.toggle("Exclude discontinued items", value=False,
                          help="Removes lines marked 'Discontinue Item' in Partial Status from every report.")

mask = pd.Series(True, index=fact.index)
for col, vals in filters.items():
    if vals:
        mask &= fact[col].isin(vals)
if excl_disc:
    mask &= ~fact["DISCONTINUED"]
f = fact[mask].copy()
inv_f = inv_all[inv_all["SODOCNO"].isin(f["SODOCNO"])]

if f.empty:
    st.warning("No rows match these filters. Clear one or more filters in the sidebar.")
    st.stop()

# =============================================================================
# Header
# =============================================================================
st.title("📊 Sales order vs invoice")
first, last = f["SODATE"].min(), f["SODATE"].max()
st.caption(f"Sales orders dated {first:%d %b %Y} – {last:%d %b %Y} · source: {source} · "
           f"refreshed {datetime.now():%d %b %Y %H:%M}")

tabs = st.tabs(["Overview", "Group ▸ Party drill-down", "Month pivot", "Items", "Order status & ageing",
                "Region & salesman", "Dispatch", "Pricing & discount", "Data checks"])

# -----------------------------------------------------------------------------
# 1. Overview
# -----------------------------------------------------------------------------
with tabs[0]:
    kpi_row(f)
    m = summarize(f, ["MONTH", "MONTH_NAME"])
    fig = go.Figure()
    fig.add_bar(x=m["MONTH_NAME"], y=m["Order Amount"], name="Order amount", marker_color=COLORS["order"])
    fig.add_bar(x=m["MONTH_NAME"], y=m["Invoice Amount"], name="Invoice amount", marker_color=COLORS["invoice"])
    fig.add_bar(x=m["MONTH_NAME"], y=m["Shortfall Amount"], name="Shortfall amount", marker_color=COLORS["short"])
    fig.add_scatter(x=m["MONTH_NAME"], y=m["Value Fill %"], name="Value fill %", yaxis="y2",
                    mode="lines+markers", line=dict(color="#222", width=2))
    fig.update_layout(barmode="group", height=430, title="Month-wise order, invoice and shortfall (by SO month)",
                      yaxis=dict(title="₹"), yaxis2=dict(title="Fill %", overlaying="y", side="right", range=[0, 110]),
                      legend=dict(orientation="h", y=1.12), margin=dict(l=10, r=10, t=70, b=10))
    st.plotly_chart(fig, key="ov_month")

    st.subheader("Month-wise summary")
    show_table(present(m), "ov_month_tbl")

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Where the value went")
        status = f.groupby("LINE_STATUS", as_index=False).agg(Lines=("ITEM_KEY", "size"), Amount=("ORDER_AMT", "sum"))
        st.plotly_chart(px.pie(status, names="LINE_STATUS", values="Amount", hole=.55, height=360,
                               color="LINE_STATUS", color_discrete_map={
                                   "Fully invoiced": COLORS["invoice"], "Partially invoiced": "#E9A23B",
                                   "Not invoiced": COLORS["short"], "Extra (not in order)": COLORS["neutral"]}),
                        key="ov_pie")
        st.caption("Order value of lines by invoicing status.")
    with c2:
        st.subheader("Shortfall split")
        disc, pend = f["SHORT_AMT_DISC"].sum(), f["SHORT_AMT_PENDING"].sum()
        st.metric("Pending supply (can still be billed)", short_inr(pend))
        st.metric("Lost to discontinued items", short_inr(disc))
        top = summarize(f, ["PARTY_NAME"]).nlargest(10, "Shortfall Amount")
        st.caption("Top 10 parties by shortfall amount")
        show_table(present(top[["PARTY_NAME", "Order Amount", "Shortfall Amount", "Value Fill %"]]), "ov_top")

# -----------------------------------------------------------------------------
# 2. Party group -> Party -> SO lines drill-down
# -----------------------------------------------------------------------------
with tabs[1]:
    st.markdown("Click a **party group** row to see its parties, then click a **party** row to see its order lines.")
    split = st.toggle("One row per month", value=True, key="split_month")
    by1 = ["PARTY_GROUP"] + (["MONTH", "MONTH_NAME"] if split else [])
    g1 = summarize(f, by1).sort_values((["MONTH"] if split else []) + ["Order Amount"],
                                       ascending=([False] if split else []) + [False])
    g1["Parties"] = g1.merge(f.groupby(by1)["PARTY_NAME"].nunique().reset_index(), on=by1, how="left")["PARTY_NAME"].values
    cols1 = ["PARTY_GROUP"] + (["MONTH_NAME"] if split else []) + ["Parties", "SOs"] + list(METRICS) + ["Qty Fill %", "Value Fill %"]
    st.subheader("Level 1 · Party group")
    sel1 = show_table(present(g1[cols1]), "lvl1", selectable=True, height=420)

    if sel1 is None:
        st.info("Select a party group above to open its party-wise report.")
    else:
        grp = sel1["Party Group"]
        sub = f[f["PARTY_GROUP"] == grp]
        title = grp
        if split:
            sub = sub[sub["MONTH_NAME"] == sel1["Month"]]
            title += f" · {sel1['Month']}"
        st.subheader(f"Level 2 · Parties in {title}")
        g2 = summarize(sub, ["PARTY_NAME", "CLIENT_CODE", "CITY", "STATE", "MONTH", "MONTH_NAME"]) \
            .sort_values(["MONTH", "Order Amount"], ascending=[False, False])
        cols2 = ["PARTY_NAME", "CLIENT_CODE", "CITY", "STATE", "MONTH_NAME", "SOs"] + list(METRICS) + ["Qty Fill %", "Value Fill %"]
        sel2 = show_table(present(g2[cols2]), f"lvl2_{grp}_{sel1.get('Month', '')}", selectable=True)

        if sel2 is None:
            st.info("Select a party to see its sales-order lines.")
        else:
            lines = sub[(sub["PARTY_NAME"] == sel2["Party"]) & (sub["MONTH_NAME"] == sel2["Month"])]
            st.subheader(f"Level 3 · Order lines of {sel2['Party']} · {sel2['Month']}")
            det = lines[["SODOCNO", "SODATE", "ORDNO", "ITEM_NAME", "ORDER_QTY", "ORDER_AVG_RATE", "ORDER_AMT",
                         "INV_QTY", "INV_AMT", "SHORT_QTY", "SHORT_AMT", "BILLS", "FIRST_BILL", "LINE_STATUS",
                         "PARTIAL_STATUS"]].sort_values(["SODOCNO", "SHORT_AMT"], ascending=[True, False])
            det = det.rename(columns={"SODATE": "SO Date", "ORDNO": "Customer PO", "ORDER_QTY": "Order Qty",
                                      "ORDER_AVG_RATE": "Order Rate", "ORDER_AMT": "Order Amount",
                                      "INV_QTY": "Invoice Qty", "INV_AMT": "Invoice Amount",
                                      "SHORT_QTY": "Shortfall Qty", "SHORT_AMT": "Shortfall Amount",
                                      "BILLS": "Bills", "FIRST_BILL": "First Bill Date"}).rename(columns=LABELS)
            det["SO Date"] = det["SO Date"].dt.strftime("%d-%m-%Y")
            det["First Bill Date"] = pd.to_datetime(det["First Bill Date"]).dt.strftime("%d-%m-%Y").fillna("–")
            show_table(det, "lvl3")

# -----------------------------------------------------------------------------
# 3. Month pivot
# -----------------------------------------------------------------------------
with tabs[2]:
    c1, c2 = st.columns(2)
    dims = {"Party Group": "PARTY_GROUP", "Party": "PARTY_NAME", "Item": "ITEM_NAME", "Zone": "ZONE",
            "State": "STATE", "Salesman": "SALESMAN", "Platform": "PLATFORM", "Type": "TYPE"}
    dim = dims[c1.selectbox("Rows", list(dims))]
    metric = c2.selectbox("Value", list(METRICS) + ["Qty Fill %", "Value Fill %"], index=5)

    def pivot(col):
        p = f.pivot_table(index=dim, columns="MONTH", values=col, aggfunc="sum", fill_value=0)
        p["Total"] = p.sum(axis=1)
        return p

    if "Fill" in metric:
        num, den = ("INV_QTY", "ORDER_QTY") if "Qty" in metric else ("INV_AMT", "ORDER_AMT")
        pn, pd_ = pivot(num), pivot(den)
        pv = (pn / pd_.where(pd_ > 0) * 100).loc[pd_["Total"].sort_values(ascending=False).index]
        suffix = " %"
    else:
        pv = pivot(METRICS[metric][0]).sort_values("Total", ascending=False)
        tot = pv.sum().to_frame("TOTAL").T
        pv = pd.concat([pv, tot])
        suffix = " amount" if "Amount" in metric else " qty"
    pv.columns = [month_map.get(c, c) + suffix if c != "Total" else "Total" + suffix for c in pv.columns]
    pv = pv.reset_index().rename(columns={"index": LABELS.get(dim, dim), dim: LABELS.get(dim, dim)})
    st.caption(f"{metric} by {LABELS.get(dim, dim).lower()} and SO month.")
    show_table(pv, "pivot", height=560)

# -----------------------------------------------------------------------------
# 4. Items
# -----------------------------------------------------------------------------
with tabs[3]:
    it = summarize(f, ["ITEM_NAME"]).sort_values("Order Amount", ascending=False)
    disc_items = set(f.loc[f["DISCONTINUED"], "ITEM_NAME"])
    it["Discontinued"] = it["ITEM_NAME"].isin(disc_items).map({True: "Yes", False: ""})
    share = it["Order Amount"].cumsum() / it["Order Amount"].sum() * 100
    it["ABC Class"] = np.select([share <= 70, share <= 90], ["A", "B"], "C")
    it["Parties"] = it["ITEM_NAME"].map(f.groupby("ITEM_NAME")["PARTY_NAME"].nunique())

    c1, c2, c3 = st.columns(3)
    c1.metric("Items ordered", f"{len(it):,}")
    c2.metric("A-class items (70% of value)", f"{(it['ABC Class'] == 'A').sum():,}")
    c3.metric("Items never invoiced", f"{(it['Invoice Qty'] == 0).sum():,}")

    top_short = it.nlargest(15, "Shortfall Amount")
    st.plotly_chart(order_vs_invoice_chart(top_short, "ITEM_NAME", "Top 15 items by shortfall amount", horizontal=True),
                    key="items_chart")
    st.subheader("Item-wise summary with ABC class")
    show_table(present(it[["ITEM_NAME", "ABC Class", "Discontinued", "Parties", "SOs"] + list(METRICS)
                          + ["Qty Fill %", "Value Fill %"]]), "items_tbl", height=480)

    st.subheader("Item × month shortfall qty")
    ip = f.pivot_table(index="ITEM_NAME", columns="MONTH", values="SHORT_QTY", aggfunc="sum", fill_value=0)
    ip = ip[ip.sum(axis=1) > 0]
    if not ip.empty:
        ip.columns = [month_map.get(c, c) for c in ip.columns]
        st.plotly_chart(px.imshow(ip.loc[ip.sum(axis=1).sort_values(ascending=False).index[:25]],
                                  color_continuous_scale="Reds", aspect="auto", text_auto=True,
                                  height=max(350, 24 * min(len(ip), 25))), key="item_heat")

# -----------------------------------------------------------------------------
# 5. Order status & ageing
# -----------------------------------------------------------------------------
with tabs[4]:
    sot = f.groupby("SODOCNO").agg(
        PARTY_NAME=("PARTY_NAME", "first"), PARTY_GROUP=("PARTY_GROUP", "first"), SODATE=("SODATE", "min"),
        ORDNO=("ORDNO", "first"), Items=("ITEM_KEY", "nunique"), **{k: v for k, v in METRICS.items()},
        FIRST_BILL=("FIRST_BILL", "min"), LAST_BILL=("LAST_BILL", "max"), Bills=("BILLS", "max")).reset_index()
    sot["Status"] = np.select([sot["Invoice Qty"] == 0, sot["Shortfall Qty"] == 0],
                              ["Not invoiced", "Fully invoiced"], "Partially invoiced")
    sot["Days to 1st bill"] = (sot["FIRST_BILL"] - sot["SODATE"]).dt.days
    today = pd.Timestamp.today().normalize()
    sot["Age (days)"] = (today - sot["SODATE"]).dt.days
    sot["Age bucket"] = pd.cut(sot["Age (days)"], [-1, 7, 15, 30, 60, 99999],
                               labels=["0–7 days", "8–15 days", "16–30 days", "31–60 days", "60+ days"]).astype(str)
    sot["Value Fill %"] = pct(sot["Invoice Amount"], sot["Order Amount"])

    c = st.columns(4)
    for i, s in enumerate(["Fully invoiced", "Partially invoiced", "Not invoiced"]):
        d = sot[sot["Status"] == s]
        c[i].metric(s, f"{len(d):,} SOs", short_inr(d["Order Amount"].sum()) + " ordered", delta_color="off")
    c[3].metric("Avg days SO → 1st bill", f"{sot['Days to 1st bill'].mean():.1f}" if sot["Days to 1st bill"].notna().any() else "–")

    pending = sot[sot["Status"] != "Fully invoiced"]
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Pending orders by age")
        age = pending.groupby("Age bucket", as_index=False).agg(SOs=("SODOCNO", "count"),
                                                                 **{"Shortfall Amount": ("Shortfall Amount", "sum")})
        order_ = ["0–7 days", "8–15 days", "16–30 days", "31–60 days", "60+ days"]
        age["Age bucket"] = pd.Categorical(age["Age bucket"], order_, ordered=True)
        show_table(age.sort_values("Age bucket").astype({"Age bucket": str}), "age_tbl")
    with c2:
        st.subheader("Billing speed by party group")
        lt = sot.dropna(subset=["Days to 1st bill"]).groupby("PARTY_GROUP", as_index=False).agg(
            SOs=("SODOCNO", "count"), **{"Avg days": ("Days to 1st bill", "mean"),
                                         "Max days": ("Days to 1st bill", "max")}).sort_values("Avg days", ascending=False)
        show_table(present(lt), "lt_tbl")
    if sot["Days to 1st bill"].notna().any():
        st.plotly_chart(px.histogram(sot, x="Days to 1st bill", nbins=30, height=300,
                                     title="Days from SO date to first bill",
                                     color_discrete_sequence=[COLORS["order"]]), key="lt_hist")

    st.subheader("Sales-order register")
    status_pick = st.multiselect("Status", ["Fully invoiced", "Partially invoiced", "Not invoiced"],
                                 default=["Partially invoiced", "Not invoiced"])
    reg = sot[sot["Status"].isin(status_pick)].sort_values("SODATE")
    reg = reg.assign(SODATE=reg["SODATE"].dt.strftime("%d-%m-%Y"),
                     FIRST_BILL=reg["FIRST_BILL"].dt.strftime("%d-%m-%Y"))
    show_table(present(reg[["SODOCNO", "SODATE", "ORDNO", "PARTY_GROUP", "PARTY_NAME", "Items", "Status",
                            "Order Qty", "Invoice Qty", "Shortfall Qty", "Order Amount", "Invoice Amount",
                            "Shortfall Amount", "Value Fill %", "FIRST_BILL", "Days to 1st bill", "Age (days)"]]
                       .rename(columns={"SODATE": "SO Date", "ORDNO": "Customer PO", "FIRST_BILL": "First Bill"})),
               "so_reg", height=450)

# -----------------------------------------------------------------------------
# 6. Region & salesman
# -----------------------------------------------------------------------------
with tabs[5]:
    dim_opts = {"Zone": "ZONE", "State": "STATE", "City": "CITY", "Salesman": "SALESMAN", "Platform": "PLATFORM",
                "Type (Direct/Indirect)": "TYPE", "Payment term": "PAYMENT_TERM"}
    d_label = st.radio("Analyse by", list(dim_opts), horizontal=True, key="region_dim")
    d = dim_opts[d_label]
    rs = summarize(f, [d]).sort_values("Order Amount", ascending=False)
    rs["Parties"] = rs[d].map(f.groupby(d)["PARTY_NAME"].nunique())
    rs["Share of orders %"] = rs["Order Amount"] / rs["Order Amount"].sum() * 100
    st.plotly_chart(order_vs_invoice_chart(rs.head(20), d, f"Order vs invoice by {d_label.lower()}"), key="reg_chart")
    show_table(present(rs[[d, "Parties", "SOs"] + list(METRICS) + ["Qty Fill %", "Value Fill %", "Share of orders %"]]),
               "reg_tbl")
    st.subheader(f"{d_label} × month (order amount)")
    rp = f.pivot_table(index=d, columns="MONTH", values="ORDER_AMT", aggfunc="sum", fill_value=0)
    rp.columns = [month_map.get(c, c) + " amount" for c in rp.columns]
    show_table(rp.reset_index().rename(columns=LABELS), "reg_pivot")

# -----------------------------------------------------------------------------
# 7. Dispatch (invoice side)
# -----------------------------------------------------------------------------
with tabs[6]:
    if inv_f.empty:
        st.info("No invoices are linked to the selected sales orders.")
    else:
        bills = inv_f.groupby("BILL_NO").agg(
            BILL_DATE=("BILL_DATE", "min"), PARTY_NAME=("PARTY_NAME", "first"), PARTY_GROUP=("PARTY_GROUP", "first"),
            TRANSPORT=("TRANSPORT", "first"), DELIVERY=("DELIVERY", "first"), SOs=("SODOCNO", "nunique"),
            Lines=("ITEM_NAME", "size"), Qty=("INV_QTY", "sum"), **{"Invoice Amount": ("INV_AMT", "sum")},
            Packages=("PACKAGE", "sum"), **{"Gross Wt (kg)": ("GROSS_WT", "max")}).reset_index()
        bills["Bill month"] = bills["BILL_DATE"].dt.strftime("%Y-%m")
        c = st.columns(4)
        c[0].metric("Bills", f"{len(bills):,}")
        c[1].metric("Packages", f"{bills['Packages'].sum():,.0f}")
        c[2].metric("Gross weight", f"{bills['Gross Wt (kg)'].sum():,.0f} kg")
        c[3].metric("Avg bill value", short_inr(bills["Invoice Amount"].mean()))
        st.caption("Gross weight is read once per bill (it repeats on every line of the same bill); packages are summed.")

        bm = bills.groupby("Bill month", as_index=False).agg(Bills=("BILL_NO", "count"),
                                                              **{"Invoice Amount": ("Invoice Amount", "sum")})
        st.plotly_chart(px.bar(bm, x="Bill month", y="Invoice Amount", text="Bills", height=330,
                               title="Invoicing by bill month (labels = number of bills)",
                               color_discrete_sequence=[COLORS["invoice"]]), key="bill_month")
        c1, c2 = st.columns(2)
        with c1:
            st.subheader("By transporter")
            tr = bills.groupby("TRANSPORT", as_index=False).agg(
                Bills=("BILL_NO", "count"), Packages=("Packages", "sum"), **{"Gross Wt (kg)": ("Gross Wt (kg)", "sum"),
                                                                            "Invoice Amount": ("Invoice Amount", "sum")})
            tr["Value per kg"] = tr["Invoice Amount"] / tr["Gross Wt (kg)"].where(tr["Gross Wt (kg)"] > 0)
            show_table(present(tr.sort_values("Invoice Amount", ascending=False)), "tr_tbl")
        with c2:
            st.subheader("By delivery location")
            dl = bills.groupby("DELIVERY", as_index=False).agg(
                Bills=("BILL_NO", "count"), **{"Invoice Amount": ("Invoice Amount", "sum")}) \
                .sort_values("Invoice Amount", ascending=False)
            show_table(present(dl), "dl_tbl", height=350)
        st.subheader("Bill register")
        br = bills.sort_values("BILL_DATE", ascending=False).assign(BILL_DATE=lambda x: x["BILL_DATE"].dt.strftime("%d-%m-%Y"))
        show_table(present(br.drop(columns=["Bill month"]).rename(columns={"BILL_DATE": "Bill Date"})), "bill_reg", height=400)

# -----------------------------------------------------------------------------
# 8. Pricing & discount
# -----------------------------------------------------------------------------
with tabs[7]:
    pg = f.groupby("PARTY_GROUP", as_index=False).agg(
        **{"Order Amount": ("ORDER_AMT", "sum"), "MRP Value": ("MRP_VALUE", "sum")})
    pg["Discount from MRP %"] = (1 - pg["Order Amount"] / pg["MRP Value"].where(pg["MRP Value"] > 0)) * 100
    pg = pg.sort_values("Discount from MRP %", ascending=False)
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Billing rate vs MRP by party group")
        st.caption("How far below MRP each group buys (order value ÷ MRP value).")
        show_table(present(pg), "mrp_tbl")
    with c2:
        st.subheader("Invoice discounts")
        if inv_f["DISC_AMT"].sum() > 0:
            dg = inv_f.groupby("PARTY_GROUP", as_index=False).agg(
                **{"Invoice Amount": ("INV_AMT", "sum"), "Discount Amount": ("DISC_AMT", "sum")})
            dg["Discount %"] = dg["Discount Amount"] / (dg["Invoice Amount"] + dg["Discount Amount"]) * 100
            show_table(present(dg.sort_values("Discount Amount", ascending=False)), "disc_tbl")
        else:
            st.info("No invoice discounts (DISC. AMOUNT) are recorded for the selected orders.")

    st.subheader("Invoice rate different from order rate")
    rv = f[(f["ORDER_QTY"] > 0) & (f["INV_QTY"] > 0)].copy()
    rv["Invoice Rate"] = rv["INV_AMT"] / rv["INV_QTY"]
    rv["Rate diff %"] = (rv["Invoice Rate"] / rv["ORDER_AVG_RATE"].where(rv["ORDER_AVG_RATE"] > 0) - 1) * 100
    tol = st.slider("Show lines where the difference is more than (%)", 0.0, 20.0, 1.0, 0.5)
    rv = rv[rv["Rate diff %"].abs() > tol]
    if rv.empty:
        st.success(f"Every invoiced line is within ±{tol}% of the order rate.")
    else:
        show_table(present(rv[["SODOCNO", "PARTY_NAME", "ITEM_NAME", "ORDER_AVG_RATE", "Invoice Rate", "Rate diff %",
                               "INV_QTY", "INV_AMT"]].rename(columns={"ORDER_AVG_RATE": "Order Rate",
                                                                     "INV_QTY": "Invoice Qty", "INV_AMT": "Invoice Amount"})
                           .sort_values("Rate diff %")), "rate_var")

# -----------------------------------------------------------------------------
# 9. Data checks
# -----------------------------------------------------------------------------
with tabs[8]:
    unm = fact[fact["PARTY_GROUP"] == "UNMAPPED PARTY"].groupby("PARTY_NAME", as_index=False) \
        .agg(SOs=("SODOCNO", "nunique"), **{"Order Amount": ("ORDER_AMT", "sum")})
    orphan = dq["orphan_inv"].groupby(["SODOCNO", "BILL_NO", "INV_PARTY_NAME"], as_index=False) \
        .agg(Lines=("ITEM_NAME", "size"), **{"Invoice Amount": ("INV_AMT", "sum")})
    extra = fact[fact["LINE_STATUS"] == "Extra (not in order)"]

    c = st.columns(4)
    c[0].metric("Parties not in client master", f"{len(unm):,}")
    c[1].metric("Invoices with unknown SODOCNO", f"{orphan['BILL_NO'].nunique():,}")
    c[2].metric("Invoiced items not on the SO", f"{len(extra):,}")
    c[3].metric("Duplicate rows ignored", f"{dq['dup_orders'] + dq['dup_invoices']:,}")

    st.subheader("Order parties missing from Client master")
    st.caption("Add these names to Client master (or fix the spelling) so they roll up to the right party group.")
    if len(unm):
        show_table(present(unm.sort_values("Order Amount", ascending=False)), "dq_unm")
    else:
        st.success("All parties are mapped.")
    st.subheader("Invoices whose SODOCNO is not in the Sale order sheet")
    st.caption("These are left out of every report above, because their SO month and party group are unknown.")
    if len(orphan):
        show_table(present(orphan.rename(columns={"INV_PARTY_NAME": "Party"})), "dq_orphan")
    else:
        st.success("Every invoice links to a sales order.")
    st.subheader("Items invoiced against an SO that does not list them")
    st.caption("Usually the item name is spelt differently in the two sheets.")
    if len(extra):
        show_table(present(extra[["SODOCNO", "PARTY_NAME", "ITEM_NAME", "INV_QTY", "INV_AMT"]]
                           .rename(columns={"INV_QTY": "Invoice Qty", "INV_AMT": "Invoice Amount"})), "dq_extra")
    else:
        st.success("Every invoiced item matches an ordered item.")
    if len(dq["multi_party_so"]):
        st.subheader("SODOCNO used for more than one party")
        show_table(present(dq["multi_party_so"]), "dq_multi")
    if dq["bad_sodate"] or dq["bad_billdate"]:
        st.warning(f"{dq['bad_sodate']} order rows have an unreadable SODATE and "
                   f"{dq['bad_billdate']} invoice rows have an unreadable BILL DATE.")

# =============================================================================
# Excel download (respects the filters)
# =============================================================================
with st.sidebar:
    st.header("Export")
    xls = to_excel({
        "Month summary": present(summarize(f, ["MONTH", "MONTH_NAME"])),
        "Group x Month": present(summarize(f, ["PARTY_GROUP", "MONTH", "MONTH_NAME"])),
        "Party x Month": present(summarize(f, ["PARTY_GROUP", "PARTY_NAME", "MONTH", "MONTH_NAME"])),
        "Item summary": present(summarize(f, ["ITEM_NAME"])),
        "Order lines": present(f[["MONTH_NAME", "SODOCNO", "SODATE", "PARTY_GROUP", "PARTY_NAME", "ITEM_NAME",
                                  "ORDER_QTY", "ORDER_AMT", "INV_QTY", "INV_AMT", "SHORT_QTY", "SHORT_AMT",
                                  "LINE_STATUS", "PARTIAL_STATUS"]]),
    })
    st.download_button("⬇️ Download Excel report", xls, file_name=f"order_vs_invoice_{datetime.now():%Y%m%d}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       width="stretch")
