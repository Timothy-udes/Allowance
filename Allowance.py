import streamlit as st
import pandas as pd
import sqlite3
from io import BytesIO
from rapidfuzz import process, fuzz
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

# ══════════════════════════════════════════════
# PAGE CONFIG
# ══════════════════════════════════════════════
st.set_page_config(
    page_title="Driver's Allowance Machine · Opti360",
    page_icon="🔶",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ══════════════════════════════════════════════
# GLOBAL CSS
# ══════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@500;600;700;800&family=DM+Sans:ital,opsz,wght@0,9..40,300;0,9..40,400;0,9..40,500;1,9..40,300&display=swap');

/* ─ Base ─ */
html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
.stApp { background: #08090F; }
section[data-testid="stSidebar"] { background: #0D0F1A !important; border-right: 1px solid #1A1E30; }
section[data-testid="stSidebar"] > div { padding-top: 0 !important; }

/* ─ Scrollbar ─ */
::-webkit-scrollbar { width: 4px; height: 4px; }
::-webkit-scrollbar-track { background: transparent; }
::-webkit-scrollbar-thumb { background: #2A2E42; border-radius: 4px; }
::-webkit-scrollbar-thumb:hover { background: #F5A623; }

/* ─ Sidebar brand ─ */
.sb-brand {
    padding: 28px 24px 20px;
    border-bottom: 1px solid #1A1E30;
    margin-bottom: 8px;
}
.sb-logo {
    display: flex; align-items: center; gap: 12px; margin-bottom: 16px;
}
.sb-logo-mark {
    width: 40px; height: 40px;
    background: #F5A623;
    border-radius: 10px;
    display: flex; align-items: center; justify-content: center;
    font-family: 'Syne', sans-serif;
    font-weight: 800; font-size: 13px; color: #08090F;
    flex-shrink: 0; letter-spacing: -0.3px;
}
.sb-brand-name {
    font-family: 'Syne', sans-serif;
    font-weight: 800; font-size: 11px;
    letter-spacing: 3px; text-transform: uppercase; color: #F5A623;
    line-height: 1; margin-bottom: 3px;
}
.sb-app-name {
    font-family: 'Syne', sans-serif;
    font-weight: 700; font-size: 14px; color: #FFFFFF; line-height: 1.2;
}
.sb-tagline {
    font-size: 11px; color: #3A3E55; font-style: italic;
    line-height: 1.5; letter-spacing: 0.1px;
}

/* ─ Sidebar workflow steps ─ */
.sb-section-label {
    font-size: 10px; font-weight: 500; letter-spacing: 2px;
    text-transform: uppercase; color: #2E3248;
    padding: 16px 24px 8px; display: block;
}
.wf-step {
    display: flex; align-items: center; gap: 12px;
    padding: 10px 24px; cursor: default;
    border-left: 2px solid transparent;
    transition: all 0.15s;
}
.wf-step.active   { border-left-color: #F5A623; background: rgba(245,166,35,0.05); }
.wf-step.done     { border-left-color: #22C55E; background: rgba(34,197,94,0.04); }
.wf-step.inactive { opacity: 0.35; }
.wf-dot {
    width: 26px; height: 26px; border-radius: 7px; flex-shrink: 0;
    display: flex; align-items: center; justify-content: center;
    font-family: 'Syne', sans-serif; font-weight: 700; font-size: 11px;
}
.wf-step.active   .wf-dot { background: #F5A623; color: #08090F; }
.wf-step.done     .wf-dot { background: #22C55E; color: #08090F; }
.wf-step.inactive .wf-dot { background: #1A1E30; color: #3A3E55; }
.wf-step-text { display: flex; flex-direction: column; }
.wf-step-name  { font-size: 12px; font-weight: 500; color: #C8CDE0; line-height: 1.2; }
.wf-step-desc  { font-size: 10px; color: #3A3E55; margin-top: 1px; }
.wf-step.active .wf-step-name { color: #FFFFFF; }
.wf-step.done   .wf-step-name { color: #86EFAC; }

/* ─ Sidebar system info ─ */
.sb-info-card {
    margin: 16px 16px 0;
    background: #0D1020; border: 1px solid #1A1E30;
    border-radius: 10px; padding: 14px 16px;
}
.sb-info-row {
    display: flex; justify-content: space-between; align-items: center;
    padding: 5px 0; border-bottom: 1px solid #13172A;
}
.sb-info-row:last-child { border-bottom: none; }
.sb-info-label { font-size: 11px; color: #2E3248; }
.sb-info-val   { font-size: 11px; font-weight: 500; color: #5A6280; }
.sb-info-val.amber { color: #F5A623; }
.sb-info-val.green { color: #22C55E; }

/* ─ Main header ─ */
.main-header {
    display: flex; align-items: center; justify-content: space-between;
    padding: 32px 0 24px;
    border-bottom: 1px solid #14172A;
    margin-bottom: 32px;
}
.mh-left { display: flex; align-items: center; gap: 16px; }
.mh-logo {
    width: 48px; height: 48px; background: #F5A623;
    border-radius: 12px; display: flex; align-items: center; justify-content: center;
    font-family: 'Syne', sans-serif; font-weight: 800; font-size: 14px; color: #08090F;
}
.mh-title {
    font-family: 'Syne', sans-serif; font-weight: 800;
    font-size: 22px; color: #FFFFFF; letter-spacing: -0.5px; line-height: 1.1;
}
.mh-subtitle { font-size: 12px; color: #2E3248; margin-top: 3px; letter-spacing: 0.3px; }
.mh-badge {
    background: rgba(245,166,35,0.1); border: 1px solid rgba(245,166,35,0.25);
    border-radius: 20px; padding: 5px 14px;
    font-size: 11px; font-weight: 500; color: #F5A623; letter-spacing: 0.5px;
}

/* ─ Section headers ─ */
.section-head {
    display: flex; align-items: center; gap: 14px; margin-bottom: 16px; margin-top: 8px;
}
.section-num {
    width: 28px; height: 28px; border-radius: 8px;
    background: #F5A623; color: #08090F;
    font-family: 'Syne', sans-serif; font-weight: 800; font-size: 12px;
    display: flex; align-items: center; justify-content: center; flex-shrink: 0;
}
.section-num.done { background: #22C55E; }
.section-label {
    font-family: 'Syne', sans-serif; font-weight: 700;
    font-size: 16px; color: #FFFFFF; letter-spacing: -0.2px;
}
.section-hint { font-size: 11px; color: #2E3248; margin-top: 1px; }

/* ─ Upload zone ─ */
.upload-zone-wrap {
    background: #0D0F1A;
    border: 1px solid #1A1E30;
    border-radius: 12px;
    padding: 4px 4px 4px;
    margin-bottom: 4px;
}
[data-testid="stFileUploader"] {
    background: transparent !important;
}
[data-testid="stFileUploaderDropzone"] {
    background: rgba(245,166,35,0.025) !important;
    border: 1.5px dashed rgba(245,166,35,0.22) !important;
    border-radius: 10px !important;
    transition: all 0.2s !important;
}
[data-testid="stFileUploaderDropzone"]:hover {
    background: rgba(245,166,35,0.05) !important;
    border-color: rgba(245,166,35,0.5) !important;
}
[data-testid="stFileUploaderDropzone"] p { color: #3A3E55 !important; font-size: 13px !important; }
[data-testid="stFileUploader"] label { color: #4A4E65 !important; font-size: 12px !important; }
[data-testid="stFileUploader"] small { color: #2A2E42 !important; }

/* ─ Column schema pills ─ */
.schema-wrap {
    display: flex; flex-wrap: wrap; gap: 6px; margin-top: 10px; margin-bottom: 6px;
}
.schema-pill {
    background: #0D0F1A; border: 1px solid #1E2235;
    border-radius: 6px; padding: 4px 10px;
    font-family: 'DM Sans', monospace; font-size: 11px;
    color: #3A4060; letter-spacing: 0.2px;
}
.schema-pill.key {
    border-color: rgba(245,166,35,0.3); color: #8A7040; background: rgba(245,166,35,0.04);
}

/* ─ Stat cards ─ */
.kpi-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(130px, 1fr));
    gap: 10px; margin: 16px 0;
}
.kpi-card {
    background: #0D0F1A;
    border: 1px solid #1A1E30;
    border-radius: 10px; padding: 14px 16px;
    position: relative; overflow: hidden;
}
.kpi-card::before {
    content: ''; position: absolute;
    top: 0; left: 0; right: 0; height: 2px;
    background: #F5A623; opacity: 0.6;
}
.kpi-card.green::before { background: #22C55E; }
.kpi-card.red::before   { background: #EF4444; }
.kpi-card.blue::before  { background: #3B82F6; }
.kpi-val {
    font-family: 'Syne', sans-serif; font-weight: 700;
    font-size: 24px; color: #FFFFFF; line-height: 1;
    margin-bottom: 5px; letter-spacing: -0.5px;
}
.kpi-val.amber { color: #F5A623; }
.kpi-val.green { color: #22C55E; }
.kpi-val.red   { color: #EF4444; }
.kpi-val.blue  { color: #60A5FA; }
.kpi-lbl {
    font-size: 10px; font-weight: 500; color: #2E3248;
    text-transform: uppercase; letter-spacing: 1px;
}

/* ─ Alert boxes ─ */
.alert {
    border-radius: 8px; padding: 12px 16px;
    font-size: 13px; line-height: 1.5; margin: 12px 0;
    display: flex; align-items: flex-start; gap: 10px;
}
.alert-icon { font-size: 14px; flex-shrink: 0; margin-top: 1px; }
.alert.success {
    background: rgba(34,197,94,0.06); border: 1px solid rgba(34,197,94,0.2);
    color: #86EFAC;
}
.alert.warning {
    background: rgba(245,166,35,0.06); border: 1px solid rgba(245,166,35,0.2);
    color: #C8A96E;
}
.alert.error {
    background: rgba(239,68,68,0.06); border: 1px solid rgba(239,68,68,0.2);
    color: #FCA5A5;
}
.alert.info {
    background: rgba(59,130,246,0.06); border: 1px solid rgba(59,130,246,0.2);
    color: #93C5FD;
}

/* ─ Table section label ─ */
.tbl-label {
    font-family: 'Syne', sans-serif; font-weight: 700;
    font-size: 13px; color: #4A4E65;
    text-transform: uppercase; letter-spacing: 1.5px;
    margin: 24px 0 10px; display: flex; align-items: center; gap: 8px;
}
.tbl-label::after {
    content: ''; flex: 1; height: 1px; background: #14172A;
}

/* ─ Download button row ─ */
.dl-row {
    display: grid; grid-template-columns: 1fr 1fr; gap: 10px;
    margin: 20px 0 8px;
}
.dl-card {
    background: #0D0F1A; border: 1px solid #1A1E30;
    border-radius: 10px; padding: 16px 18px;
}
.dl-card-title {
    font-family: 'Syne', sans-serif; font-weight: 700;
    font-size: 13px; color: #FFFFFF; margin-bottom: 3px;
}
.dl-card-desc { font-size: 11px; color: #2E3248; margin-bottom: 10px; }

/* ─ Download buttons ─ */
.stDownloadButton button {
    background: #F5A623 !important;
    color: #08090F !important;
    font-family: 'Syne', sans-serif !important;
    font-weight: 700 !important; font-size: 12px !important;
    letter-spacing: 0.3px !important;
    border: none !important; border-radius: 7px !important;
    padding: 8px 18px !important; width: 100% !important;
    transition: opacity 0.15s !important;
}
.stDownloadButton button:hover { opacity: 0.85 !important; }

/* ─ Expanders ─ */
[data-testid="stExpander"] {
    background: #0D0F1A !important;
    border: 1px solid #1A1E30 !important;
    border-radius: 10px !important;
    margin-bottom: 10px;
}
[data-testid="stExpander"] summary {
    font-size: 12px !important; font-weight: 500 !important;
    color: #4A4E65 !important; padding: 12px 16px !important;
}
[data-testid="stExpander"] summary:hover { color: #FFFFFF !important; }

/* ─ Dataframe ─ */
[data-testid="stDataFrame"] iframe {
    border-radius: 8px !important;
}

/* ─ Spinner ─ */
[data-testid="stSpinner"] p { color: #4A4E65 !important; font-size: 12px !important; }

/* ─ Separator ─ */
.rule { border: none; border-top: 1px solid #14172A; margin: 28px 0; }

/* ─ Footer ─ */
.footer {
    text-align: center; padding: 32px 0 16px;
    font-size: 11px; color: #1E2235; letter-spacing: 0.5px;
}
.footer span { color: #2E3248; }

/* ─ Match quality badge ─ */
.mq-exact  { color: #22C55E; font-weight: 600; }
.mq-fuzzy  { color: #F5A623; font-weight: 500; }
.mq-none   { color: #EF4444; font-weight: 500; }

/* ─ Stray st elements ─ */
.stAlert { border-radius: 8px !important; }
div[data-testid="stMarkdownContainer"] p { color: #6A6E85; }
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════
# STATE
# ══════════════════════════════════════════════
if "db_loaded" not in st.session_state:
    st.session_state.db_loaded = False
if "report_processed" not in st.session_state:
    st.session_state.report_processed = False
if "db_stats" not in st.session_state:
    st.session_state.db_stats = {}


# ══════════════════════════════════════════════
# SQLITE
# ══════════════════════════════════════════════
DB_FILE = "driver_db.sqlite3"
conn = sqlite3.connect(DB_FILE, check_same_thread=False)


# ══════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════
def read_file(file):
    name = file.name.lower()
    try:
        if name.endswith((".xlsx", ".xls", ".xlsm")):
            return pd.read_excel(file)
        elif name.endswith(".csv"):
            return pd.read_csv(file)
        elif name.endswith(".ods"):
            return pd.read_excel(file, engine="odf")
        else:
            st.markdown('<div class="alert error"><span class="alert-icon">✕</span>Unsupported format. Use xlsx, xls, xlsm, csv or ods.</div>', unsafe_allow_html=True)
            return None
    except Exception as e:
        st.markdown(f'<div class="alert error"><span class="alert-icon">✕</span>Could not read file: {e}</div>', unsafe_allow_html=True)
        return None


def styled_excel(df, sheet_name="Sheet1"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        wb = writer.book
        ws = writer.sheets[sheet_name]

        AMBER  = "F5A623"
        NAVY   = "08090F"
        MID    = "0D0F1A"
        DARK   = "111520"
        TEXT   = "C8CDE0"
        MUTED  = "3A3E55"

        hdr_font  = Font(bold=True, color=NAVY, name="Calibri", size=10)
        hdr_fill  = PatternFill("solid", fgColor=AMBER)
        thin_side = Side(style="thin", color="1A1E30")
        border    = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
        center    = Alignment(horizontal="center", vertical="center", wrap_text=False)

        for ci, col in enumerate(df.columns, 1):
            cell = ws.cell(row=1, column=ci)
            cell.font      = hdr_font
            cell.fill      = hdr_fill
            cell.alignment = center
            cell.border    = border
            max_w = max(df[col].astype(str).map(len).max(), len(col)) + 3
            ws.column_dimensions[get_column_letter(ci)].width = min(max_w, 38)

        ws.row_dimensions[1].height = 22

        for ri, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row), start=2):
            fill = PatternFill("solid", fgColor=MID if ri % 2 == 0 else DARK)
            for cell in row:
                cell.font      = Font(color=TEXT, name="Calibri", size=10)
                cell.fill      = fill
                cell.alignment = center
                cell.border    = border
                if "AMOUNT" in (df.columns[cell.column - 1] if cell.column <= len(df.columns) else ""):
                    cell.number_format = "#,##0.00"

        ws.freeze_panes = "A2"

        last = ws.max_row + 2
        note = ws.cell(row=last, column=1, value="Opti360 · Driver's Allowance Machine")
        note.font = Font(color="1E2235", italic=True, size=8, name="Calibri")

    return output.getvalue()


# ══════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════
with st.sidebar:
    st.markdown("""
    <div class="sb-brand">
        <div class="sb-logo">
            <div class="sb-logo-mark">O360</div>
            <div>
                <div class="sb-brand-name">Opti360</div>
                <div class="sb-app-name">Driver's Allowance<br>Machine</div>
            </div>
        </div>
        <div class="sb-tagline">Automated payment matching &amp;<br>bank transfer generation</div>
    </div>
    """, unsafe_allow_html=True)

    step1_class = "done" if st.session_state.db_loaded else "active"
    step2_class = "done" if st.session_state.report_processed else ("active" if st.session_state.db_loaded else "inactive")
    step3_class = "done" if st.session_state.report_processed else "inactive"

    st.markdown(f"""
    <span class="sb-section-label">Workflow</span>

    <div class="wf-step {step1_class}">
        <div class="wf-dot">{"✓" if st.session_state.db_loaded else "1"}</div>
        <div class="wf-step-text">
            <span class="wf-step-name">Driver Database</span>
            <span class="wf-step-desc">Upload master account list</span>
        </div>
    </div>

    <div class="wf-step {step2_class}">
        <div class="wf-dot">{"✓" if st.session_state.report_processed else "2"}</div>
        <div class="wf-step-text">
            <span class="wf-step-name">Driver Report</span>
            <span class="wf-step-desc">Upload payout amounts</span>
        </div>
    </div>

    <div class="wf-step {step3_class}">
        <div class="wf-dot">{"✓" if st.session_state.report_processed else "3"}</div>
        <div class="wf-step-text">
            <span class="wf-step-name">Export</span>
            <span class="wf-step-desc">Download payment files</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    if st.session_state.db_stats:
        s = st.session_state.db_stats
        st.markdown(f"""
        <div class="sb-info-card">
            <div class="sb-info-row">
                <span class="sb-info-label">Drivers loaded</span>
                <span class="sb-info-val amber">{s.get("total", "—")}</span>
            </div>
            <div class="sb-info-row">
                <span class="sb-info-label">Banks</span>
                <span class="sb-info-val">{s.get("banks", "—")}</span>
            </div>
            <div class="sb-info-row">
                <span class="sb-info-label">With accounts</span>
                <span class="sb-info-val green">{s.get("with_acct", "—")}</span>
            </div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<div style='height:1px'/>", unsafe_allow_html=True)
    st.markdown("""
    <div style='padding:16px 24px;'>
        <div style='font-size:10px;color:#1E2235;letter-spacing:1px;text-transform:uppercase;margin-bottom:8px;'>Supported formats</div>
        <div style='display:flex;flex-wrap:wrap;gap:5px;'>
            <span style='font-size:10px;background:#0D0F1A;border:1px solid #1A1E30;border-radius:5px;padding:3px 8px;color:#2E3248;'>.xlsx</span>
            <span style='font-size:10px;background:#0D0F1A;border:1px solid #1A1E30;border-radius:5px;padding:3px 8px;color:#2E3248;'>.xls</span>
            <span style='font-size:10px;background:#0D0F1A;border:1px solid #1A1E30;border-radius:5px;padding:3px 8px;color:#2E3248;'>.xlsm</span>
            <span style='font-size:10px;background:#0D0F1A;border:1px solid #1A1E30;border-radius:5px;padding:3px 8px;color:#2E3248;'>.csv</span>
            <span style='font-size:10px;background:#0D0F1A;border:1px solid #1A1E30;border-radius:5px;padding:3px 8px;color:#2E3248;'>.ods</span>
        </div>
    </div>
    """, unsafe_allow_html=True)


# ══════════════════════════════════════════════
# MAIN AREA — HEADER
# ══════════════════════════════════════════════
st.markdown("""
<div class="main-header">
    <div class="mh-left">
        <div class="mh-logo">O360</div>
        <div>
            <div class="mh-title">Driver's Allowance Machine</div>
            <div class="mh-subtitle">Upload · Match · Export · Pay</div>
        </div>
    </div>
    <div class="mh-badge">Opti360 Fleet Suite</div>
</div>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════
# STEP 1 — DRIVER DATABASE
# ══════════════════════════════════════════════
st.markdown(f"""
<div class="section-head">
    <div class="section-num {'done' if st.session_state.db_loaded else ''}">
        {'✓' if st.session_state.db_loaded else '1'}
    </div>
    <div>
        <div class="section-label">Driver Database</div>
        <div class="section-hint">Master list of drivers with bank account details</div>
    </div>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="schema-wrap">
    <span class="schema-pill key">S/N</span>
    <span class="schema-pill key">FMS DRIVER'S NAME</span>
    <span class="schema-pill">ACCOUNT NAME</span>
    <span class="schema-pill">ACCOUNT NO</span>
    <span class="schema-pill">BANK</span>
</div>
""", unsafe_allow_html=True)

driver_file = st.file_uploader(
    "Driver database spreadsheet",
    type=["xlsx","xls","xlsm","csv","ods"],
    key="driver_db",
    label_visibility="collapsed",
)

if driver_file:
    df_drivers = read_file(driver_file)
    if df_drivers is not None:
        df_drivers.columns = df_drivers.columns.str.strip().str.upper()
        df_drivers = df_drivers.rename(columns={
            "FMS DRIVER'S NAME": "DRIVER_NAME",
            "ACCOUNT NAME":      "ACCOUNT_NAME",
            "ACCOUNT NO":        "ACCOUNT_NO",
            "BANK":              "BANK",
        })
        required = ["S/N", "DRIVER_NAME", "ACCOUNT_NAME", "ACCOUNT_NO", "BANK"]
        if all(c in df_drivers.columns for c in required):
            df_drivers["DRIVER_NAME"] = df_drivers["DRIVER_NAME"].str.strip().str.upper()
            df_drivers["ACCOUNT_NO"]  = df_drivers["ACCOUNT_NO"].astype(str).str.zfill(10)
            df_drivers.to_sql("drivers", conn, if_exists="replace", index=False)

            total   = len(df_drivers)
            banks   = df_drivers["BANK"].nunique()
            w_acct  = int(df_drivers["ACCOUNT_NO"].notna().sum())

            st.session_state.db_loaded = True
            st.session_state.db_stats  = {"total": total, "banks": banks, "with_acct": w_acct}

            st.markdown(f"""
            <div class="alert success">
                <span class="alert-icon">✓</span>
                <span>Database loaded — <strong>{total}</strong> drivers across <strong>{banks}</strong> bank(s). All records saved.</span>
            </div>
            """, unsafe_allow_html=True)

            st.markdown(f"""
            <div class="kpi-grid">
                <div class="kpi-card">
                    <div class="kpi-val amber">{total}</div>
                    <div class="kpi-lbl">Total Drivers</div>
                </div>
                <div class="kpi-card green">
                    <div class="kpi-val green">{banks}</div>
                    <div class="kpi-lbl">Banks</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-val">{w_acct}</div>
                    <div class="kpi-lbl">With Account No.</div>
                </div>
                <div class="kpi-card {'red' if total - w_acct > 0 else 'green'}">
                    <div class="kpi-val {'red' if total - w_acct > 0 else 'green'}">{total - w_acct}</div>
                    <div class="kpi-lbl">Missing Accounts</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            with st.expander("Preview database records", expanded=False):
                st.dataframe(df_drivers.head(20), use_container_width=True, hide_index=True)
        else:
            missing_cols = [c for c in required if c not in df_drivers.columns]
            st.markdown(f'<div class="alert error"><span class="alert-icon">✕</span>Missing required columns: <strong>{", ".join(missing_cols)}</strong></div>', unsafe_allow_html=True)

st.markdown("<hr class='rule'/>", unsafe_allow_html=True)


# ══════════════════════════════════════════════
# STEP 2 — DRIVER REPORT
# ══════════════════════════════════════════════
st.markdown(f"""
<div class="section-head">
    <div class="section-num {'done' if st.session_state.report_processed else ''}">
        {'✓' if st.session_state.report_processed else '2'}
    </div>
    <div>
        <div class="section-label">Driver Report</div>
        <div class="section-hint">System Manager export with driver names and amounts due</div>
    </div>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="schema-wrap">
    <span class="schema-pill key">DRIVER NAME</span>
    <span class="schema-pill key">TOTAL AMOUNT</span>
    <span class="schema-pill" style="color:#1E2235;border-color:#14172A;">+ any other columns</span>
</div>
""", unsafe_allow_html=True)

report_file = st.file_uploader(
    "Driver report spreadsheet",
    type=["xlsx","xls","xlsm","csv","ods"],
    key="driver_report",
    label_visibility="collapsed",
)

if report_file:
    df_report = read_file(report_file)
    if df_report is not None:
        df_report.columns = df_report.columns.str.strip().str.upper()
        df_report = df_report.rename(columns={
            "DRIVER NAME":  "DRIVER_NAME",
            "TOTAL AMOUNT": "AMOUNT",
        })

        if "DRIVER_NAME" not in df_report.columns or "AMOUNT" not in df_report.columns:
            missing_cols = []
            if "DRIVER_NAME" not in df_report.columns: missing_cols.append("DRIVER NAME")
            if "AMOUNT" not in df_report.columns:      missing_cols.append("TOTAL AMOUNT")
            st.markdown(f'<div class="alert error"><span class="alert-icon">✕</span>Missing required columns: <strong>{", ".join(missing_cols)}</strong></div>', unsafe_allow_html=True)
        else:
            df_report["DRIVER_NAME"] = df_report["DRIVER_NAME"].str.strip().str.upper()
            df_agg = df_report.groupby("DRIVER_NAME", as_index=False)["AMOUNT"].sum()

            try:
                df_db = pd.read_sql("SELECT * FROM drivers", conn)
            except Exception:
                st.markdown('<div class="alert warning"><span class="alert-icon">⚠</span>No driver database found. Please complete Step 1 first.</div>', unsafe_allow_html=True)
                st.stop()

            # ── Fuzzy match ──
            with st.spinner("Matching driver names…"):
                matched, scores = [], []
                for name in df_agg["DRIVER_NAME"]:
                    result = process.extractOne(name, df_db["DRIVER_NAME"].tolist(), scorer=fuzz.token_sort_ratio)
                    if result and result[1] >= 80:
                        matched.append(result[0]); scores.append(result[1])
                    else:
                        matched.append(None); scores.append(0)
                df_agg["MATCHED_NAME"] = matched
                df_agg["MATCH_SCORE"]  = scores

            # ── Merge ──
            final_df = pd.merge(df_agg, df_db, left_on="MATCHED_NAME", right_on="DRIVER_NAME", how="left")
            final_df["S/N"] = range(1, len(final_df) + 1)
            final_df = final_df[["S/N","DRIVER_NAME_x","AMOUNT","ACCOUNT_NAME","ACCOUNT_NO","BANK","MATCH_SCORE"]]
            final_df = final_df.rename(columns={"DRIVER_NAME_x": "DRIVER_NAME"})
            final_df["ACCOUNT_NO"] = final_df["ACCOUNT_NO"].astype(str).str.zfill(10)

            st.session_state.report_processed = True

            # ── KPIs ──
            total_d    = len(final_df)
            total_amt  = final_df["AMOUNT"].sum()
            matched_n  = int(final_df["ACCOUNT_NO"].notna().sum())
            unmatch_n  = total_d - matched_n
            exact_n    = int((final_df["MATCH_SCORE"] == 100).sum())

            st.markdown(f"""
            <div class="kpi-grid">
                <div class="kpi-card">
                    <div class="kpi-val amber">{total_d}</div>
                    <div class="kpi-lbl">Drivers</div>
                </div>
                <div class="kpi-card blue">
                    <div class="kpi-val blue">₦{total_amt:,.0f}</div>
                    <div class="kpi-lbl">Total Payout</div>
                </div>
                <div class="kpi-card green">
                    <div class="kpi-val green">{matched_n}</div>
                    <div class="kpi-lbl">Matched</div>
                </div>
                <div class="kpi-card {'red' if unmatch_n > 0 else 'green'}">
                    <div class="kpi-val {'red' if unmatch_n > 0 else 'green'}">{unmatch_n}</div>
                    <div class="kpi-lbl">Unmatched</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-val">{exact_n}</div>
                    <div class="kpi-lbl">Exact Matches</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-val">{matched_n - exact_n}</div>
                    <div class="kpi-lbl">Fuzzy Matches</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            # ── Unmatched warning ──
            missing_accts = final_df[final_df["ACCOUNT_NO"].isna() | (final_df["ACCOUNT_NO"].astype(str) == "nan")]
            if not missing_accts.empty:
                st.markdown(f"""
                <div class="alert warning">
                    <span class="alert-icon">⚠</span>
                    <span><strong>{len(missing_accts)} driver(s)</strong> could not be matched to account details. 
                    They will be excluded from the bank transfer sheet. Review before exporting.</span>
                </div>
                """, unsafe_allow_html=True)
                with st.expander(f"Unmatched drivers ({len(missing_accts)})", expanded=True):
                    st.dataframe(
                        missing_accts[["S/N","DRIVER_NAME","AMOUNT","MATCH_SCORE"]],
                        use_container_width=True, hide_index=True
                    )

            # ── Payment table ──
            st.markdown('<div class="tbl-label">Final Payment Table</div>', unsafe_allow_html=True)
            display_df = final_df.drop(columns=["MATCH_SCORE"])
            st.dataframe(display_df, use_container_width=True, hide_index=True)

            # ── Match quality ──
            with st.expander("Match quality report", expanded=False):
                mq = df_agg[["DRIVER_NAME","MATCHED_NAME","MATCH_SCORE","AMOUNT"]].copy()
                mq["STATUS"] = mq["MATCH_SCORE"].apply(
                    lambda s: "Exact (100)" if s == 100 else (f"Fuzzy ({s})" if s >= 80 else "No match")
                )
                st.dataframe(mq, use_container_width=True, hide_index=True)

            # ── Downloads ──
            st.markdown("<hr class='rule'/>", unsafe_allow_html=True)
            st.markdown(f"""
            <div class="section-head" style="margin-top:0">
                <div class="section-num done">✓</div>
                <div>
                    <div class="section-label">Export</div>
                    <div class="section-hint">Download payment reports for records and bank upload</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            col1, col2 = st.columns(2, gap="medium")

            with col1:
                st.markdown("""
                <div class="dl-card">
                    <div class="dl-card-title">Full Payment Report</div>
                    <div class="dl-card-desc">All drivers · amounts · account details · match scores</div>
                </div>
                """, unsafe_allow_html=True)
                excel_full = styled_excel(display_df, "Payment Report")
                st.download_button(
                    label="Download Payment Report",
                    data=excel_full,
                    file_name="opti360_payment_report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

            with col2:
                st.markdown("""
                <div class="dl-card">
                    <div class="dl-card-title">Bank Transfer Sheet</div>
                    <div class="dl-card-desc">Matched drivers only · ready for bank portal upload</div>
                </div>
                """, unsafe_allow_html=True)
                bank_df   = final_df[["ACCOUNT_NAME","ACCOUNT_NO","BANK","AMOUNT"]].dropna()
                excel_bank = styled_excel(bank_df, "Bank Transfer")
                st.download_button(
                    label="Download Bank Transfer Sheet",
                    data=excel_bank,
                    file_name="opti360_bank_transfer.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

# ══════════════════════════════════════════════
# FOOTER
# ══════════════════════════════════════════════
st.markdown("""
<div class="footer">
    <span>Opti360</span> · Driver's Allowance Machine · Fleet Finance Operations
</div>
""", unsafe_allow_html=True)