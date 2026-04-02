import streamlit as st
import pandas as pd
import sqlite3
from io import BytesIO
from rapidfuzz import process, fuzz
import base64, os
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

# ═══════════════════════════════════════════════
# PAGE CONFIG
# ═══════════════════════════════════════════════
st.set_page_config(
    page_title="Opti360 Driver Payment System",
    page_icon="💳",
    layout="wide"
)

# ═══════════════════════════════════════════════
# REMOVE STREAMLIT UI (PROFESSIONAL LOOK)
# ═══════════════════════════════════════════════
hide_streamlit_style = """
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
[data-testid="stSidebar"] {display: none;}
[data-testid="stToolbar"] {display: none;}
.block-container {padding-top: 1rem;}
</style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# ═══════════════════════════════════════════════
# LOGO HANDLER
# ═══════════════════════════════════════════════
def _img_to_b64(path):
    if not os.path.exists(path):
        return ""
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode()

HERE = os.path.dirname(os.path.abspath(__file__))
b64_opti = _img_to_b64(os.path.join(HERE, "logo_opti360.png"))

# ═══════════════════════════════════════════════
# UI DESIGN
# ═══════════════════════════════════════════════
st.markdown(f"""
<style>
body {{background:#F4F6F9;font-family:Inter;}}

.header {{
    display:flex;justify-content:space-between;align-items:center;
    background:linear-gradient(135deg,#0D2137,#2471A3);
    padding:18px;border-radius:0 0 15px 15px;margin-bottom:25px;
}}
.header img {{height:50px;}}
.title {{color:white;font-size:1.6rem;font-weight:700;}}

.card {{
    background:white;padding:20px;border-radius:12px;
    box-shadow:0 2px 10px rgba(0,0,0,0.08);
    margin-bottom:20px;
}}

.stat {{
    background:#2471A3;color:white;padding:12px 20px;
    border-radius:10px;text-align:center;font-weight:600;
}}

.stButton>button {{
    background:#1B3A5C;color:white;font-weight:700;
    border-radius:10px;padding:12px;
}}

[data-testid="stDownloadButton"]>button {{
    background:#1A5F3C;color:white;font-weight:700;
    border-radius:10px;padding:12px;
}}
</style>

<div class="header">
    <div style="display:flex;align-items:center;gap:10px;">
        <img src="data:image/png;base64,{b64_opti}">
        <div class="title">Driver Payment Automation System</div>
    </div>
    <div style="color:white;">💳 Professional</div>
</div>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════
# DATABASE
# ═══════════════════════════════════════════════
conn = sqlite3.connect("drivers.db", check_same_thread=False)

# ═══════════════════════════════════════════════
# FILE READER
# ═══════════════════════════════════════════════
def read_file(file):
    name = file.name.lower()
    if name.endswith(("xlsx","xls","xlsm")):
        return pd.read_excel(file)
    elif name.endswith("csv"):
        return pd.read_csv(file)
    elif name.endswith("ods"):
        return pd.read_excel(file, engine="odf")
    else:
        st.error("Unsupported file format")
        return None

# ═══════════════════════════════════════════════
# EXCEL FORMATTER
# ═══════════════════════════════════════════════
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Payments")
        ws = writer.sheets["Payments"]

        for col in ws.columns:
            ws.column_dimensions[get_column_letter(col[0].column)].width = 22

        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1B3A5C")

    return output.getvalue()

# ═══════════════════════════════════════════════
# UPLOAD SECTION
# ═══════════════════════════════════════════════
col1, col2 = st.columns(2)

with col1:
    st.markdown('<div class="card"><b>Upload Driver Database</b></div>', unsafe_allow_html=True)
    driver_file = st.file_uploader("Driver Database", type=["xlsx","xls","csv","ods"])

with col2:
    st.markdown('<div class="card"><b>Upload Driver Report</b></div>', unsafe_allow_html=True)
    report_file = st.file_uploader("Driver Report", type=["xlsx","xls","csv","ods"])

# ═══════════════════════════════════════════════
# LOAD DRIVER DATABASE
# ═══════════════════════════════════════════════
if driver_file:
    df = read_file(driver_file)

    if df is not None:
        df.columns = df.columns.str.upper().str.strip()

        df.rename(columns={
            "FMS DRIVER'S NAME":"DRIVER_NAME",
            "ACCOUNT NAME":"ACCOUNT_NAME",
            "ACCOUNT NO":"ACCOUNT_NO"
        }, inplace=True)

        df["DRIVER_NAME"] = df["DRIVER_NAME"].astype(str).str.upper().str.strip()
        df["ACCOUNT_NO"] = df["ACCOUNT_NO"].astype(str).str.replace(r"\D","", regex=True).str.zfill(10)

        df.to_sql("drivers", conn, if_exists="replace", index=False)
        st.success("✅ Driver database loaded successfully")

# ═══════════════════════════════════════════════
# PROCESS REPORT
# ═══════════════════════════════════════════════
if report_file:
    df_r = read_file(report_file)

    if df_r is not None:
        df_r.columns = df_r.columns.str.upper().str.strip()

        df_r.rename(columns={
            "DRIVER NAME":"DRIVER_NAME",
            "TOTAL AMOUNT":"AMOUNT"
        }, inplace=True)

        df_r["DRIVER_NAME"] = df_r["DRIVER_NAME"].astype(str).str.upper().str.strip()
        df_r["AMOUNT"] = pd.to_numeric(df_r["AMOUNT"], errors="coerce").fillna(0)

        df_r = df_r.groupby("DRIVER_NAME", as_index=False)["AMOUNT"].sum()

        df_db = pd.read_sql("SELECT * FROM drivers", conn)

        # FUZZY MATCH
        matches = []
        for name in df_r["DRIVER_NAME"]:
            match = process.extractOne(name, df_db["DRIVER_NAME"], scorer=fuzz.ratio)
            matches.append(match[0] if match and match[1] > 80 else None)

        df_r["MATCH"] = matches

        final = pd.merge(df_r, df_db, left_on="MATCH", right_on="DRIVER_NAME", how="left")

        final["ACCOUNT_NO"] = final["ACCOUNT_NO"].astype(str).str.zfill(10)
        final["S/N"] = range(1, len(final)+1)

        final = final[["S/N","DRIVER_NAME_x","AMOUNT","ACCOUNT_NAME","ACCOUNT_NO","BANK"]]
        final.columns = ["S/N","DRIVER NAME","AMOUNT","ACCOUNT NAME","ACCOUNT NO","BANK"]

        # DISPLAY
        st.markdown('<div class="card"><b>Final Payment Table</b></div>', unsafe_allow_html=True)
        st.dataframe(final, use_container_width=True)

        # STATS
        total_amount = final["AMOUNT"].sum()

        st.markdown(f"""
        <div style="display:flex;gap:15px;">
            <div class="stat">{len(final)} Drivers</div>
            <div class="stat">₦{total_amount:,.0f}</div>
        </div>
        """, unsafe_allow_html=True)

        # DOWNLOADS
        st.download_button(
            "📥 Download Payment Report",
            to_excel(final),
            file_name="Driver_Payments.xlsx"
        )

        bank = final[["ACCOUNT NAME","ACCOUNT NO","BANK","AMOUNT"]].dropna()

        st.download_button(
            "🏦 Download Bank Sheet",
            to_excel(bank),
            file_name="Bank_Payments.xlsx"
        )

# ═══════════════════════════════════════════════
# FOOTER
# ═══════════════════════════════════════════════
st.markdown("""
<div style="margin-top:40px;padding:20px;background:#1B3A5C;color:white;border-radius:10px;text-align:center;">
Opti360 Driver Payment System | Powered by Crismel Solutions
</div>
""", unsafe_allow_html=True)