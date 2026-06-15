"""
Outstanding Report Generator — Streamlit Web App
Innovatiview India
"""

import streamlit as st
from datetime import datetime
from core import generate_report, parse_master_csv

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────

st.set_page_config(
    page_title="Outstanding Report Generator",
    page_icon="📊",
    layout="centered",
)

# ─────────────────────────────────────────────
# CUSTOM CSS
# ─────────────────────────────────────────────

st.markdown("""
<style>
    .main { max-width: 780px; }
    .stButton > button {
        background-color: #1F4E79;
        color: white;
        font-weight: bold;
        border-radius: 6px;
        padding: 0.5rem 2rem;
        border: none;
        width: 100%;
        font-size: 1rem;
    }
    .stButton > button:hover { background-color: #2E75B6; }

    .metric-card {
        background: #f8f9fa;
        border-radius: 10px;
        padding: 16px 20px;
        margin: 6px 0;
        border-left: 5px solid #1F4E79;
    }
    .metric-card.green  { border-left-color: #375623; background: #e2efda; }
    .metric-card.red    { border-left-color: #9C0006; background: #ffc7ce; }
    .metric-card.yellow { border-left-color: #7d6608; background: #fff2cc; }

    .badge {
        display: inline-block;
        padding: 3px 10px;
        border-radius: 12px;
        font-size: 0.82rem;
        font-weight: 600;
    }
    .badge-green  { background: #C6EFCE; color: #375623; }
    .badge-red    { background: #FFC7CE; color: #9C0006; }
    .badge-yellow { background: #FFF2CC; color: #7d6608; }

    .section-header {
        font-size: 1.1rem;
        font-weight: 700;
        color: #1F4E79;
        margin: 1.4rem 0 0.5rem 0;
        border-bottom: 2px solid #DEEAF1;
        padding-bottom: 4px;
    }
    footer { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────

st.markdown("## 📊 Outstanding Report Generator")
st.markdown(
    "<p style='color:#555;margin-top:-10px;'>Tally Ledger Excel → Party-wise & Project-wise Outstanding Report</p>",
    unsafe_allow_html=True,
)
st.divider()

# ─────────────────────────────────────────────
# STEP 1 — TALLY FILE
# ─────────────────────────────────────────────

st.markdown('<div class="section-header">① Tally Ledger File Upload</div>', unsafe_allow_html=True)
tally_file = st.file_uploader(
    "Tally Excel file (.xlsx) upload karo",
    type=["xlsx"],
    key="tally_upload",
)
if tally_file:
    st.success(f"✅ File loaded: **{tally_file.name}** ({tally_file.size // 1024} KB)")

st.divider()

# ─────────────────────────────────────────────
# STEP 2 — MASTER CSV (OPTIONAL TOGGLE)
# ─────────────────────────────────────────────

st.markdown('<div class="section-header">② Master CSV Matching  <span style="color:#888;font-size:0.85rem;font-weight:400;">(Optional)</span></div>', unsafe_allow_html=True)

include_csv = st.toggle("🔗 Master CSV se MATCHING tab banana hai?", value=False)

csv_file       = None
csv_col_manual = None

if include_csv:
    csv_file = st.file_uploader(
        "Master CSV file upload karo",
        type=["csv"],
        key="csv_upload",
    )
    if csv_file:
        st.success(f"✅ CSV loaded: **{csv_file.name}** ({csv_file.size // 1024} KB)")

        # Column preview
        try:
            csv_bytes  = csv_file.read(); csv_file.seek(0)
            _, _, _, fieldnames = parse_master_csv(csv_bytes)
            st.info(f"📋 Available columns: `{'` | `'.join(fieldnames)}`")

            # Auto-detect attempt
            from core import find_column, PROJECT_CODE_CANDIDATES
            auto_col = find_column(fieldnames, PROJECT_CODE_CANDIDATES)
            if auto_col:
                st.markdown(f"**Project Code column auto-detected:** `{auto_col}` ✅")
            else:
                st.warning("⚠️ 'Project Code' column auto-detect nahi hua. Manually select karo:")
                csv_col_manual = st.selectbox(
                    "Project Code column select karo:",
                    options=fieldnames,
                    key="csv_col_select",
                )
        except Exception as e:
            st.error(f"CSV preview error: {e}")

st.divider()

# ─────────────────────────────────────────────
# GENERATE BUTTON
# ─────────────────────────────────────────────

generate_clicked = st.button("🚀 Generate Outstanding Report", use_container_width=True)

if generate_clicked:
    if not tally_file:
        st.error("❌ Pehle Tally Excel file upload karo!")
        st.stop()

    if include_csv and not csv_file:
        st.error("❌ CSV toggle ON hai lekin file upload nahi ki!")
        st.stop()

    with st.spinner("⏳ Processing... please wait"):
        tally_bytes = tally_file.read()
        csv_bytes   = None
        csv_fname   = None

        if include_csv and csv_file:
            csv_bytes = csv_file.read()
            csv_fname = csv_file.name

        excel_bytes, summary, match_counts, error = generate_report(
            tally_bytes    = tally_bytes,
            tally_filename = tally_file.name,
            csv_bytes      = csv_bytes,
            csv_filename   = csv_fname,
            csv_project_col= csv_col_manual,
        )

    # ── CSV Column Error → let user pick column ──
    if error and error.startswith("CSV_COL_ERROR"):
        st.error("⚠️ " + error.split("|", 1)[1])
        st.info("Upar CSV section me manually column select karo, phir dobara Generate karo.")
        st.stop()

    if error:
        st.error(f"❌ Error: {error}")
        st.stop()

    # ─────────────────────────────────────────────
    # RESULTS
    # ─────────────────────────────────────────────

    st.success("✅ Report ready!")
    st.markdown('<div class="section-header">📈 Summary</div>', unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        st.markdown(f"""
        <div class="metric-card">
          <div style="font-size:0.8rem;color:#555;">Opening Balance</div>
          <div style="font-size:1.3rem;font-weight:700;">₹{summary['opening']:,.2f}</div>
        </div>
        <div class="metric-card green">
          <div style="font-size:0.8rem;color:#375623;">Total Invoiced</div>
          <div style="font-size:1.3rem;font-weight:700;color:#375623;">₹{summary['total_invoiced']:,.2f}</div>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown(f"""
        <div class="metric-card green">
          <div style="font-size:0.8rem;color:#375623;">Total Paid</div>
          <div style="font-size:1.3rem;font-weight:700;color:#375623;">₹{summary['total_paid']:,.2f}</div>
        </div>
        <div class="metric-card yellow">
          <div style="font-size:0.8rem;color:#7d6608;">TDS Deducted</div>
          <div style="font-size:1.3rem;font-weight:700;color:#7d6608;">₹{summary['total_tds']:,.2f}</div>
        </div>
        """, unsafe_allow_html=True)

    outstanding = summary['outstanding']
    if outstanding <= 0:
        badge_cls  = "badge-green"
        badge_text = "✅ CLEAR"
        card_cls   = "green"
        msg        = "Party ka koi outstanding nahi hai"
    else:
        badge_cls  = "badge-red"
        badge_text = "⚠️ OUTSTANDING"
        card_cls   = "red"
        msg        = f"₹{outstanding:,.2f} receivable from party"

    st.markdown(f"""
    <div class="metric-card {card_cls}" style="margin-top:10px;">
      <div style="display:flex;justify-content:space-between;align-items:center;">
        <div>
          <div style="font-size:0.8rem;color:#555;">NET OUTSTANDING</div>
          <div style="font-size:1.6rem;font-weight:800;">₹{outstanding:,.2f}</div>
          <div style="font-size:0.85rem;margin-top:2px;">{msg}</div>
        </div>
        <span class="badge {badge_cls}">{badge_text}</span>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # Match counts (if CSV used)
    if match_counts:
        st.markdown('<div class="section-header">🔗 MATCHING Summary</div>', unsafe_allow_html=True)
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("✅ Matched",       match_counts.get('matched', 0))
        m2.metric("⚠️ Mismatch",      match_counts.get('mismatch', 0))
        m3.metric("❌ Not in CSV",     match_counts.get('not_csv', 0))
        m4.metric("❌ Not in Tally",   match_counts.get('not_tally', 0))

    # ── SHEETS INFO ──
    sheets = ["SUMMARY", "LEDGER", "INVOICE vs PAYMENT", "PROJECT WISE"]
    if include_csv and csv_file:
        sheets.append("MATCHING")
    st.markdown(
        f"**Sheets in report:** " + "  •  ".join(f"`{s}`" for s in sheets),
        unsafe_allow_html=False,
    )

    # ── DOWNLOAD ──
    st.divider()
    out_name = f"Outstanding_{tally_file.name.rsplit('.',1)[0]}_{datetime.now().strftime('%d%b%Y')}.xlsx"
    st.download_button(
        label     = "⬇️ Download Excel Report",
        data      = excel_bytes,
        file_name = out_name,
        mime      = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width = True,
    )
    st.caption(f"File: {out_name}")

# ─────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────

st.divider()
st.markdown(
    "<p style='text-align:center;color:#aaa;font-size:0.8rem;'>"
    "Innovatiview India &nbsp;|&nbsp; Outstanding Report Generator v4"
    "</p>",
    unsafe_allow_html=True,
)
