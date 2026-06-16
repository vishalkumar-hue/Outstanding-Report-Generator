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
# DARK THEME CSS
# ─────────────────────────────────────────────

st.markdown("""
<style>
    /* ── Base dark background ── */
    .stApp {
        background: #0f1117 !important;
    }
    section[data-testid="stSidebar"] {
        background: #1a1d27 !important;
    }

    /* ── All text light ── */
    html, body, [class*="css"], .stMarkdown, p, label, span {
        color: #e0e6f0 !important;
    }

    /* ── File uploader ── */
    [data-testid="stFileUploader"] {
        background: #1a1d27 !important;
        border: 1.5px dashed #2e75b6 !important;
        border-radius: 10px !important;
        padding: 10px !important;
    }
    [data-testid="stFileUploader"] * { color: #a0b4cc !important; }

    /* ── Divider ── */
    hr { border-color: #2a2d3a !important; }

    /* ── Generate button ── */
    .stButton > button {
        background: linear-gradient(135deg, #1a3a5c, #2e75b6) !important;
        color: #ffffff !important;
        font-weight: 700 !important;
        border-radius: 8px !important;
        padding: 0.6rem 2rem !important;
        border: none !important;
        font-size: 1rem !important;
        letter-spacing: 0.3px !important;
        box-shadow: 0 4px 15px rgba(46,117,182,0.35) !important;
        transition: all 0.2s ease !important;
    }
    .stButton > button:hover {
        background: linear-gradient(135deg, #2e75b6, #41a0e0) !important;
        box-shadow: 0 6px 20px rgba(46,117,182,0.5) !important;
        transform: translateY(-1px) !important;
    }

    /* ── Download button ── */
    [data-testid="stDownloadButton"] > button {
        background: linear-gradient(135deg, #1a5c3a, #27a361) !important;
        color: #ffffff !important;
        font-weight: 700 !important;
        border-radius: 8px !important;
        padding: 0.6rem 2rem !important;
        border: none !important;
        font-size: 1rem !important;
        box-shadow: 0 4px 15px rgba(39,163,97,0.35) !important;
        transition: all 0.2s ease !important;
    }
    [data-testid="stDownloadButton"] > button:hover {
        background: linear-gradient(135deg, #27a361, #2ecc71) !important;
        transform: translateY(-1px) !important;
    }

    /* ── Toggle ── */
    [data-testid="stToggle"] span { color: #a0b4cc !important; }

    /* ── Metric cards ── */
    .metric-card {
        background: #1a1d27;
        border-radius: 12px;
        padding: 16px 20px;
        margin: 6px 0;
        border-left: 4px solid #2e75b6;
        box-shadow: 0 2px 8px rgba(0,0,0,0.3);
    }
    .metric-card.green  { border-left-color: #27a361; background: #0d1f15; }
    .metric-card.red    { border-left-color: #e05c6a; background: #1f0d10; }
    .metric-card.yellow { border-left-color: #f0b429; background: #1f1a0d; }
    .metric-card.blue   { border-left-color: #2e75b6; background: #0d1520; }

    .metric-label { font-size: 0.75rem; color: #7a8a9a !important; letter-spacing: 0.5px; text-transform: uppercase; }
    .metric-value { font-size: 1.4rem; font-weight: 800; margin-top: 2px; }
    .metric-value.green  { color: #27a361 !important; }
    .metric-value.red    { color: #e05c6a !important; }
    .metric-value.yellow { color: #f0b429 !important; }
    .metric-value.blue   { color: #4da3e8 !important; }
    .metric-value.white  { color: #e0e6f0 !important; }

    /* ── Badge ── */
    .badge {
        display: inline-block;
        padding: 4px 12px;
        border-radius: 20px;
        font-size: 0.8rem;
        font-weight: 700;
        letter-spacing: 0.3px;
    }
    .badge-green  { background: #0d3320; color: #27a361; border: 1px solid #27a361; }
    .badge-red    { background: #3a0d12; color: #e05c6a; border: 1px solid #e05c6a; }

    /* ── Section headers ── */
    .section-header {
        font-size: 0.85rem;
        font-weight: 700;
        color: #4da3e8 !important;
        margin: 1.6rem 0 0.6rem 0;
        letter-spacing: 1px;
        text-transform: uppercase;
        border-bottom: 1px solid #1e2a3a;
        padding-bottom: 6px;
    }

    /* ── App header box ── */
    .app-header {
        background: linear-gradient(135deg, #0d1a2e, #112240);
        border-radius: 14px;
        padding: 24px 28px;
        margin-bottom: 8px;
        border: 1px solid #1e3a5c;
        box-shadow: 0 4px 20px rgba(0,0,0,0.4);
    }
    .app-title {
        font-size: 1.7rem;
        font-weight: 800;
        color: #e0e6f0 !important;
        margin: 0;
        letter-spacing: -0.3px;
    }
    .app-subtitle {
        font-size: 0.9rem;
        color: #6a8aaa !important;
        margin-top: 4px;
    }

    /* ── Success / error / info boxes ── */
    [data-testid="stAlert"] {
        border-radius: 8px !important;
        border: none !important;
    }

    /* ── Streamlit metric widget ── */
    [data-testid="metric-container"] {
        background: #1a1d27 !important;
        border-radius: 10px !important;
        padding: 12px 16px !important;
        border: 1px solid #2a2d3a !important;
    }
    [data-testid="metric-container"] label { color: #7a8a9a !important; }
    [data-testid="metric-container"] [data-testid="stMetricValue"] { color: #e0e6f0 !important; }

    /* ── Selectbox ── */
    [data-testid="stSelectbox"] > div > div {
        background: #1a1d27 !important;
        border-color: #2e75b6 !important;
        color: #e0e6f0 !important;
        border-radius: 8px !important;
    }

    /* ── Spinner ── */
    .stSpinner { color: #4da3e8 !important; }

    /* ── Code / caption ── */
    code { background: #1e2535 !important; color: #4da3e8 !important; border-radius: 4px; padding: 1px 5px; }
    .stCaption { color: #4a5a6a !important; }

    /* ── Sheets tag row ── */
    .sheet-tag {
        display: inline-block;
        background: #1a2535;
        border: 1px solid #2e3a50;
        color: #4da3e8 !important;
        border-radius: 6px;
        padding: 3px 10px;
        font-size: 0.8rem;
        font-weight: 600;
        margin: 2px 3px;
    }

    footer { visibility: hidden; }
    #MainMenu { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────

st.markdown("""
<div class="app-header">
  <div class="app-title">📊 Outstanding Report Generator</div>
  <div class="app-subtitle">Tally Ledger Excel → Party-wise &amp; Project-wise Outstanding Report &nbsp;|&nbsp; Innovatiview India</div>
</div>
""", unsafe_allow_html=True)

st.markdown("")

# ─────────────────────────────────────────────
# STEP 1 — TALLY FILE
# ─────────────────────────────────────────────

st.markdown('<div class="section-header">① Tally Ledger File</div>', unsafe_allow_html=True)
tally_file = st.file_uploader(
    "Tally Excel file (.xlsx) upload karo",
    type=["xlsx"],
    key="tally_upload",
)
if tally_file:
    st.success(f"✅ **{tally_file.name}** loaded ({tally_file.size // 1024} KB)")

st.divider()

# ─────────────────────────────────────────────
# STEP 2 — MASTER CSV (OPTIONAL TOGGLE)
# ─────────────────────────────────────────────

st.markdown('<div class="section-header">② Master CSV Matching &nbsp;<span style="color:#4a5a6a;font-size:0.75rem;font-weight:400;text-transform:none;letter-spacing:0;">(Optional)</span></div>', unsafe_allow_html=True)

include_csv    = st.toggle("🔗 Master CSV se MATCHING tab banana hai?", value=False)
csv_file       = None
csv_col_manual = None

if include_csv:
    csv_file = st.file_uploader(
        "Master CSV file upload karo",
        type=["csv"],
        key="csv_upload",
    )
    if csv_file:
        st.success(f"✅ **{csv_file.name}** loaded ({csv_file.size // 1024} KB)")
        try:
            csv_bytes  = csv_file.read(); csv_file.seek(0)
            _, _, _, fieldnames = parse_master_csv(csv_bytes)
            st.info(f"📋 Columns: `{'` | `'.join(fieldnames)}`")

            from core import find_column, PROJECT_CODE_CANDIDATES
            auto_col = find_column(fieldnames, PROJECT_CODE_CANDIDATES)
            if auto_col:
                st.markdown(f"**Project Code column detected:** `{auto_col}` ✅")
            else:
                st.warning("⚠️ Project Code column auto-detect nahi hua — manually select karo:")
                csv_col_manual = st.selectbox(
                    "Project Code column:",
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
            tally_bytes     = tally_bytes,
            tally_filename  = tally_file.name,
            csv_bytes       = csv_bytes,
            csv_filename    = csv_fname,
            csv_project_col = csv_col_manual,
        )

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

    st.markdown('<div class="section-header">📈 Summary</div>', unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        st.markdown(f"""
        <div class="metric-card blue">
          <div class="metric-label">Opening Balance</div>
          <div class="metric-value white">₹{summary['opening']:,.2f}</div>
        </div>
        <div class="metric-card green">
          <div class="metric-label">Total Invoiced</div>
          <div class="metric-value green">₹{summary['total_invoiced']:,.2f}</div>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown(f"""
        <div class="metric-card green">
          <div class="metric-label">Total Paid</div>
          <div class="metric-value green">₹{summary['total_paid']:,.2f}</div>
        </div>
        <div class="metric-card yellow">
          <div class="metric-label">TDS Deducted</div>
          <div class="metric-value yellow">₹{summary['total_tds']:,.2f}</div>
        </div>
        """, unsafe_allow_html=True)

    outstanding = summary['outstanding']
    if outstanding <= 0:
        card_cls   = "green"
        badge_cls  = "badge-green"
        badge_text = "✅ CLEAR"
        val_cls    = "green"
        msg        = "Party ka koi outstanding nahi hai"
    else:
        card_cls   = "red"
        badge_cls  = "badge-red"
        badge_text = "⚠️ OUTSTANDING"
        val_cls    = "red"
        msg        = f"Receivable from party"

    st.markdown(f"""
    <div class="metric-card {card_cls}" style="margin-top:12px;padding:20px 24px;">
      <div style="display:flex;justify-content:space-between;align-items:center;">
        <div>
          <div class="metric-label">Net Outstanding</div>
          <div class="metric-value {val_cls}" style="font-size:1.8rem;">₹{outstanding:,.2f}</div>
          <div style="font-size:0.82rem;color:#6a7a8a;margin-top:3px;">{msg}</div>
        </div>
        <span class="badge {badge_cls}">{badge_text}</span>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # Match counts
    if match_counts:
        st.markdown('<div class="section-header">🔗 Matching Summary</div>', unsafe_allow_html=True)
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("✅ Matched",     match_counts.get('matched', 0))
        m2.metric("⚠️ Mismatch",    match_counts.get('mismatch', 0))
        m3.metric("❌ Not in CSV",   match_counts.get('not_csv', 0))
        m4.metric("❌ Not in Tally", match_counts.get('not_tally', 0))

    # Sheets info
    sheets = ["SUMMARY", "LEDGER", "INVOICE vs PAYMENT", "PROJECT WISE"]
    if include_csv and csv_file:
        sheets.append("MATCHING")
    tags = "".join(f'<span class="sheet-tag">{s}</span>' for s in sheets)
    st.markdown(
        f'<div style="margin:10px 0 4px 0;font-size:0.8rem;color:#4a5a6a;">Sheets in report:</div>{tags}',
        unsafe_allow_html=True,
    )

    st.divider()

    out_name = f"Outstanding_{tally_file.name.rsplit('.',1)[0]}_{datetime.now().strftime('%d%b%Y')}.xlsx"
    st.download_button(
        label               = "⬇️  Download Excel Report",
        data                = excel_bytes,
        file_name           = out_name,
        mime                = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width = True,
    )
    st.caption(f"📄 {out_name}")

# ─────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────

st.markdown("<br>", unsafe_allow_html=True)
st.markdown(
    "<p style='text-align:center;color:#2a3a4a;font-size:0.78rem;'>"
    "Innovatiview India &nbsp;·&nbsp; Outstanding Report Generator v4"
    "</p>",
    unsafe_allow_html=True,
)
