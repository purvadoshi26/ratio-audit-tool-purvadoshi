"""
Audit Analytical Procedures Tool
Built by Purva Doshi
"""

import streamlit as st
import pandas as pd
import tempfile, os

st.set_page_config(
    page_title="Audit AP Tool — Purva Doshi",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap');

html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }

.stApp { background: #F7F8FA; }

[data-testid="stSidebar"] {
    background: #0D1B2A !important;
    border-right: 1px solid #1E3048;
}
[data-testid="stSidebar"] * { color: #CBD5E1 !important; }
[data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3 { color: #F1F5F9 !important; }
[data-testid="stSidebar"] .stSelectbox > div > div {
    background: #1E3048 !important;
    border: 1px solid #334155 !important;
    color: #F1F5F9 !important;
}
[data-testid="stSidebar"] .stTextInput > div > div > input {
    background: #1E3048 !important;
    border: 1px solid #334155 !important;
    color: #F1F5F9 !important;
}

.metric-card {
    background: white;
    border: 1px solid #E2E8F0;
    border-radius: 10px;
    padding: 18px 20px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.06);
}
.metric-label { font-size: 11px; font-weight: 600; color: #64748B; text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 4px; }
.metric-value { font-size: 22px; font-weight: 600; color: #0D1B2A; font-family: 'DM Mono', monospace; }
.metric-delta { font-size: 12px; margin-top: 2px; }
.delta-pos { color: #16A34A; }
.delta-neg { color: #DC2626; }

.section-header {
    font-size: 13px; font-weight: 600; color: #64748B;
    text-transform: uppercase; letter-spacing: 0.08em;
    border-bottom: 2px solid #E2E8F0;
    padding-bottom: 8px; margin: 24px 0 16px 0;
}

.risk-pill {
    display: inline-block; padding: 3px 10px;
    border-radius: 20px; font-size: 11px; font-weight: 600;
}
.risk-high { background: #FEE2E2; color: #991B1B; }
.risk-mod  { background: #FEF9C3; color: #854D0E; }
.risk-low  { background: #DCFCE7; color: #14532D; }
.risk-crit { background: #7F1D1D; color: white; }

.footer-bar {
    margin-top: 40px; padding: 16px 0;
    border-top: 1px solid #E2E8F0;
    display: flex; align-items: center; gap: 8px;
    font-size: 12px; color: #94A3B8;
}

.stDownloadButton > button {
    background: #0D1B2A !important; color: white !important;
    border: none !important; border-radius: 8px !important;
    font-weight: 600 !important; padding: 10px 24px !important;
    font-family: 'DM Sans', sans-serif !important;
}
.stDownloadButton > button:hover { background: #1E3A5F !important; }

.stButton > button {
    border-radius: 8px !important;
    font-family: 'DM Sans', sans-serif !important;
    font-weight: 500 !important;
}

div[data-testid="stExpander"] {
    border: 1px solid #E2E8F0 !important;
    border-radius: 10px !important;
    background: white !important;
}

.stAlert { border-radius: 8px !important; }

.hero-title {
    font-size: 26px; font-weight: 600; color: #0D1B2A;
    margin-bottom: 4px;
}
.hero-sub {
    font-size: 14px; color: #64748B; margin-bottom: 0;
}
.standard-badge {
    display: inline-block; background: #EFF6FF;
    color: #1D4ED8; font-size: 11px; font-weight: 600;
    padding: 3px 10px; border-radius: 4px;
    font-family: 'DM Mono', monospace;
}
</style>
""", unsafe_allow_html=True)


# ── Sidebar ───────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 📊 Audit AP Tool")
    st.markdown("---")

    st.markdown("**Engagement**")
    client_name = st.text_input("Client Name", placeholder="e.g. ABC Pvt Ltd")
    period      = st.text_input("Period", placeholder="e.g. FY 2024-25")

    st.markdown("---")
    st.markdown("**Reporting Standard**")
    standard = st.selectbox("Standard", ["Ind AS", "IFRS", "US GAAP"])

    st.markdown("---")
    st.markdown("**Materiality**")
    auto_mat = st.checkbox("Auto-detect threshold", value=True)
    if not auto_mat:
        mat_pct = st.slider("Variance threshold (%)", 5, 50, 20, 5)
    else:
        mat_pct = 20

    st.markdown("---")
    st.markdown("""
    <div style='font-size:11px; color:#64748B; line-height:1.8;'>
    <b style='color:#94A3B8'>Standards Applied</b><br>
    ISA/SA 520 · Analytical Procedures<br>
    ISA/SA 570 · Going Concern<br>
    ISA/SA 240 · Fraud (Benford's)<br>
    ISA/SA 315 · Risk Assessment
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("""
    <div style='font-size:11px; color:#475569;'>
    Built by <a href='https://www.linkedin.com/in/purvadoshi26/' target='_blank'
    style='color:#60A5FA; text-decoration:none; font-weight:600;'>Purva Doshi</a>
    </div>
    """, unsafe_allow_html=True)


# ── Header ────────────────────────────────────────────────────
col_h1, col_h2 = st.columns([3, 1])
with col_h1:
    name_display = f" — {client_name}" if client_name else ""
    period_display = f" · {period}" if period else ""
    st.markdown(f"<div class='hero-title'>Audit Analytical Procedures{name_display}</div>", unsafe_allow_html=True)
    st.markdown(f"<div class='hero-sub'><span class='standard-badge'>{standard}</span>&nbsp; ISA/SA 520 · 570 · 240{period_display}</div>", unsafe_allow_html=True)
with col_h2:
    st.markdown("")

st.markdown("---")

# ── Tabs ──────────────────────────────────────────────────────
tab_main, tab_bf, tab_help = st.tabs(["📋 Trial Balance Analysis", "🔍 Benford's Law", "❓ Format Guide"])


# ════════════════════════════════════════════════════════════════
# TAB 1 — MAIN ANALYSIS
# ════════════════════════════════════════════════════════════════
with tab_main:

    col_up, col_info = st.columns([3, 2])

    with col_up:
        st.markdown("<div class='section-header'>Upload Trial Balance</div>", unsafe_allow_html=True)
        uploaded = st.file_uploader(
            "Excel or CSV file",
            type=["xlsx", "xls", "csv"],
            help="Required columns: Account Name, CY Amount, PY Amount. See Format Guide tab.",
            label_visibility="collapsed"
        )

    with col_info:
        st.markdown("<div class='section-header'>Required Columns</div>", unsafe_allow_html=True)
        st.markdown("""
        | Column | Required |
        |---|---|
        | `Account Name` | ✅ Yes |
        | `CY Amount` | ✅ Yes |
        | `PY Amount` | ✅ Yes |
        | `Category` | Optional |
        """)
        st.caption("Amounts can be in ₹, Lakhs, or Crores — auto-detected. Category auto-classified if blank.")

    if uploaded is not None:
        from engine import load_tb, run_analysis, risk_summary

        # Save to temp
        suffix = ".csv" if uploaded.name.endswith(".csv") else ".xlsx"
        with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
            tmp.write(uploaded.read())
            fpath = tmp.name

        with st.spinner("Loading file..."):
            try:
                df, unit_info, tb_warnings = load_tb(fpath)
            except ValueError as e:
                st.error(f"❌ {e}")
                st.stop()

        unit_label = unit_info["label"]
        threshold_abs = unit_info["threshold"] if auto_mat else None

        st.success(f"✅ {len(df)} accounts loaded · Units: **{unit_label}** · Materiality: **{unit_info['threshold']:,.2f}**")

        # Show warnings
        for w in tb_warnings:
            st.warning(f"⚠️ {w}")

        # Warn unclassified
        unclass = df[df["Category"] == "unclassified"]
        if len(unclass):
            with st.expander(f"⚠️ {len(unclass)} accounts could not be auto-classified — click to review"):
                st.dataframe(unclass[["Account Name", "CY Amount", "PY Amount"]], use_container_width=True)
                st.caption("Add a Category column to your file or rename account names to match common keywords.")

        with st.spinner("Running analysis..."):
            results = run_analysis(df, unit_info, mat_pct, standard)

        agg    = results["agg"]
        ratios = results["ratios"]
        df_var = results["variance"]
        gc     = results["going_concern"]

        has_revenue  = agg["revenue"]["cy"] is not None and agg["revenue"]["cy"] != 0
        has_bs       = agg["total_assets"]["cy"] is not None and agg["total_assets"]["cy"] != 0
        has_profit   = agg["net_profit"]["cy"] is not None

        # ── KPI Cards ──────────────────────────────────────────
        st.markdown("<div class='section-header'>Key Metrics</div>", unsafe_allow_html=True)

        kpi_cols = st.columns(5)

        def fmt(v, dec=2):
            if v is None: return "—"
            try: return f"{float(v):,.{dec}f}"
            except: return str(v)

        def delta_html(cy, py):
            if py is None or py == 0 or cy is None: return ""
            pct = (cy - py) / abs(py) * 100
            cls = "delta-pos" if pct >= 0 else "delta-neg"
            arrow = "▲" if pct >= 0 else "▼"
            return f"<div class='metric-delta {cls}'>{arrow} {abs(pct):.1f}% YoY</div>"

        kpis = [
            ("Revenue", fmt(agg["revenue"]["cy"]), delta_html(agg["revenue"]["cy"], agg["revenue"]["py"]), has_revenue),
            ("Net Profit", fmt(agg["net_profit"]["cy"]), delta_html(agg["net_profit"]["cy"], agg["net_profit"]["py"]), has_profit),
            ("Total Assets", fmt(agg["total_assets"]["cy"]), delta_html(agg["total_assets"]["cy"], agg["total_assets"]["py"]), has_bs),
            ("Working Capital", fmt(agg["working_capital"]["cy"]), "", has_bs),
            ("Going Concern", gc["overall_risk"], "", True),
        ]

        gc_color_map = {"LOW": "#16A34A", "MODERATE": "#D97706", "HIGH": "#EA580C", "CRITICAL": "#DC2626"}
        for col, (label, value, dlt, available) in zip(kpi_cols, kpis):
            with col:
                if not available:
                    st.markdown(f"""<div class='metric-card'>
                        <div class='metric-label'>{label}</div>
                        <div class='metric-value' style='color:#94A3B8;font-size:14px;'>No data</div>
                    </div>""", unsafe_allow_html=True)
                elif label == "Going Concern":
                    color = gc_color_map.get(value, "#0D1B2A")
                    st.markdown(f"""<div class='metric-card'>
                        <div class='metric-label'>{label}</div>
                        <div class='metric-value' style='color:{color};font-size:18px;'>{value}</div>
                        <div class='metric-delta' style='color:{color};'>{gc['score']}/10 risk score</div>
                    </div>""", unsafe_allow_html=True)
                elif label == "Working Capital":
                    wc = agg["working_capital"]["cy"]
                    color = "#DC2626" if wc and wc < 0 else "#0D1B2A"
                    note = "<div class='metric-delta delta-neg'>⚠ Negative</div>" if wc and wc < 0 else ""
                    st.markdown(f"""<div class='metric-card'>
                        <div class='metric-label'>{label} ({unit_label})</div>
                        <div class='metric-value' style='color:{color};'>{value}</div>{note}
                    </div>""", unsafe_allow_html=True)
                else:
                    st.markdown(f"""<div class='metric-card'>
                        <div class='metric-label'>{label} ({unit_label})</div>
                        <div class='metric-value'>{value}</div>{dlt}
                    </div>""", unsafe_allow_html=True)

        # ── Analysis Tabs ──────────────────────────────────────
        st.markdown("<div class='section-header'>Analysis</div>", unsafe_allow_html=True)
        r1, r2, r3, r4 = st.tabs(["📐 Ratios", "📋 All Accounts", "🚩 Flagged Items", "🏥 Going Concern"])

        def style_flag(val):
            if "🔴" in str(val): return "background-color:#FEE2E2"
            if "🟡" in str(val): return "background-color:#FEF9C3"
            if "🟢" in str(val): return "background-color:#DCFCE7"
            return ""

        with r1:
            if not ratios:
                st.info("No ratios could be calculated with the available data.")
            else:
                ratio_df = pd.DataFrame(ratios)
                available_ratios = ratio_df[ratio_df["Available"] == True].copy()
                if available_ratios.empty:
                    st.info("No ratios available — check your data has revenue, assets, and liability figures.")
                else:
                    show_cols = ["Ratio", "Formula", "CY", "PY", "YoY (%)", "Flag", "Note"]
                    available_cols = [c for c in show_cols if c in available_ratios.columns]
                    disp = available_ratios[available_cols].copy()
                    for c in ["CY", "PY"]:
                        if c in disp.columns:
                            disp[c] = disp[c].apply(lambda x: f"{x:.2f}" if isinstance(x, float) else x)
                    if "YoY (%)" in disp.columns:
                        disp["YoY (%)"] = disp["YoY (%)"].apply(lambda x: f"{x:+.1f}%" if isinstance(x, float) else x)
                    st.dataframe(
                        disp.style.applymap(style_flag, subset=["Flag"] if "Flag" in disp.columns else []),
                        use_container_width=True, height=460, hide_index=True
                    )
                    st.caption(f"🔴 HIGH risk  ·  🟡 MODERATE  ·  🟢 OK  ·  Standard: {standard}")

        with r2:
            disp_all = df_var[["Account Name", "Category", "CY Amount", "PY Amount", "Change", "Change (%)", "Flag"]].copy()
            for c in ["CY Amount", "PY Amount", "Change"]:
                disp_all[c] = disp_all[c].apply(lambda x: f"{x:,.2f}" if isinstance(x, (int, float)) else x)
            disp_all["Change (%)"] = disp_all["Change (%)"].apply(
                lambda x: "New" if x == 999.0 else (f"{x:+.1f}%" if isinstance(x, float) else x))
            st.dataframe(
                disp_all.style.applymap(style_flag, subset=["Flag"]),
                use_container_width=True, height=500, hide_index=True
            )

        with r3:
            flagged = df_var[df_var["Flag"].str.contains("🔴|🟡", na=False)]
            if flagged.empty:
                st.success("✅ No flagged items — all accounts within materiality threshold.")
            else:
                high = flagged[flagged["Flag"].str.contains("🔴", na=False)]
                mod  = flagged[flagged["Flag"].str.contains("🟡", na=False)]
                col_f1, col_f2 = st.columns(2)
                with col_f1:
                    st.markdown(f"🔴 **{len(high)} HIGH risk** — immediate attention required")
                with col_f2:
                    st.markdown(f"🟡 **{len(mod)} MODERATE** — obtain explanation")
                disp_f = flagged[["Account Name", "Category", "CY Amount", "PY Amount", "Change", "Change (%)", "Flag"]].copy()
                for c in ["CY Amount", "PY Amount", "Change"]:
                    disp_f[c] = disp_f[c].apply(lambda x: f"{x:,.2f}" if isinstance(x, (int, float)) else x)
                disp_f["Change (%)"] = disp_f["Change (%)"].apply(
                    lambda x: "New" if x == 999.0 else (f"{x:+.1f}%" if isinstance(x, float) else x))
                st.dataframe(
                    disp_f.style.applymap(style_flag, subset=["Flag"]),
                    use_container_width=True, height=400, hide_index=True
                )

        with r4:
            gc_colors = {"CRITICAL": "#DC2626", "HIGH": "#EA580C", "MODERATE": "#D97706", "LOW": "#16A34A"}
            bg = gc_colors.get(gc["overall_risk"], "#0D1B2A")
            st.markdown(f"""
            <div style='background:{bg};color:white;padding:16px 20px;border-radius:10px;
                        font-size:17px;font-weight:600;margin-bottom:12px;'>
                {gc["overall_risk"]} RISK &nbsp;·&nbsp; Score: {gc["score"]}/10
            </div>""", unsafe_allow_html=True)
            st.markdown(f"> {gc['conclusion']}")
            gc_df = pd.DataFrame(gc["indicators"])
            if not gc_df.empty:
                st.dataframe(
                    gc_df.style.applymap(style_flag, subset=["Status"] if "Status" in gc_df.columns else []),
                    use_container_width=True, hide_index=True
                )

        # ── Generate Workpaper ─────────────────────────────────
        st.markdown("<div class='section-header'>Generate Workpaper</div>", unsafe_allow_html=True)
        bf_result = st.session_state.get("bf_result", {
            "sufficient": False,
            "message": "Benford's Law analysis not run. Upload transaction data in the Benford's Law tab.",
            "risk_flag": "⚪ Not Analysed",
        })

        if st.button("⬇️ Generate Excel Workpaper", type="primary"):
            from reporter import generate_report
            with st.spinner("Building workpaper..."):
                out = os.path.join(tempfile.gettempdir(), "AP_Workpaper.xlsx")
                generate_report(
                    df_variance=df_var,
                    ratios=ratios,
                    agg=agg,
                    gc=gc,
                    bf=bf_result,
                    output_path=out,
                    client_name=client_name or "Client",
                    period=period or "—",
                    standard=standard,
                    unit_label=unit_label,
                )
            with open(out, "rb") as f:
                client_slug = (client_name or "Client").replace(" ", "_")
                st.download_button(
                    f"⬇️ Download — {client_name or 'Client'} Workpaper",
                    data=f,
                    file_name=f"AP_Workpaper_{client_slug}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            st.success("✅ Workpaper ready! Sheets: Cover · Ratios · Variance · Flagged · Going Concern · Benford's · Notes")

    else:
        st.markdown("""
        <div style='background:white;border:1px solid #E2E8F0;border-radius:10px;
                    padding:40px;text-align:center;color:#94A3B8;margin-top:20px;'>
            <div style='font-size:32px;margin-bottom:8px;'>📂</div>
            <div style='font-size:15px;font-weight:500;color:#64748B;'>Upload a Trial Balance to begin</div>
            <div style='font-size:13px;margin-top:6px;'>See the Format Guide tab for the required input format</div>
        </div>
        """, unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════
# TAB 2 — BENFORD'S LAW
# ════════════════════════════════════════════════════════════════
with tab_bf:
    st.markdown("<div class='section-header'>Benford's Law — ISA/SA 240 Fraud Risk Indicator</div>", unsafe_allow_html=True)
    st.markdown("""
    In naturally occurring financial data, digit **1** leads ~30% of the time; digit **9** only ~4.6%.
    Deviations from this pattern may indicate manipulation, rounding, or fabricated entries.
    Upload individual transaction amounts (journal entries, invoices, payments) — not the trial balance.
    """)

    col_bf1, col_bf2 = st.columns([3, 2])
    with col_bf1:
        bf_file = st.file_uploader(
            "Transaction amounts file (.xlsx / .csv)",
            type=["xlsx", "xls", "csv"],
            key="bf_upload",
            label_visibility="collapsed"
        )
    with col_bf2:
        st.markdown("""
        **Input format:**
        Excel with one column named `Amount`

        **Good sources:**
        - Journal entries from Tally/SAP
        - Purchase invoices register
        - Payment vouchers ledger
        - GL transaction listing

        Minimum 20 rows. Best with 500+.
        """)

    if bf_file is not None:
        from engine import run_benford

        suffix_bf = ".csv" if bf_file.name.endswith(".csv") else ".xlsx"
        with tempfile.NamedTemporaryFile(suffix=suffix_bf, delete=False) as tmp_bf:
            tmp_bf.write(bf_file.read())
            bf_path = tmp_bf.name

        try:
            bf_df = pd.read_csv(bf_path) if suffix_bf == ".csv" else pd.read_excel(bf_path)
        except Exception as e:
            st.error(f"❌ Could not read file: {e}")
            st.stop()

        amt_col = next((c for c in bf_df.columns if "amount" in c.lower() or "value" in c.lower()), bf_df.columns[0])
        amounts = pd.to_numeric(bf_df[amt_col], errors="coerce").dropna().abs().tolist()
        amounts = [a for a in amounts if a > 0]

        with st.spinner("Running Benford's Law analysis..."):
            bf_result = run_benford(amounts, label=bf_file.name)

        st.session_state["bf_result"] = bf_result

        if not bf_result.get("sufficient"):
            st.warning(bf_result.get("message", "Insufficient data."))
        else:
            flag = bf_result.get("risk_flag", "")
            def style_flag_bf(val):
                if "🔴" in str(val): return "background-color:#FEE2E2"
                if "🟡" in str(val): return "background-color:#FEF9C3"
                if "🟢" in str(val): return "background-color:#DCFCE7"
                return ""
            st.markdown(f"### {flag}")
            st.markdown(f"> {bf_result.get('interpretation', '')}")
            st.markdown(f"**Chi-Square: {bf_result['chi_square']:.2f}** · Critical values: p<0.05 → 15.51 · p<0.01 → 20.09 (df=8)")
            digit_df = pd.DataFrame(bf_result["digit_data"])
            st.dataframe(
                digit_df.style.applymap(style_flag_bf, subset=["Deviation (pp)"] if "Deviation (pp)" in digit_df.columns else []),
                use_container_width=True, height=350, hide_index=True
            )
            st.caption("Deviation >5pp = 🔴 Investigate · 2–5pp = 🟡 Review · <2pp = 🟢 Within range")
    else:
        st.info("Upload a transaction amounts file above to run Benford's Law analysis. This is optional — the main workpaper will note if it was not run.")


# ════════════════════════════════════════════════════════════════
# TAB 3 — FORMAT GUIDE
# ════════════════════════════════════════════════════════════════
with tab_help:
    st.markdown("<div class='section-header'>Input Format — Trial Balance</div>", unsafe_allow_html=True)

    st.markdown("""
    Your Excel file must have these **exact column names** in row 1:

    | Column | Required | Notes |
    |---|---|---|
    | `Account Name` | ✅ | Account name as in Tally / SAP |
    | `CY Amount` | ✅ | Current year balance |
    | `PY Amount` | ✅ | Prior year comparative |
    | `Category` | Optional | Auto-detected if blank |

    **Important:**
    - Do **not** include subtotal / total rows (they are filtered automatically)
    - Negative amounts are fine (e.g. credit balances)
    - Amounts in any unit — ₹, Lakhs, Crores — the tool auto-detects
    - Column names must match exactly (case-insensitive)
    """)

    st.markdown("<div class='section-header'>Valid Category Values</div>", unsafe_allow_html=True)

    cat_data = {
        "Category Value": [
            "current_assets", "fixed_assets",
            "current_liabilities", "long_term_liabilities",
            "equity", "revenue", "cogs", "expenses", "interest_expense"
        ],
        "Use For": [
            "Assets due within 1 year",
            "Long-term tangible / intangible assets",
            "Liabilities due within 1 year",
            "Liabilities due after 1 year",
            "Owner's funds",
            "Revenue / top line",
            "Direct cost of goods sold",
            "Operating / indirect expenses",
            "Finance costs"
        ],
        "Examples": [
            "Cash, Debtors, Stock, TDS Receivable",
            "Plant & Machinery, Building, CWIP, Goodwill",
            "Creditors, GST Payable, OD, Advance from customers",
            "Term Loan, Deferred Tax Liability",
            "Share Capital, Reserves & Surplus",
            "Revenue from Operations, Other Income",
            "Purchases, Changes in Inventories, Direct Labour",
            "Salaries, Rent, Depreciation, Admin Expenses",
            "Interest on Loan, Bank Charges, Processing Fees"
        ]
    }
    st.dataframe(pd.DataFrame(cat_data), use_container_width=True, hide_index=True)

    st.markdown("<div class='section-header'>Benford's Law Input Format</div>", unsafe_allow_html=True)
    st.markdown("""
    Separate Excel file with one column named `Amount`.
    Each row = one individual transaction (not aggregated).
    Minimum 20 amounts. Works best with 500+.

    Good sources: journal entries, purchase invoices, payment register, GL transaction listing.
    """)

    st.markdown("<div class='section-header'>Standards Reference</div>", unsafe_allow_html=True)
    std_data = {
        "Procedure": ["Analytical Procedures", "Going Concern", "Fraud Risk", "Risk Assessment", "Revenue", "Impairment"],
        "Ind AS": ["SA 520", "SA 570", "SA 240", "SA 315", "Ind AS 115", "Ind AS 36"],
        "IFRS": ["ISA 520", "ISA 570", "ISA 240", "ISA 315", "IFRS 15", "IAS 36"],
        "US GAAP": ["AU-C 520", "AU-C 570", "AU-C 240", "AU-C 315", "ASC 606", "ASC 350"],
    }
    st.dataframe(pd.DataFrame(std_data), use_container_width=True, hide_index=True)


# ── Footer ────────────────────────────────────────────────────
st.markdown("""
<div class='footer-bar'>
    <span>Audit AP Tool</span>
    <span style='color:#CBD5E1'>·</span>
    <span>Built by <a href='https://www.linkedin.com/in/purvadoshi26/' target='_blank'
    style='color:#3B82F6;text-decoration:none;font-weight:600;'>Purva Doshi</a></span>
    <span style='color:#CBD5E1'>·</span>
    <span>ISA/SA 520 · 570 · 240 · 315</span>
</div>
""", unsafe_allow_html=True)
