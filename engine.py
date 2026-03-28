"""
engine.py — Financial Analysis Engine
Supports Ind AS / IFRS / US GAAP
"""

import pandas as pd
import math

# ── Standard-aware references ─────────────────────────────────
STD_REFS = {
    "Ind AS": {
        "revenue":    "Ind AS 115",
        "impairment": "Ind AS 36",
        "leases":     "Ind AS 116",
        "ap":         "SA 520",
        "gc":         "SA 570",
        "fraud":      "SA 240",
        "risk":       "SA 315",
    },
    "IFRS": {
        "revenue":    "IFRS 15",
        "impairment": "IAS 36",
        "leases":     "IFRS 16",
        "ap":         "ISA 520",
        "gc":         "ISA 570",
        "fraud":      "ISA 240",
        "risk":       "ISA 315",
    },
    "US GAAP": {
        "revenue":    "ASC 606",
        "impairment": "ASC 350",
        "leases":     "ASC 842",
        "ap":         "AU-C 520",
        "gc":         "AU-C 570",
        "fraud":      "AU-C 240",
        "risk":       "AU-C 315",
    },
}

SUBTOTAL_KEYWORDS = [
    "total", "subtotal", "grand total", "net total", "sum",
    "total income", "total expense", "total assets", "total liabilities",
    "profit before", "profit after", "ebitda", "ebit",
]

CATEGORY_MAP = {
    "current_assets": [
        "cash", "bank", "debtor", "receivable", "inventory", "stock",
        "advance paid", "tds", "gst receivable", "prepaid", "short term investment",
        "other current", "loan given", "accrued income", "bills receivable",
    ],
    "fixed_assets": [
        "fixed asset", "plant", "machinery", "building", "equipment", "vehicle",
        "furniture", "computer", "cwip", "capital wip", "capital work",
        "intangible", "goodwill", "patent", "trademark", "right of use",
        "net block", "gross block", "accumulated depreciation",
    ],
    "current_liabilities": [
        "creditor", "payable", "gst payable", "tds payable", "advance received",
        "short term loan", "overdraft", "od limit", "bank od", "current portion",
        "outstanding expense", "provision", "statutory due",
    ],
    "long_term_liabilities": [
        "term loan", "long term loan", "debenture", "bond", "mortgage",
        "deferred tax liability", "long term provision", "security deposit received",
    ],
    "equity": [
        "share capital", "equity share", "preference share", "reserve",
        "surplus", "retained earning", "capital reserve", "general reserve",
        "securities premium", "other equity",
    ],
    "revenue": [
        "revenue", "sales", "turnover", "income from operation", "service income",
        "other income", "interest income", "dividend income", "rental income",
        "commission income", "export", "domestic sales",
    ],
    "cogs": [
        "purchase", "cost of goods", "cost of material", "raw material consumed",
        "direct material", "direct labour", "changes in inventor",
        "opening stock", "closing stock", "freight inward",
    ],
    "expenses": [
        "salary", "wage", "employee benefit", "staff expense", "depreciation",
        "amortisation", "rent", "electricity", "telephone", "printing",
        "stationary", "travelling", "conveyance", "advertisement", "marketing",
        "legal", "professional", "audit fee", "repair", "maintenance",
        "insurance", "other expense", "administrative", "selling expense",
        "distribution", "general expense",
    ],
    "interest_expense": [
        "finance cost", "interest expense", "interest on loan", "bank charge",
        "processing fee", "loan charge", "borrowing cost",
    ],
}


def _auto_classify(name: str) -> str:
    n = name.lower().strip()
    for cat, keywords in CATEGORY_MAP.items():
        for kw in keywords:
            if kw in n:
                return cat
    return "unclassified"


def _is_subtotal(name: str) -> bool:
    n = name.lower().strip()
    return any(kw in n for kw in SUBTOTAL_KEYWORDS)


def _clean_amount(val) -> float:
    if val is None: return 0.0
    s = str(val).strip()
    if s in ["-", "—", "", "nil", "n/a", "na"]: return 0.0
    s = s.replace(",", "").replace("₹", "").replace("$", "").strip()
    s = s.replace("(", "-").replace(")", "")
    try: return float(s)
    except: return 0.0


def load_tb(filepath: str):
    try:
        if filepath.endswith(".csv"):
            df = pd.read_csv(filepath)
        else:
            df = pd.read_excel(filepath)
    except Exception as e:
        raise ValueError(f"Cannot read file: {e}")

    df.columns = [str(c).strip() for c in df.columns]

    # Find columns flexibly
    def find_col(df, candidates):
        for cand in candidates:
            for col in df.columns:
                if cand.lower() in col.lower():
                    return col
        return None

    name_col = find_col(df, ["account name", "account", "particulars", "description", "head"])
    cy_col   = find_col(df, ["cy amount", "cy", "current year", "2024", "2025", "2023"])
    py_col   = find_col(df, ["py amount", "py", "prior year", "previous year", "2023", "2022", "2024"])
    cat_col  = find_col(df, ["category", "type", "classification", "group"])

    if name_col is None:
        raise ValueError("Cannot find 'Account Name' column. Please check the Format Guide for required column names.")
    if cy_col is None:
        raise ValueError("Cannot find 'CY Amount' column. Please check the Format Guide for required column names.")
    if py_col is None:
        raise ValueError("Cannot find 'PY Amount' column. Please check the Format Guide for required column names.")

    df = df.rename(columns={
        name_col: "Account Name",
        cy_col:   "CY Amount",
        py_col:   "PY Amount",
    })
    if cat_col:
        df = df.rename(columns={cat_col: "Category"})
    else:
        df["Category"] = ""

    # Drop subtotals and empty rows
    df = df[df["Account Name"].notna()].copy()
    df = df[~df["Account Name"].apply(_is_subtotal)].copy()
    df = df[df["Account Name"].str.strip() != ""].copy()

    df["CY Amount"] = df["CY Amount"].apply(_clean_amount)
    df["PY Amount"] = df["PY Amount"].apply(_clean_amount)

    # Auto-classify
    df["Category"] = df.apply(
        lambda r: r["Category"].strip().lower() if str(r["Category"]).strip() else _auto_classify(r["Account Name"]),
        axis=1
    )
    df["Category"] = df["Category"].apply(
        lambda c: c if c in CATEGORY_MAP else _auto_classify(c) if c not in CATEGORY_MAP else c
    )

    # Unit detection
    max_val = df[["CY Amount", "PY Amount"]].abs().max().max()
    if max_val >= 100:
        label, threshold = "₹ Crores", 2.0
    elif max_val >= 1:
        label, threshold = "₹ Lakhs", 5.0
    else:
        label, threshold = "₹", 50000.0

    unit_info = {"label": label, "threshold": threshold, "max_val": max_val}
    df = df.reset_index(drop=True)
    return df, unit_info


def _sum_cat(df, cats):
    if isinstance(cats, str): cats = [cats]
    subset = df[df["Category"].isin(cats)]
    return {"cy": subset["CY Amount"].sum(), "py": subset["PY Amount"].sum()}


def _safe_div(a, b, scale=1):
    if b is None or b == 0 or a is None: return None
    return (a / b) * scale


def _yoy(cy, py):
    if py is None or py == 0 or cy is None: return None
    return (cy - py) / abs(py) * 100


def aggregate(df):
    rev   = _sum_cat(df, "revenue")
    cogs  = _sum_cat(df, "cogs")
    exp   = _sum_cat(df, "expenses")
    int_e = _sum_cat(df, "interest_expense")
    ca    = _sum_cat(df, "current_assets")
    fa    = _sum_cat(df, "fixed_assets")
    cl    = _sum_cat(df, "current_liabilities")
    ltl   = _sum_cat(df, "long_term_liabilities")
    eq    = _sum_cat(df, "equity")

    gp = {"cy": rev["cy"] - cogs["cy"], "py": rev["py"] - cogs["py"]}
    op = {"cy": gp["cy"] - exp["cy"],   "py": gp["py"] - exp["py"]}
    np_ = {"cy": op["cy"] - int_e["cy"], "py": op["py"] - int_e["py"]}
    ta  = {"cy": ca["cy"] + fa["cy"],    "py": ca["py"] + fa["py"]}
    tl  = {"cy": cl["cy"] + ltl["cy"],  "py": cl["py"] + ltl["py"]}
    wc  = {"cy": ca["cy"] - cl["cy"],   "py": ca["py"] - cl["py"]}
    ebitda = {"cy": op["cy"] + exp["cy"] * 0,  "py": op["py"] + exp["py"] * 0}  # approximation

    return {
        "revenue": rev, "cogs": cogs, "expenses": exp, "interest_expense": int_e,
        "gross_profit": gp, "operating_profit": op, "net_profit": np_,
        "current_assets": ca, "fixed_assets": fa,
        "current_liabilities": cl, "long_term_liabilities": ltl,
        "equity": eq, "total_assets": ta, "total_liabilities": tl,
        "working_capital": wc,
    }


def _make_ratio(name, formula, cy, py, flag_thresholds, note, ref, available=True):
    if not available or cy is None:
        return {"Ratio": name, "Formula": formula, "CY": None, "PY": None,
                "YoY (%)": None, "Flag": "⚪ N/A", "Note": note, "Ref": ref, "Available": False}
    yoy = _yoy(cy, py)
    # Flag logic
    flag = "🟢 OK"
    for (op, val, lbl) in flag_thresholds:
        try:
            if op == "<" and cy < val:   flag = lbl; break
            if op == ">" and cy > val:   flag = lbl; break
            if op == "yoy>" and yoy is not None and abs(yoy) > val: flag = lbl; break
        except: pass
    return {
        "Ratio": name, "Formula": formula,
        "CY": round(cy, 2) if cy is not None else None,
        "PY": round(py, 2) if py is not None else None,
        "YoY (%)": round(yoy, 1) if yoy is not None else None,
        "Flag": flag, "Note": note, "Ref": ref, "Available": True,
    }


def calculate_ratios(agg, standard):
    ref = STD_REFS.get(standard, STD_REFS["IFRS"])
    rev_cy = agg["revenue"]["cy"];      rev_py = agg["revenue"]["py"]
    gp_cy  = agg["gross_profit"]["cy"]; gp_py  = agg["gross_profit"]["py"]
    np_cy  = agg["net_profit"]["cy"];   np_py  = agg["net_profit"]["py"]
    op_cy  = agg["operating_profit"]["cy"]; op_py = agg["operating_profit"]["py"]
    ca_cy  = agg["current_assets"]["cy"];   ca_py  = agg["current_assets"]["py"]
    cl_cy  = agg["current_liabilities"]["cy"]; cl_py = agg["current_liabilities"]["py"]
    ta_cy  = agg["total_assets"]["cy"]; ta_py  = agg["total_assets"]["py"]
    tl_cy  = agg["total_liabilities"]["cy"]; tl_py = agg["total_liabilities"]["py"]
    eq_cy  = agg["equity"]["cy"];       eq_py  = agg["equity"]["py"]
    int_cy = agg["interest_expense"]["cy"]; int_py = agg["interest_expense"]["py"]
    fa_cy  = agg["fixed_assets"]["cy"]; fa_py  = agg["fixed_assets"]["py"]
    cogs_cy = agg["cogs"]["cy"];        cogs_py = agg["cogs"]["py"]
    wc_cy  = agg["working_capital"]["cy"]; wc_py = agg["working_capital"]["py"]

    has_rev = rev_cy != 0
    has_bs  = ta_cy != 0
    has_cl  = cl_cy != 0

    ratios = [
        # Profitability
        _make_ratio("Gross Profit Margin (%)",
            f"Gross Profit / Revenue × 100 ({ref['revenue']})",
            _safe_div(gp_cy, rev_cy, 100), _safe_div(gp_py, rev_py, 100),
            [("yoy>", 10, "🟡 MODERATE — Explain"), ("<", 0, "🔴 HIGH — Investigate")],
            f"Declining margin → rising COGS, pricing pressure, or revenue understatement. Verify under {ref['revenue']}.",
            ref["ap"], has_rev),

        _make_ratio("Net Profit Margin (%)",
            "Net Profit / Revenue × 100",
            _safe_div(np_cy, rev_cy, 100), _safe_div(np_py, rev_py, 100),
            [("<", 0, "🔴 HIGH — Investigate"), ("yoy>", 20, "🟡 MODERATE — Explain")],
            f"Sustained losses → going concern trigger under {ref['gc']}. Check interest burden and provisions.",
            ref["ap"], has_rev),

        _make_ratio("Operating Profit Margin (%)",
            "Operating Profit / Revenue × 100",
            _safe_div(op_cy, rev_cy, 100), _safe_div(op_py, rev_py, 100),
            [("<", 0, "🔴 HIGH — Investigate"), ("yoy>", 15, "🟡 MODERATE — Explain")],
            "Operating loss before interest → verify expense completeness and revenue cut-off.",
            ref["ap"], has_rev),

        _make_ratio("Revenue Growth (%)",
            "(CY Revenue − PY Revenue) / PY Revenue × 100",
            _yoy(rev_cy, rev_py), None,
            [("yoy>", 25, "🟡 MODERATE — Explain")],
            f"Significant growth → verify cut-off, returns, credit notes. Ref: {ref['revenue']} five-step model.",
            ref["ap"], has_rev and rev_py != 0),

        _make_ratio("COGS % of Revenue",
            "COGS / Revenue × 100",
            _safe_div(cogs_cy, rev_cy, 100), _safe_div(cogs_py, rev_py, 100),
            [("yoy>", 10, "🟡 MODERATE — Explain")],
            f"Rising COGS% → cost inflation, waste, incorrect inventory valuation, or fictitious purchases.",
            ref["ap"], has_rev),

        _make_ratio("Return on Equity (%)",
            "Net Profit / Total Equity × 100",
            _safe_div(np_cy, eq_cy, 100), _safe_div(np_py, eq_py, 100),
            [("<", 0, "🔴 HIGH — Investigate"), ("yoy>", 20, "🟡 MODERATE — Explain")],
            "Negative ROE → entity is loss-making. Verify equity balances and profit allocation.",
            ref["ap"], has_bs and eq_cy != 0),

        _make_ratio("Return on Assets (%)",
            "Net Profit / Total Assets × 100",
            _safe_div(np_cy, ta_cy, 100), _safe_div(np_py, ta_py, 100),
            [("<", 0, "🔴 HIGH — Investigate")],
            f"Declining ROA → assets not generating returns. Consider {ref['impairment']} impairment review.",
            ref["ap"], has_bs),

        # Liquidity
        _make_ratio("Current Ratio",
            "Current Assets / Current Liabilities",
            _safe_div(ca_cy, cl_cy), _safe_div(ca_py, cl_py),
            [("<", 1.0, "🔴 HIGH — Investigate"), ("<", 1.5, "🟡 MODERATE — Explain")],
            f"Ratio <1 → {ref['gc']} going concern trigger. Verify receivable recoverability and creditor terms.",
            ref["gc"], has_bs and has_cl),

        _make_ratio("Quick Ratio",
            "(Current Assets − Inventory) / Current Liabilities",
            _safe_div(ca_cy - cogs_cy * 0.3, cl_cy), _safe_div(ca_py - cogs_py * 0.3, cl_py),
            [("<", 0.8, "🔴 HIGH — Investigate"), ("<", 1.0, "🟡 MODERATE — Explain")],
            "Low quick ratio → liquidity depends heavily on inventory realisation. Verify stock valuation.",
            ref["gc"], has_bs and has_cl),

        _make_ratio("Working Capital Adequacy",
            "Working Capital / Revenue",
            _safe_div(wc_cy, rev_cy), _safe_div(wc_py, rev_py),
            [("<", 0, "🔴 HIGH — Investigate")],
            "Negative working capital → operational funding risk. Flag under going concern assessment.",
            ref["gc"], has_rev and has_cl),

        # Leverage
        _make_ratio("Debt-to-Equity Ratio",
            "Total Liabilities / Total Equity",
            _safe_div(tl_cy, eq_cy), _safe_div(tl_py, eq_py),
            [(">", 3.0, "🔴 HIGH — Investigate"), (">", 2.0, "🟡 MODERATE — Explain")],
            f"High leverage → repayment risk. Verify loan covenants and {ref['gc']} going concern.",
            ref["gc"], has_bs and eq_cy != 0),

        _make_ratio("Interest Coverage Ratio",
            "Operating Profit / Interest Expense",
            _safe_div(op_cy, int_cy), _safe_div(op_py, int_py),
            [("<", 1.5, "🔴 HIGH — Investigate"), ("<", 2.5, "🟡 MODERATE — Explain")],
            f"Low coverage → entity may struggle to service debt. {ref['gc']} going concern indicator.",
            ref["gc"], int_cy != 0),

        _make_ratio("Debt-to-Assets Ratio",
            "Total Liabilities / Total Assets",
            _safe_div(tl_cy, ta_cy), _safe_div(tl_py, ta_py),
            [(">", 0.8, "🔴 HIGH — Investigate"), (">", 0.6, "🟡 MODERATE — Explain")],
            "High debt-to-assets → solvency concern. Verify asset valuations and debt completeness.",
            ref["ap"], has_bs),

        # Efficiency
        _make_ratio("Asset Turnover",
            "Revenue / Total Assets",
            _safe_div(rev_cy, ta_cy), _safe_div(rev_py, ta_py),
            [("yoy>", 25, "🟡 MODERATE — Explain")],
            f"Significant change → review asset additions/disposals under {ref['impairment']}.",
            ref["ap"], has_rev and has_bs),

        _make_ratio("Fixed Asset Intensity",
            "Fixed Assets / Total Assets",
            _safe_div(fa_cy, ta_cy), _safe_div(fa_py, ta_py),
            [("yoy>", 20, "🟡 MODERATE — Explain")],
            f"Large change → verify capital expenditure, disposals and depreciation policy under {ref['impairment']}.",
            ref["ap"], has_bs and fa_cy != 0),
    ]
    return ratios


def calculate_variance(df, mat_pct, unit_info):
    rows = []
    threshold_abs = unit_info["threshold"]
    for _, row in df.iterrows():
        cy = row["CY Amount"]
        py = row["PY Amount"]
        change = cy - py
        if py == 0 and cy == 0:
            pct = 0.0
        elif py == 0:
            pct = 999.0
        else:
            pct = (change / abs(py)) * 100

        flag = "🟢 OK"
        if py == 0 and abs(cy) > threshold_abs:
            flag = "🔴 HIGH — Investigate"
        elif abs(pct) >= mat_pct and abs(change) >= threshold_abs:
            flag = "🔴 HIGH — Investigate" if abs(pct) >= mat_pct * 1.5 else "🟡 MODERATE — Explain"

        rows.append({
            "Account Name": row["Account Name"],
            "Category": row["Category"],
            "CY Amount": cy,
            "PY Amount": py,
            "Change": change,
            "Change (%)": pct,
            "Flag": flag,
        })

    result = pd.DataFrame(rows)
    # Sort: flagged first
    priority = {"🔴 HIGH — Investigate": 0, "🟡 MODERATE — Explain": 1, "🟢 OK": 2}
    result["_sort"] = result["Flag"].map(priority).fillna(3)
    result = result.sort_values("_sort").drop("_sort", axis=1).reset_index(drop=True)
    return result


def going_concern(agg, ratios, standard):
    ref = STD_REFS.get(standard, STD_REFS["IFRS"])
    ratio_map = {r["Ratio"]: r for r in ratios if r.get("Available")}

    indicators = []
    score = 0

    def ind(name, status, finding, reference):
        return {"Indicator": name, "Status": status, "Finding": finding, "Reference": reference}

    # 1. Profitability
    np_cy = agg["net_profit"]["cy"]
    if np_cy is not None and np_cy < 0:
        indicators.append(ind("Net loss in current year", "🔴 CONCERN", f"Net loss of {np_cy:,.2f}. Entity is loss-making.", f"{ref['gc']} para A2(a)"))
        score += 2
    else:
        indicators.append(ind("Net loss in current year", "🟢 CLEAR", f"Entity is profitable. Net profit: {np_cy:,.2f}." if np_cy else "Profit data not available.", f"{ref['gc']} para A2(a)"))

    # 2. Current ratio
    cr_ratio = ratio_map.get("Current Ratio")
    if cr_ratio and cr_ratio["CY"] is not None:
        if cr_ratio["CY"] < 1.0:
            indicators.append(ind("Current Ratio < 1 (liquidity failure)", "🔴 CONCERN", f"Current ratio = {cr_ratio['CY']:.2f}. Entity cannot meet short-term obligations.", f"{ref['gc']} para A2(b)"))
            score += 2
        elif cr_ratio["CY"] < 1.5:
            indicators.append(ind("Current Ratio < 1.5 (liquidity warning)", "🟡 MONITOR", f"Current ratio = {cr_ratio['CY']:.2f}. Monitor closely.", f"{ref['gc']} para A2(b)"))
            score += 1
        else:
            indicators.append(ind("Current Ratio < 1 (liquidity failure)", "🟢 CLEAR", f"Current ratio = {cr_ratio['CY']:.2f}. Adequate liquidity.", f"{ref['gc']} para A2(b)"))
    else:
        indicators.append(ind("Current Ratio", "⚪ N/A", "Insufficient data.", f"{ref['gc']} para A2(b)"))

    # 3. Working capital
    wc_cy = agg["working_capital"]["cy"]
    if wc_cy is not None and wc_cy < 0:
        indicators.append(ind("Negative working capital", "🔴 CONCERN", f"Working capital = {wc_cy:,.2f}. Negative working capital.", f"{ref['gc']} para A2(b)"))
        score += 2
    else:
        indicators.append(ind("Negative working capital", "🟢 CLEAR", f"Working capital = {wc_cy:,.2f}. Positive." if wc_cy else "N/A.", f"{ref['gc']} para A2(b)"))

    # 4. Interest coverage
    ic = ratio_map.get("Interest Coverage Ratio")
    if ic and ic["CY"] is not None:
        if ic["CY"] < 1.0:
            indicators.append(ind("Interest coverage < 1 (debt service failure)", "🔴 CONCERN", f"Coverage = {ic['CY']:.2f}. Cannot cover interest from operations.", f"{ref['gc']} para A2(c)"))
            score += 2
        elif ic["CY"] < 2.0:
            indicators.append(ind("Interest coverage < 2 (thin coverage)", "🟡 MONITOR", f"Coverage = {ic['CY']:.2f}. Vulnerable to earnings decline.", f"{ref['gc']} para A2(c)"))
            score += 1
        else:
            indicators.append(ind("Interest coverage < 1 (debt service failure)", "🟢 CLEAR", f"Coverage = {ic['CY']:.2f}. Adequate.", f"{ref['gc']} para A2(c)"))
    else:
        indicators.append(ind("Interest Coverage", "⚪ N/A", "No interest expense data.", f"{ref['gc']} para A2(c)"))

    # 5. Debt-to-equity
    de = ratio_map.get("Debt-to-Equity Ratio")
    if de and de["CY"] is not None:
        if de["CY"] > 3.0:
            indicators.append(ind("Excessive leverage (D/E > 3)", "🔴 CONCERN", f"D/E = {de['CY']:.2f}. High debt burden.", f"{ref['gc']} para A2(d)"))
            score += 1
        else:
            indicators.append(ind("Excessive leverage (D/E > 3)", "🟢 CLEAR", f"D/E = {de['CY']:.2f}. Within acceptable range.", f"{ref['gc']} para A2(d)"))
    else:
        indicators.append(ind("Excessive leverage", "⚪ N/A", "Insufficient balance sheet data.", f"{ref['gc']} para A2(d)"))

    # 6. Revenue trend
    rev_growth = ratio_map.get("Revenue Growth (%)")
    if rev_growth and rev_growth["CY"] is not None:
        if rev_growth["CY"] < -20:
            indicators.append(ind("Significant revenue decline (>20%)", "🔴 CONCERN", f"Revenue fell {rev_growth['CY']:.1f}%. Investigate customer loss, market conditions.", f"{ref['gc']} para A2(a)"))
            score += 1
        elif rev_growth["CY"] < -10:
            indicators.append(ind("Revenue decline (10–20%)", "🟡 MONITOR", f"Revenue fell {rev_growth['CY']:.1f}%. Monitor trend.", f"{ref['gc']} para A2(a)"))
        else:
            indicators.append(ind("Significant revenue decline (>20%)", "🟢 CLEAR", f"Revenue growth = {rev_growth['CY']:.1f}%. No adverse trend.", f"{ref['gc']} para A2(a)"))
    else:
        indicators.append(ind("Revenue trend", "⚪ N/A", "Insufficient revenue data.", f"{ref['gc']} para A2(a)"))

    # Overall
    if score == 0:
        risk = "LOW"
        conclusion = f"LOW GOING CONCERN RISK — No significant indicators identified. Document that going concern basis is appropriate under {ref['gc']}. No material uncertainty exists."
    elif score <= 2:
        risk = "MODERATE"
        conclusion = f"MODERATE GOING CONCERN RISK — Some indicators present. Obtain management's assessment and future cash flow projections. Document procedures under {ref['gc']} para 16–17."
    elif score <= 5:
        risk = "HIGH"
        conclusion = f"HIGH GOING CONCERN RISK — Multiple indicators identified. Perform extended going concern procedures under {ref['gc']}. Consider whether a material uncertainty paragraph is required."
    else:
        risk = "CRITICAL"
        conclusion = f"CRITICAL GOING CONCERN RISK — Entity's ability to continue as a going concern is in serious doubt. Consider modified opinion under {ref['gc']} para 21–24. Discuss with engagement partner."

    return {"overall_risk": risk, "score": score, "conclusion": conclusion, "indicators": indicators}


def run_analysis(df, unit_info, mat_pct, standard):
    agg    = aggregate(df)
    ratios = calculate_ratios(agg, standard)
    df_var = calculate_variance(df, mat_pct, unit_info)
    gc     = going_concern(agg, ratios, standard)
    return {"agg": agg, "ratios": ratios, "variance": df_var, "going_concern": gc}


def run_benford(amounts, label="Transactions"):
    amounts = [a for a in amounts if a > 0]
    n = len(amounts)

    if n < 20:
        return {
            "sufficient": False,
            "message": f"Only {n} amounts found. Benford's Law requires at least 20 amounts for meaningful analysis.",
            "risk_flag": "⚪ Insufficient Data",
        }

    # First digit counts
    counts = {d: 0 for d in range(1, 10)}
    for a in amounts:
        fd = int(str(abs(a)).replace(".", "").lstrip("0")[0])
        if fd in counts:
            counts[fd] += 1

    expected = {d: math.log10(1 + 1/d) for d in range(1, 10)}

    digit_data = []
    chi_sq = 0
    for d in range(1, 10):
        obs_count = counts[d]
        obs_pct   = obs_count / n * 100
        exp_pct   = expected[d] * 100
        dev_pp    = obs_pct - exp_pct
        dev_pct   = (dev_pp / exp_pct * 100) if exp_pct else 0
        chi_sq   += ((obs_count - expected[d] * n) ** 2) / (expected[d] * n)
        digit_data.append({
            "Digit":         d,
            "Observed Count": obs_count,
            "Observed (%)":  round(obs_pct, 2),
            "Expected (%)":  round(exp_pct, 2),
            "Deviation (pp)": round(dev_pp, 2),
            "Deviation (%)": round(dev_pct, 1),
        })

    # p<0.05 critical value df=8 = 15.507
    if chi_sq > 20.09:
        risk_flag = "🔴 HIGH — Significant deviation from Benford's Law"
        interp = f"Chi-square = {chi_sq:.2f} exceeds critical value at p<0.01 (20.09). Significant deviation detected. This may indicate manipulation, rounding to specific amounts, or fictitious entries. Extend fraud procedures under ISA/SA 240."
    elif chi_sq > 15.51:
        risk_flag = "🟡 MODERATE — Notable deviation, review recommended"
        interp = f"Chi-square = {chi_sq:.2f} exceeds critical value at p<0.05 (15.51). Some deviation present. Obtain explanation for unusual digit patterns and consider targeted transaction testing."
    elif chi_sq > 13.36:
        risk_flag = "🟡 LOW-MODERATE — Minor deviation"
        interp = f"Chi-square = {chi_sq:.2f} slightly elevated. No strong evidence of manipulation but worth noting in working papers."
    else:
        risk_flag = "🟢 LOW — Distribution consistent with Benford's Law"
        interp = f"Chi-square = {chi_sq:.2f}. Distribution is consistent with Benford's Law. No significant fraud indicator from first-digit analysis."

    return {
        "sufficient": True,
        "n": n,
        "label": label,
        "chi_square": round(chi_sq, 2),
        "digit_data": digit_data,
        "risk_flag": risk_flag,
        "interpretation": interp,
    }
