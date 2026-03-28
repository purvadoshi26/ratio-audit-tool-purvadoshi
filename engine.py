"""
engine.py — Audit Analytical Procedures Engine
Supports Ind AS / IFRS / US GAAP
Features: fuzzy classification, validation, materiality logic,
          risk scoring, automated audit commentary
"""

import pandas as pd
import math
import re

# ── Standard references ───────────────────────────────────────
STD_REFS = {
    "Ind AS":  {"revenue": "Ind AS 115", "impairment": "Ind AS 36", "leases": "Ind AS 116",
                "ap": "SA 520", "gc": "SA 570", "fraud": "SA 240", "risk": "SA 315"},
    "IFRS":    {"revenue": "IFRS 15",    "impairment": "IAS 36",    "leases": "IFRS 16",
                "ap": "ISA 520", "gc": "ISA 570", "fraud": "ISA 240", "risk": "ISA 315"},
    "US GAAP": {"revenue": "ASC 606",    "impairment": "ASC 350",   "leases": "ASC 842",
                "ap": "AU-C 520", "gc": "AU-C 570", "fraud": "AU-C 240", "risk": "AU-C 315"},
}

SUBTOTAL_KEYWORDS = [
    "total", "subtotal", "grand total", "net total", "profit before",
    "profit after", "loss before", "loss after", "ebitda", "ebit", "pbt", "pat",
    "total income", "total expense", "total revenue", "total cost",
    "total assets", "total liabilities", "total equity", "total capital",
    "net worth", "shareholders fund", "net profit", "net loss",
]

# ── Fuzzy keyword classification ──────────────────────────────
# Each entry: (keywords, category, weight)
# Higher weight = stronger match
KEYWORD_RULES = [
    # Current Assets
    (["cash", "bank balance", "petty cash", "cash in hand", "cash at bank"], "current_assets", 3),
    (["debtor", "receivable", "trade receivable", "sundry debtor", "book debt", "bills receivable"], "current_assets", 3),
    (["inventory", "stock", "stock in trade", "raw material", "wip", "work in progress", "finished good", "closing stock"], "current_assets", 3),
    (["advance paid", "advance to supplier", "prepaid", "advance given", "vendor advance"], "current_assets", 2),
    (["tds receivable", "tds refund", "income tax refundable", "advance tax", "tax asset"], "current_assets", 2),
    (["gst receivable", "input tax credit", "itc", "gst credit", "cenvat"], "current_assets", 2),
    (["short term investment", "liquid fund", "mutual fund", "fdr", "fixed deposit"], "current_assets", 2),
    (["accrued income", "interest receivable", "dividend receivable"], "current_assets", 2),
    (["other current asset", "miscellaneous current"], "current_assets", 1),

    # Fixed Assets
    (["plant", "machinery", "equipment", "machine"], "fixed_assets", 3),
    (["land", "building", "premise", "property"], "fixed_assets", 3),
    (["furniture", "fixture", "office furniture"], "fixed_assets", 2),
    (["vehicle", "car", "truck", "motor"], "fixed_assets", 2),
    (["computer", "laptop", "hardware", "software", "it asset"], "fixed_assets", 2),
    (["cwip", "capital wip", "capital work in progress", "construction in progress"], "fixed_assets", 3),
    (["intangible", "goodwill", "patent", "trademark", "copyright", "license"], "fixed_assets", 3),
    (["net block", "gross block", "fixed asset", "tangible asset"], "fixed_assets", 3),
    (["right of use", "rou asset", "lease asset"], "fixed_assets", 2),

    # Current Liabilities
    (["creditor", "payable", "trade payable", "sundry creditor", "supplier payable", "bills payable"], "current_liabilities", 3),
    (["bank od", "overdraft", "cash credit", "working capital loan", "short term loan", "short term borrowing"], "current_liabilities", 3),
    (["gst payable", "igst", "cgst", "sgst", "tax payable", "tds payable", "tcs payable"], "current_liabilities", 2),
    (["advance received", "advance from customer", "customer deposit", "unearned revenue", "deferred income"], "current_liabilities", 2),
    (["outstanding expense", "accrued expense", "accrued liability", "payroll payable", "salary payable"], "current_liabilities", 2),
    (["current portion", "current maturity", "installment due"], "current_liabilities", 2),
    (["provision for tax", "income tax payable", "current tax"], "current_liabilities", 2),
    (["other current liability", "misc payable"], "current_liabilities", 1),

    # Long-term Liabilities
    (["term loan", "long term loan", "long term borrowing", "secured loan", "unsecured loan"], "long_term_liabilities", 3),
    (["debenture", "bond", "ncd", "non convertible debenture"], "long_term_liabilities", 3),
    (["deferred tax liability", "dtl"], "long_term_liabilities", 3),
    (["security deposit received", "long term deposit"], "long_term_liabilities", 2),
    (["long term provision", "gratuity liability", "leave encashment liability"], "long_term_liabilities", 2),
    (["mortgage", "hypothecation"], "long_term_liabilities", 2),
    (["lease liability", "finance lease liability"], "long_term_liabilities", 2),

    # Equity
    (["share capital", "equity share", "preference share", "paid up capital", "authorised capital"], "equity", 3),
    (["reserve", "surplus", "general reserve", "capital reserve", "retained earning", "profit and loss account"], "equity", 3),
    (["securities premium", "share premium"], "equity", 3),
    (["other equity", "other comprehensive income", "oci"], "equity", 2),
    (["minority interest", "non controlling interest"], "equity", 2),

    # Revenue
    (["revenue from operation", "sales", "turnover", "net sales", "gross sales"], "revenue", 3),
    (["other income", "miscellaneous income", "sundry income"], "revenue", 3),
    (["interest income", "interest received", "interest earned"], "revenue", 2),
    (["dividend income", "dividend received"], "revenue", 2),
    (["rental income", "rent income", "lease income"], "revenue", 2),
    (["commission income", "commission received"], "revenue", 2),
    (["export sale", "domestic sale", "service income", "service revenue"], "revenue", 2),
    (["gain on sale", "profit on sale"], "revenue", 2),

    # COGS
    (["purchase", "raw material consumed", "material cost", "cost of material"], "cogs", 3),
    (["changes in inventor", "change in stock", "opening stock", "closing stock adjustment"], "cogs", 3),
    (["direct labour", "job work", "contract labour", "direct wage"], "cogs", 3),
    (["cost of goods sold", "cogs", "cost of production", "cost of sales"], "cogs", 3),
    (["freight inward", "carriage inward", "import duty"], "cogs", 2),

    # Expenses
    (["salary", "wage", "staff cost", "employee benefit", "employee expense", "remuneration", "manpower"], "expenses", 3),
    (["depreciation", "amortisation", "amortization", "d&a"], "expenses", 3),
    (["rent", "lease rent", "office rent"], "expenses", 2),
    (["electricity", "power", "fuel", "utilities"], "expenses", 2),
    (["telephone", "communication", "internet expense"], "expenses", 2),
    (["repair", "maintenance", "amc"], "expenses", 2),
    (["advertisement", "marketing", "promotion", "branding"], "expenses", 2),
    (["travelling", "travel", "conveyance", "lodging"], "expenses", 2),
    (["legal", "professional fee", "consultancy", "audit fee", "retainer"], "expenses", 2),
    (["insurance premium", "insurance expense"], "expenses", 2),
    (["printing", "stationery", "office supply"], "expenses", 2),
    (["selling expense", "distribution", "freight outward", "carriage outward"], "expenses", 2),
    (["administrative", "general expense", "miscellaneous expense", "other expense"], "expenses", 1),
    (["csr", "donation", "charity"], "expenses", 1),

    # Interest Expense
    (["interest expense", "interest on loan", "finance cost", "borrowing cost", "interest paid"], "interest_expense", 3),
    (["bank charge", "processing fee", "bank commission", "loan charge"], "interest_expense", 2),
    (["interest on od", "interest on cc", "interest on working capital"], "interest_expense", 2),
]


def _classify(name: str) -> str:
    n = name.lower().strip()
    n = re.sub(r'[^a-z0-9 /&\-()]', ' ', n)
    n = re.sub(r'\s+', ' ', n).strip()

    scores = {}
    for keywords, category, weight in KEYWORD_RULES:
        for kw in keywords:
            if kw in n:
                scores[category] = scores.get(category, 0) + weight

    if not scores:
        return "unclassified"
    return max(scores, key=scores.get)


def _is_subtotal(name: str) -> bool:
    n = name.lower().strip()
    for kw in SUBTOTAL_KEYWORDS:
        if kw in n:
            return True
    # Also skip rows that are ONLY numbers or dashes
    if re.match(r'^[\d\s\-\.,]+$', n):
        return True
    return False


def _clean_amount(val) -> float:
    if val is None or (isinstance(val, float) and math.isnan(val)):
        return 0.0
    s = str(val).strip()
    if s.lower() in ["-", "—", "", "nil", "n/a", "na", "none", "null", "nan"]:
        return 0.0
    s = s.replace(",", "").replace("₹", "").replace("$", "").replace("£", "").strip()
    s = s.replace("(", "-").replace(")", "")
    s = s.replace(" cr", "").replace(" lakh", "").strip()
    try:
        return float(s)
    except:
        return 0.0


def _find_col(df, candidates):
    for cand in candidates:
        for col in df.columns:
            if cand.lower() in str(col).lower():
                return col
    return None


# ── VALIDATION ────────────────────────────────────────────────
def validate_df(df):
    errors = []
    warnings = []

    required = ["Account Name", "CY Amount", "PY Amount"]
    for r in required:
        if r not in df.columns:
            errors.append(f"Missing required column: '{r}'")

    if errors:
        return errors, warnings

    # Check numeric
    non_num_cy = df[pd.to_numeric(df["CY Amount"], errors="coerce").isna() & df["CY Amount"].notna()]
    if len(non_num_cy):
        warnings.append(f"{len(non_num_cy)} rows have non-numeric CY Amount — treated as 0")

    non_num_py = df[pd.to_numeric(df["PY Amount"], errors="coerce").isna() & df["PY Amount"].notna()]
    if len(non_num_py):
        warnings.append(f"{len(non_num_py)} rows have non-numeric PY Amount — treated as 0")

    # Check all zeros
    if (df["CY Amount"] == 0).all():
        warnings.append("All CY Amounts are zero — check your file")

    # Check unclassified
    unclass = df[df["Category"] == "unclassified"]
    if len(unclass) > 0:
        warnings.append(f"{len(unclass)} accounts could not be auto-classified — review manually")

    return errors, warnings


# ── FILE LOADER ───────────────────────────────────────────────
def load_tb(filepath: str):
    try:
        if filepath.endswith(".csv"):
            df = pd.read_csv(filepath)
        else:
            # Try first sheet, then look for sheet named 'Trial Balance'
            xl = pd.ExcelFile(filepath)
            sheet = "Trial Balance" if "Trial Balance" in xl.sheet_names else xl.sheet_names[0]
            df = pd.read_excel(filepath, sheet_name=sheet)
    except Exception as e:
        raise ValueError(f"Cannot read file: {e}")

    df.columns = [str(c).strip() for c in df.columns]

    name_col = _find_col(df, ["account name", "account", "particulars", "description", "head", "ledger"])
    cy_col   = _find_col(df, ["cy amount", "cy", "current year", "2025", "2024", "2023", "fy 25", "fy 24"])
    py_col   = _find_col(df, ["py amount", "py", "prior year", "previous year", "2024", "2023", "2022", "fy 24", "fy 23"])
    cat_col  = _find_col(df, ["category", "type", "classification", "group", "head of account"])

    # cy and py might be same col if only 2 number cols — disambiguate
    if cy_col == py_col and cy_col is not None:
        num_cols = [c for c in df.columns if c != (name_col or "") and pd.to_numeric(df[c], errors="coerce").notna().sum() > len(df) * 0.5]
        if len(num_cols) >= 2:
            cy_col, py_col = num_cols[0], num_cols[1]
        else:
            cy_col = num_cols[0] if num_cols else None
            py_col = None

    if name_col is None:
        raise ValueError("Cannot find 'Account Name' column. Expected column named: Account Name, Account, Particulars, or Description.")
    if cy_col is None:
        raise ValueError("Cannot find 'CY Amount' column. Expected column named: CY Amount, Current Year, or similar.")
    if py_col is None:
        raise ValueError("Cannot find 'PY Amount' column. Expected column named: PY Amount, Prior Year, or similar.")

    df = df.rename(columns={name_col: "Account Name", cy_col: "CY Amount", py_col: "PY Amount"})
    if cat_col and cat_col not in ["CY Amount", "PY Amount"]:
        df = df.rename(columns={cat_col: "Category"})
    else:
        df["Category"] = ""

    # Drop nulls and subtotals
    df = df[df["Account Name"].notna()].copy()
    df["Account Name"] = df["Account Name"].astype(str).str.strip()
    df = df[df["Account Name"] != ""].copy()
    df = df[~df["Account Name"].apply(_is_subtotal)].copy()

    df["CY Amount"] = df["CY Amount"].apply(_clean_amount)
    df["PY Amount"] = df["PY Amount"].apply(_clean_amount)

    # Drop rows that are all zero
    df = df[(df["CY Amount"] != 0) | (df["PY Amount"] != 0)].copy()

    # Classify
    valid_cats = {"current_assets", "fixed_assets", "current_liabilities", "long_term_liabilities",
                  "equity", "revenue", "cogs", "expenses", "interest_expense"}
    df["Category"] = df.apply(
        lambda r: r["Category"].strip().lower()
        if str(r.get("Category", "")).strip().lower() in valid_cats
        else _classify(r["Account Name"]),
        axis=1
    )

    # Unit detection
    max_val = df[["CY Amount", "PY Amount"]].abs().max().max()
    if max_val >= 1_00_00_000:   # >= 1 crore
        label, threshold = "₹ Crores", 2.0
        df["CY Amount"] = df["CY Amount"] / 1_00_00_000
        df["PY Amount"] = df["PY Amount"] / 1_00_00_000
    elif max_val >= 1_00_000:    # >= 1 lakh
        label, threshold = "₹ Lakhs", 5.0
        df["CY Amount"] = df["CY Amount"] / 1_00_000
        df["PY Amount"] = df["PY Amount"] / 1_00_000
    elif max_val >= 1:
        label, threshold = "₹ Crores", 0.5
    else:
        label, threshold = "₹", 50000.0

    unit_info = {"label": label, "threshold": threshold, "max_val": max_val}

    errors, warnings = validate_df(df)
    if errors:
        raise ValueError(" | ".join(errors))

    df = df.reset_index(drop=True)
    return df, unit_info, warnings


# ── AGGREGATION ───────────────────────────────────────────────
def aggregate(df):
    def s(cats):
        sub = df[df["Category"].isin(cats if isinstance(cats, list) else [cats])]
        return {"cy": sub["CY Amount"].sum(), "py": sub["PY Amount"].sum()}

    rev   = s("revenue")
    cogs  = s("cogs")
    exp   = s("expenses")
    int_e = s("interest_expense")
    ca    = s("current_assets")
    fa    = s("fixed_assets")
    cl    = s("current_liabilities")
    ltl   = s("long_term_liabilities")
    eq    = s("equity")

    gp   = {"cy": rev["cy"] - cogs["cy"],          "py": rev["py"] - cogs["py"]}
    op   = {"cy": gp["cy"]  - exp["cy"],            "py": gp["py"]  - exp["py"]}
    np_  = {"cy": op["cy"]  - int_e["cy"],          "py": op["py"]  - int_e["py"]}
    ta   = {"cy": ca["cy"]  + fa["cy"],             "py": ca["py"]  + fa["py"]}
    tl   = {"cy": cl["cy"]  + ltl["cy"],            "py": cl["py"]  + ltl["py"]}
    wc   = {"cy": ca["cy"]  - cl["cy"],             "py": ca["py"]  - cl["py"]}

    return {
        "revenue": rev, "cogs": cogs, "expenses": exp, "interest_expense": int_e,
        "gross_profit": gp, "operating_profit": op, "net_profit": np_,
        "current_assets": ca, "fixed_assets": fa,
        "current_liabilities": cl, "long_term_liabilities": ltl,
        "equity": eq, "total_assets": ta, "total_liabilities": tl, "working_capital": wc,
    }


# ── RATIOS ────────────────────────────────────────────────────
def _safe_div(a, b, scale=1):
    if b is None or b == 0 or a is None: return None
    return (a / b) * scale

def _yoy(cy, py):
    if py is None or py == 0 or cy is None: return None
    return (cy - py) / abs(py) * 100

def _make_ratio(name, formula, cy, py, flag_rules, note, ref, section):
    if cy is None:
        return {"Ratio": name, "Formula": formula, "CY": None, "PY": None,
                "YoY (%)": None, "Flag": "⚪ N/A", "Note": note,
                "Ref": ref, "Section": section, "Available": False}
    yoy = _yoy(cy, py)
    flag = "🟢 OK"
    for (op, val, lbl) in flag_rules:
        try:
            if op == "<"    and cy < val:                              flag = lbl; break
            if op == ">"    and cy > val:                              flag = lbl; break
            if op == "yoy>" and yoy is not None and abs(yoy) > val:   flag = lbl; break
            if op == "yoy<" and yoy is not None and yoy < val:        flag = lbl; break
        except: pass
    return {
        "Ratio": name, "Formula": formula,
        "CY": round(cy, 2), "PY": round(py, 2) if py is not None else None,
        "YoY (%)": round(yoy, 1) if yoy is not None else None,
        "Flag": flag, "Note": note, "Ref": ref,
        "Section": section, "Available": True,
    }


def calculate_ratios(agg, standard):
    ref = STD_REFS.get(standard, STD_REFS["IFRS"])
    r = agg
    rev_cy = r["revenue"]["cy"];          rev_py = r["revenue"]["py"]
    gp_cy  = r["gross_profit"]["cy"];     gp_py  = r["gross_profit"]["py"]
    np_cy  = r["net_profit"]["cy"];       np_py  = r["net_profit"]["py"]
    op_cy  = r["operating_profit"]["cy"]; op_py  = r["operating_profit"]["py"]
    ca_cy  = r["current_assets"]["cy"];   ca_py  = r["current_assets"]["py"]
    cl_cy  = r["current_liabilities"]["cy"]; cl_py = r["current_liabilities"]["py"]
    ta_cy  = r["total_assets"]["cy"];     ta_py  = r["total_assets"]["py"]
    tl_cy  = r["total_liabilities"]["cy"]; tl_py = r["total_liabilities"]["py"]
    eq_cy  = r["equity"]["cy"];           eq_py  = r["equity"]["py"]
    int_cy = r["interest_expense"]["cy"]; int_py = r["interest_expense"]["py"]
    fa_cy  = r["fixed_assets"]["cy"];     fa_py  = r["fixed_assets"]["py"]
    cogs_cy= r["cogs"]["cy"];             cogs_py= r["cogs"]["py"]
    wc_cy  = r["working_capital"]["cy"];  wc_py  = r["working_capital"]["py"]

    has_rev = rev_cy != 0
    has_bs  = ta_cy  != 0
    has_cl  = cl_cy  != 0

    ratios = [
        _make_ratio("Gross Profit Margin (%)",
            f"(Revenue − COGS) / Revenue × 100",
            _safe_div(gp_cy, rev_cy, 100) if has_rev else None,
            _safe_div(gp_py, rev_py, 100),
            [("<", 0, "🔴 HIGH"), ("yoy>", 10, "🟡 MODERATE")],
            f"Decline → rising COGS, pricing pressure, or revenue understatement. Verify under {ref['revenue']}.",
            ref["ap"], "Profitability"),

        _make_ratio("Net Profit Margin (%)",
            "Net Profit / Revenue × 100",
            _safe_div(np_cy, rev_cy, 100) if has_rev else None,
            _safe_div(np_py, rev_py, 100),
            [("<", 0, "🔴 HIGH"), ("yoy>", 20, "🟡 MODERATE")],
            f"Net loss → going concern trigger. Check interest burden, unusual provisions, and expense cut-off.",
            ref["ap"], "Profitability"),

        _make_ratio("Operating Profit Margin (%)",
            "Operating Profit / Revenue × 100",
            _safe_div(op_cy, rev_cy, 100) if has_rev else None,
            _safe_div(op_py, rev_py, 100),
            [("<", 0, "🔴 HIGH"), ("yoy>", 15, "🟡 MODERATE")],
            "Operating loss → verify expense completeness and revenue recognition cut-off.",
            ref["ap"], "Profitability"),

        _make_ratio("Revenue Growth (%)",
            "(CY − PY) / PY × 100",
            _yoy(rev_cy, rev_py) if has_rev and rev_py != 0 else None, None,
            [("yoy>", 25, "🟡 MODERATE"), ("yoy<", -20, "🔴 HIGH")],
            f"Significant change → verify cut-off, returns, credit notes. Ref: {ref['revenue']}.",
            ref["ap"], "Profitability"),

        _make_ratio("COGS % of Revenue",
            "COGS / Revenue × 100",
            _safe_div(cogs_cy, rev_cy, 100) if has_rev else None,
            _safe_div(cogs_py, rev_py, 100),
            [("yoy>", 10, "🟡 MODERATE")],
            "Rising COGS% → cost inflation, inventory misstatement, or fictitious purchases.",
            ref["ap"], "Profitability"),

        _make_ratio("Return on Equity (%)",
            "Net Profit / Equity × 100",
            _safe_div(np_cy, eq_cy, 100) if eq_cy != 0 else None,
            _safe_div(np_py, eq_py, 100),
            [("<", 0, "🔴 HIGH"), ("yoy>", 20, "🟡 MODERATE")],
            "Negative ROE → loss-making. Verify equity balances and profit allocation.",
            ref["ap"], "Profitability"),

        _make_ratio("Return on Assets (%)",
            "Net Profit / Total Assets × 100",
            _safe_div(np_cy, ta_cy, 100) if has_bs else None,
            _safe_div(np_py, ta_py, 100),
            [("<", 0, "🔴 HIGH")],
            f"Declining ROA → assets underperforming. Consider {ref['impairment']} impairment review.",
            ref["ap"], "Profitability"),

        _make_ratio("Current Ratio",
            "Current Assets / Current Liabilities",
            _safe_div(ca_cy, cl_cy) if has_cl else None,
            _safe_div(ca_py, cl_py),
            [("<", 1.0, "🔴 HIGH"), ("<", 1.5, "🟡 MODERATE")],
            f"< 1 → {ref['gc']} going concern trigger. Verify receivable recoverability.",
            ref["gc"], "Liquidity"),

        _make_ratio("Quick Ratio",
            "(Current Assets − Inventory est.) / Current Liabilities",
            _safe_div(ca_cy - cogs_cy * 0.15, cl_cy) if has_cl else None,
            _safe_div(ca_py - cogs_py * 0.15, cl_py),
            [("<", 0.8, "🔴 HIGH"), ("<", 1.0, "🟡 MODERATE")],
            "Low quick ratio → liquidity depends on inventory. Verify stock realisability.",
            ref["gc"], "Liquidity"),

        _make_ratio("Working Capital / Revenue",
            "Working Capital / Revenue",
            _safe_div(wc_cy, rev_cy) if has_rev else None,
            _safe_div(wc_py, rev_py),
            [("<", 0, "🔴 HIGH")],
            "Negative → operational funding risk. Flag under going concern assessment.",
            ref["gc"], "Liquidity"),

        _make_ratio("Debt-to-Equity",
            "Total Liabilities / Equity",
            _safe_div(tl_cy, eq_cy) if eq_cy != 0 else None,
            _safe_div(tl_py, eq_py),
            [(">", 3.0, "🔴 HIGH"), (">", 2.0, "🟡 MODERATE")],
            f"High leverage → repayment risk. Verify loan covenants and {ref['gc']}.",
            ref["gc"], "Leverage"),

        _make_ratio("Interest Coverage",
            "Operating Profit / Interest Expense",
            _safe_div(op_cy, int_cy) if int_cy != 0 else None,
            _safe_div(op_py, int_py),
            [("<", 1.5, "🔴 HIGH"), ("<", 2.5, "🟡 MODERATE")],
            f"Low coverage → cannot service debt comfortably. {ref['gc']} indicator.",
            ref["gc"], "Leverage"),

        _make_ratio("Debt-to-Assets",
            "Total Liabilities / Total Assets",
            _safe_div(tl_cy, ta_cy) if has_bs else None,
            _safe_div(tl_py, ta_py),
            [(">", 0.8, "🔴 HIGH"), (">", 0.6, "🟡 MODERATE")],
            "High ratio → solvency concern. Verify asset valuations and debt completeness.",
            ref["ap"], "Leverage"),

        _make_ratio("Asset Turnover",
            "Revenue / Total Assets",
            _safe_div(rev_cy, ta_cy) if has_bs else None,
            _safe_div(rev_py, ta_py),
            [("yoy>", 25, "🟡 MODERATE")],
            f"Large change → review asset additions/disposals under {ref['impairment']}.",
            ref["ap"], "Efficiency"),

        _make_ratio("Fixed Asset Intensity",
            "Fixed Assets / Total Assets",
            _safe_div(fa_cy, ta_cy) if has_bs and fa_cy != 0 else None,
            _safe_div(fa_py, ta_py),
            [("yoy>", 20, "🟡 MODERATE")],
            f"Large change → verify capex, disposals, and depreciation policy. Ref: {ref['impairment']}.",
            ref["ap"], "Efficiency"),
    ]
    return ratios


# ── AUDIT COMMENTARY ──────────────────────────────────────────
def _auto_commentary(name, category, cy, py, change_pct, flag, agg):
    rev_cy = agg["revenue"]["cy"]
    rev_py = agg["revenue"]["py"]
    rev_growth = _yoy(rev_cy, rev_py) or 0
    name_l = name.lower()

    if "🔴" not in flag and "🟡" not in flag:
        return "Movement within materiality threshold. No further procedures required at this stage."

    abs_pct = abs(change_pct) if change_pct != 999.0 else None
    direction = "increased" if cy > py else "decreased"

    # Specific account commentary
    if any(k in name_l for k in ["debtor", "receivable", "trade receivable"]):
        if cy > py and abs_pct and abs_pct > rev_growth + 10:
            return f"Trade receivables {direction} by {abs_pct:.1f}%, which exceeds revenue growth of {rev_growth:.1f}%. This inconsistency may indicate collection issues, fictitious debtors, or premature revenue recognition. Obtain debtor ageing schedule and review post-balance sheet collections."
        return f"Trade receivables {direction} by {abs_pct:.1f}%. Verify recoverability by obtaining debtor ageing and reviewing post-balance sheet receipts."

    if any(k in name_l for k in ["inventor", "stock"]):
        return f"Inventory {direction} by {abs_pct:.1f}%. Verify physical stock count, cost vs NRV valuation, and slow-moving/obsolete items."

    if any(k in name_l for k in ["creditor", "payable", "trade payable"]):
        return f"Trade payables {direction} by {abs_pct:.1f}%. Confirm completeness of liabilities and check for unrecorded payables at year-end."

    if any(k in name_l for k in ["cash", "bank"]):
        return f"Cash & Bank {direction} by {abs_pct:.1f}%. Agree to bank reconciliation statements and review large unexplained movements."

    if any(k in name_l for k in ["loan", "borrowing", "term loan"]):
        return f"Loan balance {direction} by {abs_pct:.1f}%. Agree to loan sanction letter, confirm outstanding balance with lender, and verify interest accruals."

    if any(k in name_l for k in ["revenue", "sales", "turnover"]):
        return f"Revenue {direction} by {abs_pct:.1f}%. Perform cut-off testing around year-end and verify revenue recognition policy compliance."

    if any(k in name_l for k in ["salary", "wage", "employee", "staff"]):
        return f"Staff costs {direction} by {abs_pct:.1f}%. Agree to payroll records and HR headcount. Verify any new joinees or leavers."

    if any(k in name_l for k in ["reserve", "surplus", "retained"]):
        return f"Equity reserve {direction} by {abs_pct:.1f}%. Agree to profit appropriation and board resolution. Verify dividend declarations if any."

    if any(k in name_l for k in ["depreciation", "amortis"]):
        return f"Depreciation {direction} by {abs_pct:.1f}%. Agree to fixed asset register and verify rate consistency with prior year."

    if any(k in name_l for k in ["advance", "prepaid"]):
        return f"Advance/Prepaid {direction} by {abs_pct:.1f}%. Verify nature, recoverability, and whether any should be expensed in the current period."

    if any(k in name_l for k in ["other income", "misc income"]):
        return f"Other income {direction} by {abs_pct:.1f}%. Enquire about nature and verify occurrence. Assess whether correctly classified vs revenue."

    if change_pct == 999.0:
        return f"New account with no prior year balance. Confirm nature, obtain supporting documentation, and verify appropriateness of classification."

    # Generic
    cat_map = {
        "current_assets": "Verify existence, completeness, and valuation at year-end.",
        "fixed_assets": f"Review additions, disposals, and ensure impairment assessment under {STD_REFS.get('Ind AS', {}).get('impairment', 'applicable standard')}.",
        "current_liabilities": "Confirm completeness of liabilities and verify no unrecorded obligations.",
        "long_term_liabilities": "Agree to loan agreements and verify classification between current and non-current.",
        "equity": "Agree to statutory records and board minutes.",
        "revenue": "Perform cut-off testing and verify recognition policy.",
        "cogs": "Verify against purchase records, inventory counts, and supplier invoices.",
        "expenses": "Agree to supporting invoices and verify period allocation.",
        "interest_expense": "Agree to loan statements and verify accrual of interest at year-end.",
    }
    base = cat_map.get(category, "Obtain supporting documentation and verify balance.")
    return f"Balance {direction} by {abs_pct:.1f}% which is material. {base}"


# ── VARIANCE WITH MATERIALITY & COMMENTARY ───────────────────
def calculate_variance(df, mat_pct, unit_info, agg):
    threshold_abs = unit_info["threshold"]
    overall_mat = agg["revenue"]["cy"] * 0.05 if agg["revenue"]["cy"] else threshold_abs * 10

    rows = []
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

        is_material = abs(change) >= threshold_abs
        mat_label = "Yes" if is_material else "No"

        flag = "🟢 Within Threshold"
        if pct == 999.0 and abs(cy) >= threshold_abs:
            flag = "🔴 HIGH — Investigate"
        elif is_material and abs(pct) >= mat_pct * 1.5:
            flag = "🔴 HIGH — Investigate"
        elif is_material and abs(pct) >= mat_pct:
            flag = "🟡 MODERATE — Explain"

        # Risk score per account
        risk_score = 0
        if "🔴" in flag: risk_score += 3
        elif "🟡" in flag: risk_score += 1

        commentary = _auto_commentary(
            row["Account Name"], row["Category"], cy, py, pct, flag, agg
        )

        rows.append({
            "Account Name": row["Account Name"],
            "Category": row["Category"],
            "CY Amount": cy,
            "PY Amount": py,
            "Change": change,
            "Change (%)": pct,
            "Material?": mat_label,
            "Flag": flag,
            "Risk Score": risk_score,
            "Audit Commentary": commentary,
        })

    result = pd.DataFrame(rows)
    priority = {"🔴 HIGH — Investigate": 0, "🟡 MODERATE — Explain": 1, "🟢 Within Threshold": 2}
    result["_sort"] = result["Flag"].map(priority).fillna(3)
    result = result.sort_values("_sort").drop("_sort", axis=1).reset_index(drop=True)
    return result


# ── GOING CONCERN ─────────────────────────────────────────────
def going_concern(agg, ratios, standard):
    ref = STD_REFS.get(standard, STD_REFS["IFRS"])
    ratio_map = {r["Ratio"]: r for r in ratios if r.get("Available")}

    indicators = []
    score = 0

    def ind(name, status, finding, reference, detail=""):
        return {"Indicator": name, "Status": status, "Finding": finding,
                "Reference": reference, "Detail": detail}

    np_cy = agg["net_profit"]["cy"]
    np_py = agg["net_profit"]["py"]
    if np_cy is not None and np_cy < 0:
        indicators.append(ind("Net loss in current year", "🔴 CONCERN",
            f"Net loss of {np_cy:,.2f}. Entity is loss-making.", f"{ref['gc']} para A2(a)",
            "Obtain management's explanation. Review future projections."))
        score += 2
        if np_py is not None and np_py < 0:
            score += 1  # consecutive loss
    else:
        indicators.append(ind("Net loss in current year", "🟢 CLEAR",
            f"Entity is profitable. Net profit: {np_cy:,.2f}." if np_cy else "Profit data not available.",
            f"{ref['gc']} para A2(a)"))

    cr = ratio_map.get("Current Ratio")
    if cr and cr["CY"] is not None:
        if cr["CY"] < 1.0:
            indicators.append(ind("Current Ratio < 1", "🔴 CONCERN",
                f"CR = {cr['CY']:.2f}. Entity cannot meet short-term obligations.", f"{ref['gc']} para A2(b)",
                "Obtain cash flow projections and evidence of committed facilities.")); score += 2
        elif cr["CY"] < 1.5:
            indicators.append(ind("Current Ratio < 1.5", "🟡 MONITOR",
                f"CR = {cr['CY']:.2f}. Tight liquidity — monitor closely.", f"{ref['gc']} para A2(b)")); score += 1
        else:
            indicators.append(ind("Current Ratio", "🟢 CLEAR",
                f"CR = {cr['CY']:.2f}. Adequate liquidity.", f"{ref['gc']} para A2(b)"))
    else:
        indicators.append(ind("Current Ratio", "⚪ N/A", "Insufficient data.", f"{ref['gc']} para A2(b)"))

    wc_cy = agg["working_capital"]["cy"]
    if wc_cy is not None and wc_cy < 0:
        indicators.append(ind("Negative working capital", "🔴 CONCERN",
            f"WC = {wc_cy:,.2f}. Negative.", f"{ref['gc']} para A2(b)",
            "Review debt maturity profile and refinancing plans.")); score += 2
    else:
        indicators.append(ind("Negative working capital", "🟢 CLEAR",
            f"WC = {wc_cy:,.2f}. Positive." if wc_cy is not None else "N/A.", f"{ref['gc']} para A2(b)"))

    ic = ratio_map.get("Interest Coverage")
    if ic and ic["CY"] is not None:
        if ic["CY"] < 1.0:
            indicators.append(ind("Interest coverage < 1", "🔴 CONCERN",
                f"Coverage = {ic['CY']:.2f}. Cannot cover interest from operations.", f"{ref['gc']} para A2(c)",
                "Review loan covenant compliance and risk of default.")); score += 2
        elif ic["CY"] < 2.0:
            indicators.append(ind("Interest coverage < 2", "🟡 MONITOR",
                f"Coverage = {ic['CY']:.2f}. Vulnerable to earnings decline.", f"{ref['gc']} para A2(c)")); score += 1
        else:
            indicators.append(ind("Interest coverage", "🟢 CLEAR",
                f"Coverage = {ic['CY']:.2f}. Adequate.", f"{ref['gc']} para A2(c)"))
    else:
        indicators.append(ind("Interest coverage", "⚪ N/A", "No interest data.", f"{ref['gc']} para A2(c)"))

    de = ratio_map.get("Debt-to-Equity")
    if de and de["CY"] is not None:
        if de["CY"] > 3.0:
            indicators.append(ind("Excessive leverage (D/E > 3)", "🔴 CONCERN",
                f"D/E = {de['CY']:.2f}. High debt burden.", f"{ref['gc']} para A2(d)",
                "Verify covenant compliance and assess risk of lender withdrawal.")); score += 1
        else:
            indicators.append(ind("Excessive leverage", "🟢 CLEAR",
                f"D/E = {de['CY']:.2f}.", f"{ref['gc']} para A2(d)"))
    else:
        indicators.append(ind("Excessive leverage", "⚪ N/A", "Insufficient data.", f"{ref['gc']} para A2(d)"))

    rev_gr = ratio_map.get("Revenue Growth (%)")
    if rev_gr and rev_gr["CY"] is not None:
        if rev_gr["CY"] < -20:
            indicators.append(ind("Significant revenue decline > 20%", "🔴 CONCERN",
                f"Revenue fell {rev_gr['CY']:.1f}%.", f"{ref['gc']} para A2(a)",
                "Investigate customer loss, market contraction, pricing issues.")); score += 1
        elif rev_gr["CY"] < -10:
            indicators.append(ind("Revenue decline 10–20%", "🟡 MONITOR",
                f"Revenue fell {rev_gr['CY']:.1f}%.", f"{ref['gc']} para A2(a)")); score += 0
        else:
            indicators.append(ind("Significant revenue decline", "🟢 CLEAR",
                f"Revenue growth = {rev_gr['CY']:.1f}%.", f"{ref['gc']} para A2(a)"))
    else:
        indicators.append(ind("Revenue trend", "⚪ N/A", "Insufficient data.", f"{ref['gc']} para A2(a)"))

    if score == 0:
        risk, conclusion = "LOW", f"No significant going concern indicators. Document basis of preparation under {ref['gc']} para 9. No material uncertainty disclosure required at this stage."
    elif score <= 2:
        risk, conclusion = "MODERATE", f"Some indicators present. Under {ref['gc']} para 12–15, obtain management's written going concern assessment and supporting cash flow projections for 12 months."
    elif score <= 5:
        risk, conclusion = "HIGH", f"Multiple indicators identified. Perform extended procedures under {ref['gc']} para 16–17. Consider whether a material uncertainty disclosure is required in the financial statements."
    else:
        risk, conclusion = "CRITICAL", f"Serious doubt about the entity's ability to continue as a going concern. Consider modified opinion under {ref['gc']} para 21–24. Escalate to engagement partner immediately."

    return {"overall_risk": risk, "score": score, "conclusion": conclusion, "indicators": indicators}


# ── RISK SCORING SUMMARY ──────────────────────────────────────
def risk_summary(df_var, ratios, gc, bf):
    total_score = 0
    breakdown = []

    flagged_high = df_var[df_var["Flag"].str.contains("🔴", na=False)]
    flagged_mod  = df_var[df_var["Flag"].str.contains("🟡", na=False)]

    if len(flagged_high):
        pts = len(flagged_high) * 2
        total_score += pts
        breakdown.append(f"Variance — {len(flagged_high)} HIGH items (+{pts})")

    if len(flagged_mod):
        pts = len(flagged_mod) * 1
        total_score += pts
        breakdown.append(f"Variance — {len(flagged_mod)} MODERATE items (+{pts})")

    ratio_flags = [r for r in ratios if r.get("Available") and "🔴" in str(r.get("Flag", ""))]
    if ratio_flags:
        pts = len(ratio_flags) * 3
        total_score += pts
        breakdown.append(f"Ratios — {len(ratio_flags)} HIGH risk ratios (+{pts})")

    gc_pts = {"LOW": 0, "MODERATE": 2, "HIGH": 5, "CRITICAL": 8}.get(gc["overall_risk"], 0)
    if gc_pts:
        total_score += gc_pts
        breakdown.append(f"Going Concern — {gc['overall_risk']} (+{gc_pts})")

    if bf.get("sufficient"):
        if "🔴" in bf.get("risk_flag", ""):
            total_score += 3; breakdown.append("Benford's Law — HIGH deviation (+3)")
        elif "🟡" in bf.get("risk_flag", ""):
            total_score += 1; breakdown.append("Benford's Law — MODERATE deviation (+1)")

    if total_score <= 3:     level = "🟢 LOW"
    elif total_score <= 8:   level = "🟡 MODERATE"
    elif total_score <= 15:  level = "🔴 HIGH"
    else:                    level = "🚨 CRITICAL"

    return {"total_score": total_score, "level": level, "breakdown": breakdown}


# ── MAIN ENTRY POINT ──────────────────────────────────────────
def run_analysis(df, unit_info, mat_pct, standard):
    agg_data = aggregate(df)
    ratios   = calculate_ratios(agg_data, standard)
    df_var   = calculate_variance(df, mat_pct, unit_info, agg_data)
    gc       = going_concern(agg_data, ratios, standard)
    return {"agg": agg_data, "ratios": ratios, "variance": df_var, "going_concern": gc}


# ── BENFORD'S LAW ─────────────────────────────────────────────
def run_benford(amounts, label="Transactions"):
    amounts = [a for a in amounts if isinstance(a, (int, float)) and a > 0]
    n = len(amounts)

    if n < 20:
        return {"sufficient": False,
                "message": f"Only {n} valid amounts. Need at least 20 for meaningful analysis.",
                "risk_flag": "⚪ Insufficient Data"}

    counts = {d: 0 for d in range(1, 10)}
    for a in amounts:
        s = str(abs(a)).replace(".", "").lstrip("0")
        if s:
            fd = int(s[0])
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
            "Digit": d, "Observed Count": obs_count,
            "Observed (%)": round(obs_pct, 2), "Expected (%)": round(exp_pct, 2),
            "Deviation (pp)": round(dev_pp, 2), "Deviation (%)": round(dev_pct, 1),
        })

    if chi_sq > 20.09:
        flag = "🔴 HIGH — Significant deviation from Benford's Law"
        interp = f"χ² = {chi_sq:.2f} exceeds p<0.01 critical value (20.09). Significant deviation. May indicate manipulation, rounding to specific amounts, or fictitious entries. Extend fraud procedures under {STD_REFS['Ind AS']['fraud']}."
    elif chi_sq > 15.51:
        flag = "🟡 MODERATE — Notable deviation, review recommended"
        interp = f"χ² = {chi_sq:.2f} exceeds p<0.05 critical value (15.51). Some deviation present. Obtain explanation for unusual patterns and consider targeted transaction testing."
    elif chi_sq > 13.36:
        flag = "🟡 LOW-MODERATE — Minor deviation"
        interp = f"χ² = {chi_sq:.2f} slightly elevated. No strong manipulation indicator but note in working papers."
    else:
        flag = "🟢 LOW — Consistent with Benford's Law"
        interp = f"χ² = {chi_sq:.2f}. Distribution consistent with Benford's Law. No significant fraud indicator from first-digit analysis."

    return {"sufficient": True, "n": n, "label": label,
            "chi_square": round(chi_sq, 2), "digit_data": digit_data,
            "risk_flag": flag, "interpretation": interp}
