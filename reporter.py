"""
reporter.py — Excel Workpaper Generator
All openpyxl imports are lazy (inside generate_report) to avoid startup crashes.
"""

import pandas as pd
from datetime import date

P = {
    "navy":      "0D1B2A",
    "navy2":     "1E3A5F",
    "slate":     "334155",
    "red_bg":    "FEE2E2",
    "yel_bg":    "FEF9C3",
    "grn_bg":    "DCFCE7",
    "grey":      "F8FAFC",
    "white":     "FFFFFF",
    "border":    "E2E8F0",
    "text":      "0F172A",
    "sub":       "475569",
    "light":     "94A3B8",
}


def generate_report(df_variance, ratios, agg, gc, bf,
                    output_path, client_name, period, standard, unit_label):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    def bdr(color="E2E8F0"):
        s = Side(style="thin", color=color)
        return Border(left=s, right=s, top=s, bottom=s)

    def fill(color):
        return PatternFill("solid", fgColor=color)

    def hdr_cell(ws, row, col, val, bg=P["navy"], fg="FFFFFF", sz=10, bold=True, span=None):
        c = ws.cell(row=row, column=col, value=val)
        c.font = Font(name="Calibri", size=sz, bold=bold, color=fg)
        c.fill = fill(bg)
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1, wrap_text=True)
        c.border = bdr()
        if span:
            ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + span - 1)
        return c

    def data_cell(ws, row, col, val, bg=P["white"], bold=False, align="left", num_fmt=None, sz=10):
        c = ws.cell(row=row, column=col, value=val)
        c.font = Font(name="Calibri", size=sz, bold=bold, color=P["text"])
        c.fill = fill(bg)
        c.alignment = Alignment(horizontal=align, vertical="center", indent=1 if align == "left" else 0, wrap_text=True)
        c.border = bdr()
        if num_fmt:
            c.number_format = num_fmt
        return c

    def flag_bg(flag):
        s = str(flag)
        if "🔴" in s: return P["red_bg"]
        if "🟡" in s: return P["yel_bg"]
        if "🟢" in s: return P["grn_bg"]
        return P["white"]

    def row_height(ws, row, h):
        ws.row_dimensions[row].height = h

    def col_width(ws, col, w):
        ws.column_dimensions[get_column_letter(col)].width = w

    wb = Workbook()
    wb.remove(wb.active)

    # ══════════════════════════════════════════════════════════
    # SHEET 1 — COVER
    # ══════════════════════════════════════════════════════════
    ws1 = wb.create_sheet("Cover")
    ws1.sheet_view.showGridLines = False
    N = 6

    ws1.merge_cells(f"A1:F1")
    t = ws1["A1"]
    t.value = "AUDIT ANALYTICAL PROCEDURES WORKPAPER"
    t.font = Font(name="Calibri", size=18, bold=True, color="FFFFFF")
    t.fill = fill(P["navy"])
    t.alignment = Alignment(horizontal="center", vertical="center")
    t.border = bdr()
    row_height(ws1, 1, 44)

    ws1.merge_cells("A2:F2")
    s = ws1["A2"]
    s.value = f"{standard}  ·  ISA/SA 520 · 570 · 240 · 315  ·  Analytical Procedures Workpaper"
    s.font = Font(name="Calibri", size=10, italic=True, color=P["light"])
    s.fill = fill(P["navy2"])
    s.alignment = Alignment(horizontal="center", vertical="center")
    row_height(ws1, 2, 18)
    row_height(ws1, 3, 8)

    hdr_cell(ws1, 4, 1, "ENGAGEMENT DETAILS", bg=P["navy2"], span=N, sz=9)
    row_height(ws1, 4, 20)

    details = [
        ("Client", client_name),
        ("Period / Year End", period),
        ("Reporting Standard", standard),
        ("Amounts Presented In", unit_label),
        ("Report Date", date.today().strftime("%d %B %Y")),
        ("Prepared Using", "Audit AP Tool — Built by Purva Doshi"),
    ]
    for i, (lbl, val) in enumerate(details, 5):
        row_height(ws1, i, 18)
        lc = ws1.cell(row=i, column=1, value=lbl)
        lc.font = Font(name="Calibri", size=10, bold=True, color=P["text"])
        lc.fill = fill(P["grey"])
        lc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        lc.border = bdr()
        ws1.merge_cells(start_row=i, start_column=2, end_row=i, end_column=N)
        vc = ws1.cell(row=i, column=2, value=val)
        vc.font = Font(name="Calibri", size=10, color=P["text"])
        vc.fill = fill(P["white"])
        vc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        vc.border = bdr()

    row_height(ws1, 12, 10)
    hdr_cell(ws1, 13, 1, "RISK SUMMARY", bg=P["navy2"], span=N, sz=9)
    row_height(ws1, 13, 20)

    gc_risk = gc["overall_risk"]
    gc_score = gc["score"]
    flag_high = sum(1 for _, r in df_variance.iterrows() if "🔴" in str(r.get("Flag", "")))
    flag_mod  = sum(1 for _, r in df_variance.iterrows() if "🟡" in str(r.get("Flag", "")))
    available_ratios = sum(1 for r in ratios if r.get("Available"))
    flagged_ratios   = sum(1 for r in ratios if r.get("Available") and "🔴" in str(r.get("Flag", "")))

    summary_rows = [
        ("Going Concern Risk Level", gc_risk),
        ("Going Concern Score", f"{gc_score}/10"),
        ("Variance — HIGH Risk Items", str(flag_high)),
        ("Variance — MODERATE Items", str(flag_mod)),
        ("Ratios Calculated", str(available_ratios)),
        ("Ratios Flagged (HIGH)", str(flagged_ratios)),
        ("Benford's Law", bf.get("risk_flag", "Not Analysed")),
    ]
    for i, (lbl, val) in enumerate(summary_rows, 14):
        row_height(ws1, i, 18)
        lc = ws1.cell(row=i, column=1, value=lbl)
        lc.font = Font(name="Calibri", size=10, bold=True, color=P["text"])
        lc.fill = fill(P["grey"])
        lc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        lc.border = bdr()
        ws1.merge_cells(start_row=i, start_column=2, end_row=i, end_column=N)
        vc = ws1.cell(row=i, column=2, value=val)
        vc.font = Font(name="Calibri", size=10, color=P["text"])
        vc.fill = fill(flag_bg(val) if any(e in val for e in ["🔴","🟡","🟢"]) else P["white"])
        vc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        vc.border = bdr()

    for c, w in [(1, 32), (2, 40), (3, 10), (4, 10), (5, 10), (6, 10)]:
        col_width(ws1, c, w)

    # ══════════════════════════════════════════════════════════
    # SHEET 2 — RATIO ANALYSIS
    # ══════════════════════════════════════════════════════════
    ws2 = wb.create_sheet("Ratio Analysis")
    ws2.sheet_view.showGridLines = False
    hdrs2 = ["Ratio / Metric", "Formula", f"CY ({unit_label})", f"PY ({unit_label})", "YoY (%)", "Flag", "Auditor Note", "Ref"]
    widths2 = [32, 44, 14, 14, 11, 26, 52, 20]

    hdr_cell(ws2, 1, 1, f"RATIO ANALYSIS — {standard} | Analytical Procedures", span=8, sz=12)
    row_height(ws2, 1, 30)

    row_height(ws2, 2, 22)
    for col, (h, w) in enumerate(zip(hdrs2, widths2), 1):
        c = ws2.cell(row=2, column=col, value=h)
        c.font = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
        c.fill = fill(P["slate"])
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = bdr()
        col_width(ws2, col, w)

    sections = {0: "PROFITABILITY", 5: "RETURN METRICS", 7: "LIQUIDITY", 10: "LEVERAGE & SOLVENCY", 13: "EFFICIENCY"}
    avail = [r for r in ratios if r.get("Available")]
    data_row = 3
    for idx, r in enumerate(avail):
        if idx in sections:
            ws2.merge_cells(start_row=data_row, start_column=1, end_row=data_row, end_column=8)
            sc = ws2.cell(row=data_row, column=1, value=f"  ── {sections[idx]}")
            sc.font = Font(name="Calibri", size=9, bold=True, color=P["sub"])
            sc.fill = fill(P["grey"])
            sc.border = bdr()
            row_height(ws2, data_row, 16)
            data_row += 1

        bg = flag_bg(r.get("Flag", ""))
        row_height(ws2, data_row, 20)
        vals = [r.get("Ratio"), r.get("Formula"), r.get("CY"), r.get("PY"),
                r.get("YoY (%)"), r.get("Flag"), r.get("Note"), r.get("Ref")]
        for col, val in enumerate(vals, 1):
            c = ws2.cell(row=data_row, column=col, value=val)
            c.font = Font(name="Calibri", size=9, color=P["text"], bold=(col == 1))
            c.fill = fill(bg)
            c.border = bdr()
            c.alignment = Alignment(
                horizontal="right" if col in [3, 4, 5] else "left",
                vertical="center", wrap_text=True, indent=0 if col in [3,4,5] else 1
            )
            if col in [3, 4] and isinstance(val, float): c.number_format = "0.00"
            if col == 5 and isinstance(val, float): c.number_format = "+0.0;-0.0"
        data_row += 1

    ws2.freeze_panes = ws2.cell(row=3, column=1)

    # ══════════════════════════════════════════════════════════
    # SHEET 3 — ALL ACCOUNTS (VARIANCE)
    # ══════════════════════════════════════════════════════════
    ws3 = wb.create_sheet("Variance — All Accounts")
    ws3.sheet_view.showGridLines = False
    hdrs3 = ["Account Name", "Category", f"CY ({unit_label})", f"PY ({unit_label})", f"Change ({unit_label})", "Change (%)", "Flag"]
    widths3 = [40, 24, 16, 16, 16, 13, 28]

    hdr_cell(ws3, 1, 1, f"VARIANCE ANALYSIS — All Accounts | {standard}", span=7, sz=11)
    row_height(ws3, 1, 28)
    row_height(ws3, 2, 20)
    for col, (h, w) in enumerate(zip(hdrs3, widths3), 1):
        c = ws3.cell(row=2, column=col, value=h)
        c.font = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
        c.fill = fill(P["slate"])
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = bdr()
        col_width(ws3, col, w)

    for i, (_, row) in enumerate(df_variance.iterrows(), 3):
        flag = str(row.get("Flag", ""))
        bg = flag_bg(flag)
        if bg == P["white"]: bg = P["grey"] if i % 2 == 0 else P["white"]
        row_height(ws3, i, 18)
        vals = [row["Account Name"], row["Category"], row["CY Amount"], row["PY Amount"],
                row["Change"], row["Change (%)"], flag]
        for col, val in enumerate(vals, 1):
            c = ws3.cell(row=i, column=col, value=val if not (col == 6 and val == 999.0) else "New")
            c.font = Font(name="Calibri", size=9, color=P["text"])
            c.fill = fill(bg)
            c.border = bdr()
            c.alignment = Alignment(horizontal="right" if col in [3,4,5,6] else "left", vertical="center", indent=1 if col in [1,2,7] else 0)
            if col in [3, 4, 5] and isinstance(val, (int, float)): c.number_format = "#,##0.00"
            if col == 6 and isinstance(val, float) and val != 999.0: c.number_format = "+0.0;-0.0"

    ws3.freeze_panes = ws3.cell(row=3, column=1)

    # ══════════════════════════════════════════════════════════
    # SHEET 4 — FLAGGED ITEMS
    # ══════════════════════════════════════════════════════════
    ws4 = wb.create_sheet("Flagged Items")
    ws4.sheet_view.showGridLines = False
    flagged_df = df_variance[df_variance["Flag"].str.contains("🔴|🟡", na=False)]

    hdr_cell(ws4, 1, 1, f"FLAGGED ITEMS — Requires Auditor Attention | {standard}", span=7, sz=11)
    row_height(ws4, 1, 28)
    row_height(ws4, 2, 20)
    for col, (h, w) in enumerate(zip(hdrs3, widths3), 1):
        c = ws4.cell(row=2, column=col, value=h)
        c.font = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
        c.fill = fill(P["slate"])
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = bdr()
        col_width(ws4, col, w)

    if flagged_df.empty:
        ws4.merge_cells("A3:G3")
        nc = ws4["A3"]
        nc.value = "✅ No flagged items — all accounts within materiality threshold."
        nc.font = Font(name="Calibri", size=10, color="14532D")
        nc.fill = fill(P["grn_bg"])
        nc.border = bdr()
    else:
        for i, (_, row) in enumerate(flagged_df.iterrows(), 3):
            flag = str(row.get("Flag", ""))
            bg = flag_bg(flag)
            row_height(ws4, i, 18)
            vals = [row["Account Name"], row["Category"], row["CY Amount"], row["PY Amount"],
                    row["Change"], row["Change (%)"], flag]
            for col, val in enumerate(vals, 1):
                c = ws4.cell(row=i, column=col, value=val if not (col == 6 and val == 999.0) else "New")
                c.font = Font(name="Calibri", size=9, color=P["text"])
                c.fill = fill(bg)
                c.border = bdr()
                c.alignment = Alignment(horizontal="right" if col in [3,4,5,6] else "left", vertical="center", indent=1 if col in [1,2,7] else 0)
                if col in [3, 4, 5] and isinstance(val, (int, float)): c.number_format = "#,##0.00"
                if col == 6 and isinstance(val, float) and val != 999.0: c.number_format = "+0.0;-0.0"

    ws4.freeze_panes = ws4.cell(row=3, column=1)

    # ══════════════════════════════════════════════════════════
    # SHEET 5 — GOING CONCERN
    # ══════════════════════════════════════════════════════════
    ws5 = wb.create_sheet("Going Concern")
    ws5.sheet_view.showGridLines = False
    gc_color_map = {"CRITICAL": "991B1B", "HIGH": "C2410C", "MODERATE": "92400E", "LOW": "14532D"}
    gc_bg_map    = {"CRITICAL": "FEE2E2", "HIGH": "FEF3C7", "MODERATE": "FEF9C3", "LOW": "DCFCE7"}
    risk = gc["overall_risk"]

    hdr_cell(ws5, 1, 1, f"GOING CONCERN ASSESSMENT — {standard}", span=5, sz=12)
    row_height(ws5, 1, 30)
    row_height(ws5, 2, 8)
    row_height(ws5, 3, 34)
    ws5.merge_cells("A3:E3")
    vc = ws5["A3"]
    vc.value = f"OVERALL: {risk}  ·  Risk Score: {gc['score']}/10"
    vc.font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    vc.fill = fill(gc_color_map.get(risk, P["navy"]))
    vc.alignment = Alignment(horizontal="center", vertical="center")
    vc.border = bdr()

    row_height(ws5, 4, 52)
    ws5.merge_cells("A4:E4")
    cc = ws5["A4"]
    cc.value = gc["conclusion"]
    cc.font = Font(name="Calibri", size=10, color=P["text"])
    cc.fill = fill(gc_bg_map.get(risk, P["grey"]))
    cc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
    cc.border = bdr()

    row_height(ws5, 5, 10)
    hdr_cell(ws5, 6, 1, "DETAILED INDICATORS", bg=P["slate"], span=5, sz=9)
    row_height(ws5, 6, 20)
    for col, h in enumerate(["Indicator", "Status", "Finding", "", "Reference"], 1):
        c = ws5.cell(row=7, column=col, value=h)
        c.font = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
        c.fill = fill(P["slate"])
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = bdr()
    row_height(ws5, 7, 20)

    for i, ind in enumerate(gc["indicators"], 8):
        flag = str(ind.get("Status", ""))
        bg = flag_bg(flag)
        ws5.merge_cells(start_row=i, start_column=3, end_row=i, end_column=4)
        row_height(ws5, i, 34)
        for col, val in enumerate([ind["Indicator"], ind["Status"], ind["Finding"], "", ind["Reference"]], 1):
            if col == 4: continue
            c = ws5.cell(row=i, column=col, value=val)
            c.font = Font(name="Calibri", size=9, color=P["text"], bold=(col == 1))
            c.fill = fill(bg)
            c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
            c.border = bdr()

    for c, w in [(1, 40), (2, 22), (3, 48), (4, 4), (5, 28)]:
        col_width(ws5, c, w)

    # ══════════════════════════════════════════════════════════
    # SHEET 6 — BENFORD'S LAW
    # ══════════════════════════════════════════════════════════
    ws6 = wb.create_sheet("Benford's Law")
    ws6.sheet_view.showGridLines = False
    hdr_cell(ws6, 1, 1, "BENFORD'S LAW ANALYSIS — ISA/SA 240 Fraud Risk Indicator", span=7, sz=12)
    row_height(ws6, 1, 30)

    if not bf.get("sufficient"):
        ws6.merge_cells("A2:G2")
        nc = ws6["A2"]
        nc.value = f"⚠ NOT ANALYSED — {bf.get('message', 'No transaction data uploaded.')}"
        nc.font = Font(name="Calibri", size=10, color=P["text"])
        nc.fill = fill(P["yel_bg"])
        nc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
        nc.border = bdr()
        row_height(ws6, 2, 48)
    else:
        flag = bf.get("risk_flag", "")
        ws6.merge_cells("A2:G2")
        rc = ws6["A2"]
        rc.value = f"{flag}  ·  {bf['n']:,} amounts analysed  ·  Chi-Square: {bf['chi_square']:.2f}"
        rc.font = Font(name="Calibri", size=11, bold=True, color=P["text"])
        rc.fill = fill(flag_bg(flag))
        rc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        rc.border = bdr()
        row_height(ws6, 2, 26)

        ws6.merge_cells("A3:G3")
        ic = ws6["A3"]
        ic.value = bf.get("interpretation", "")
        ic.font = Font(name="Calibri", size=10, color=P["text"])
        ic.fill = fill(P["grey"])
        ic.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
        ic.border = bdr()
        row_height(ws6, 3, 48)

        row_height(ws6, 4, 10)
        hdr_cell(ws6, 5, 1, "DIGIT-BY-DIGIT ANALYSIS", bg=P["slate"], span=7, sz=9)
        row_height(ws6, 5, 20)
        bf_hdrs = ["Digit", "Count", "Observed (%)", "Expected (%)", "Deviation (pp)", "Deviation (%)", "Flag"]
        bf_widths = [10, 12, 14, 22, 16, 14, 22]
        for col, (h, w) in enumerate(zip(bf_hdrs, bf_widths), 1):
            c = ws6.cell(row=6, column=col, value=h)
            c.font = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
            c.fill = fill(P["slate"])
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = bdr()
            col_width(ws6, col, w)
        row_height(ws6, 6, 20)

        for i, drow in enumerate(bf["digit_data"], 7):
            dev = abs(drow["Deviation (pp)"])
            if dev > 5:   df_bg, df_flag = P["red_bg"], "🔴 Investigate"
            elif dev > 2: df_bg, df_flag = P["yel_bg"], "🟡 Review"
            else:         df_bg, df_flag = P["grn_bg"], "🟢 OK"
            row_height(ws6, i, 18)
            for col, val in enumerate([drow["Digit"], drow["Observed Count"], drow["Observed (%)"],
                                        drow["Expected (%)"], drow["Deviation (pp)"], drow["Deviation (%)"], df_flag], 1):
                c = ws6.cell(row=i, column=col, value=val)
                c.font = Font(name="Calibri", size=9, color=P["text"])
                c.fill = fill(df_bg)
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.border = bdr()
                if col in [3, 4]: c.number_format = "0.00"
                if col == 5: c.number_format = "+0.00;-0.00"

        hdr_cell(ws6, 17, 1, f"Critical values (df=8): p<0.10 → 13.36  |  p<0.05 → 15.51  |  p<0.01 → 20.09  |  Your statistic: {bf['chi_square']:.2f}", bg=P["slate"], span=7, sz=9)
        row_height(ws6, 17, 20)

    # ══════════════════════════════════════════════════════════
    # SHEET 7 — AUDIT NOTES
    # ══════════════════════════════════════════════════════════
    ws7 = wb.create_sheet("Audit Notes")
    ws7.sheet_view.showGridLines = False
    hdr_cell(ws7, 1, 1, "AUDIT NOTES & STANDARD REFERENCES", span=4, sz=12)
    row_height(ws7, 1, 28)

    from engine import STD_REFS
    refs = STD_REFS.get(standard, STD_REFS["IFRS"])
    notes = [
        ("Standard Applied", standard),
        ("Analytical Procedures", refs["ap"]),
        ("Going Concern", refs["gc"]),
        ("Fraud Risk", refs["fraud"]),
        ("Risk Assessment", refs["risk"]),
        ("Revenue Recognition", refs["revenue"]),
        ("Impairment", refs["impairment"]),
        ("Leases", refs["leases"]),
        ("", ""),
        ("Materiality basis", "Threshold auto-scaled to detected amount units"),
        ("Variance threshold", "HIGH = >1.5x threshold % AND above absolute threshold"),
        ("Going concern scoring", "Score 0 = LOW · 1-2 = MODERATE · 3-5 = HIGH · 6+ = CRITICAL"),
        ("Benford's Law", "Chi-square test df=8 · p<0.05 → notable deviation · p<0.01 → significant"),
        ("", ""),
        ("Built by", "Purva Doshi"),
        ("LinkedIn", "https://www.linkedin.com/in/purvadoshi26/"),
        ("Tool", "Audit Analytical Procedures Tool"),
        ("Generated", date.today().strftime("%d %B %Y")),
    ]
    for i, (lbl, val) in enumerate(notes, 2):
        row_height(ws7, i, 18)
        if not lbl:
            row_height(ws7, i, 8)
            continue
        lc = ws7.cell(row=i, column=1, value=lbl)
        lc.font = Font(name="Calibri", size=10, bold=True, color=P["text"])
        lc.fill = fill(P["grey"])
        lc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        lc.border = bdr()
        ws7.merge_cells(start_row=i, start_column=2, end_row=i, end_column=4)
        vc = ws7.cell(row=i, column=2, value=val)
        vc.font = Font(name="Calibri", size=10, color=P["text"])
        vc.fill = fill(P["white"])
        vc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        vc.border = bdr()

    for c, w in [(1, 30), (2, 60), (3, 10), (4, 10)]:
        col_width(ws7, c, w)

    wb.save(output_path)
    return output_path
