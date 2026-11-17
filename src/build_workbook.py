#!/usr/bin/env python3
"""P21 — build the deliverable workbook for PakBazaar (SAMPLE):

  Dashboard              — KPIs (orders, returns, return rate, gross
                           revenue, refunded revenue, return shipping,
                           money lost = refunded + shipping, net
                           revenue, revenue kept, refund vs
                           replacement counts), worst/best category and
                           channel, top return reason, worst products,
                           worst month, worst/best category return rate
                           via MAXIFS/MINIFS (written with the _xlfn.
                           prefix so LibreOffice and Excel calculate
                           them), monthly return-rate trend + chart —
                           all live formulas over Returns
  Returns                — all 1,200 orders (values for the source
                           columns, live formulas for is_returned and
                           money_lost_pkr)
  By Category            — 5 categories: orders, returns, return rate,
                           gross, refunded, shipping, money lost, net,
                           kept, average order value — all live
                           COUNTIFS / SUMIFS over Returns
  By Channel             — same, 4 channels
  By Reason              — 6 return reasons: returns, share of
                           returns, refunded, shipping, money lost,
                           share of money lost
  Products               — 20 products: orders, returns, return rate,
                           gross, refunded, shipping, money lost, net,
                           kept, top reason
  Monthly Trend          — 6 months, same as By Category plus
                           month-on-month return-rate change
  Method                 — how to read it, the return-rate and
                           money-lost rules, their limits, how to
                           re-verify

    python3 src/build_workbook.py
"""
import csv
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill

PROJECT = Path(__file__).resolve().parent.parent
RAW = PROJECT / "data" / "raw"
OUT = PROJECT / "outputs" / "returns_report_SAMPLE.xlsx"

CATEGORIES = ["Fashion", "Electronics", "Home & Kitchen", "Beauty",
              "Sports"]
CHANNELS = ["Website", "Daraz", "Facebook", "WhatsApp"]
MONTHS = [f"2026-{m:02d}" for m in range(1, 7)]
REASONS = ["Wrong size", "Defective/damaged", "Wrong item sent",
           "Late delivery", "Quality not as expected", "Changed mind"]
HDR_FILL = PatternFill("solid", fgColor="1F4E79")
HDR_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, size=14, color="1F4E79")
COUNT_FMT = "#,##0"
PKR_FMT = "#,##0"
PCT_FMT = "0.00%"
RATE_FMT = "0.00"
AVG_FMT = "#,##0.00"


def style_header(ws, ncols, row=1):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT


def main():
    raw = list(csv.DictReader(
        open(RAW / "orders_SAMPLE.csv", encoding="utf-8")))
    n = len(raw)
    last = n + 1  # last data row on Returns (row 1 = header)

    # top reason per product, from the raw file (labels, not totals)
    top_reason = {}
    for r in raw:
        if r["return_status"] == "Returned":
            top_reason.setdefault(r["product_sku"], {})
            d = top_reason[r["product_sku"]]
            d[r["return_reason"]] = d.get(r["return_reason"], 0) + 1
    top_reason = {k: max(v, key=lambda x: v[x])
                  for k, v in top_reason.items()}

    wb = Workbook()

    # ── Returns (values for source columns, formulas for derived) ──
    rs = wb.active
    rs.title = "Returns"
    rs.append(["order_id", "order_date", "month", "category",
               "channel", "product_name", "product_sku", "quantity",
               "order_value_pkr", "return_status", "return_reason",
               "resolution", "refund_amount_pkr",
               "return_shipping_cost_pkr", "is_returned",
               "money_lost_pkr"])
    for r in raw:
        rs.append([r["order_id"], r["order_date"], r["month"],
                   r["category"], r["channel"], r["product_name"],
                   r["product_sku"], int(r["quantity"]),
                   int(r["order_value_pkr"]), r["return_status"],
                   r["return_reason"], r["resolution"],
                   int(r["refund_amount_pkr"]),
                   int(r["return_shipping_cost_pkr"])])
    for i in range(2, last + 1):
        rs.cell(row=i, column=15).value = f'=IF(J{i}="Returned",1,0)'
        rs.cell(row=i, column=16).value = f"=IF(O{i}=1,M{i}+N{i},0)"
    style_header(rs, 16)
    rs.freeze_panes = "A2"
    widths = {"A": 10, "B": 12, "C": 9, "D": 15, "E": 10, "F": 26,
              "G": 10, "H": 9, "I": 15, "J": 12, "K": 22, "L": 11,
              "M": 15, "N": 20, "O": 11, "P": 14}
    for col, w in widths.items():
        rs.column_dimensions[col].width = w
    for r in range(2, last + 1):
        for col in (8, 9, 13, 14, 16):
            rs.cell(row=r, column=col).number_format = PKR_FMT

    # ── shared group sheet (category / channel / month) ─────────────
    # Columns: A label, B orders, C returns, D return rate,
    # E gross, F refunded, G shipping, H money lost, I net,
    # J kept, K avg order value
    def group_sheet(name, title, labels, key_letter, nrows):
        ws = wb.create_sheet(name)
        ws["A1"] = title
        ws["A1"].font = TITLE_FONT
        headers = ["Label", "Orders", "Returns", "Return rate (%)",
                   "Gross revenue (PKR)", "Refunded (PKR)",
                   "Return shipping (PKR)", "Money lost (PKR)",
                   "Net revenue (PKR)",
                   "Revenue kept (PKR)", "Avg order value (PKR)"]
        for c, h in enumerate(headers, start=1):
            ws.cell(row=3, column=c, value=h)
        style_header(ws, 11, row=3)
        ws.freeze_panes = "A4"
        for i, label in enumerate(labels):
            r = 4 + i
            ws.cell(row=r, column=1, value=label)
            ws.cell(row=r, column=2).value = (
                f"=COUNTIFS(Returns!${key_letter}$2:${key_letter}"
                f"${last},$A{r})")
            ws.cell(row=r, column=3).value = (
                f"=SUMIFS(Returns!$O$2:$O${last},"
                f"Returns!${key_letter}$2:${key_letter}${last},$A{r})")
            ws.cell(row=r, column=4).value = (
                f"=IF(B{r}=0,0,100*C{r}/B{r})")
            ws.cell(row=r, column=4).number_format = RATE_FMT
            ws.cell(row=r, column=5).value = (
                f"=SUMIFS(Returns!$I$2:$I${last},"
                f"Returns!${key_letter}$2:${key_letter}${last},$A{r})")
            ws.cell(row=r, column=6).value = (
                f"=SUMIFS(Returns!$M$2:$M${last},"
                f"Returns!${key_letter}$2:${key_letter}${last},$A{r})")
            ws.cell(row=r, column=7).value = (
                f"=SUMIFS(Returns!$N$2:$N${last},"
                f"Returns!${key_letter}$2:${key_letter}${last},$A{r})")
            ws.cell(row=r, column=8).value = f"=F{r}+G{r}"
            ws.cell(row=r, column=9).value = f"=E{r}-F{r}"
            ws.cell(row=r, column=10).value = f"=E{r}-H{r}"
            ws.cell(row=r, column=11).value = (
                f"=IF(B{r}=0,0,E{r}/B{r})")
            ws.cell(row=r, column=11).number_format = AVG_FMT
            for col in (2, 3, 5, 6, 7, 8, 9, 10):
                ws.cell(row=r, column=col).number_format = PKR_FMT
        t = 4 + nrows  # total row
        ws.cell(row=t, column=1, value="TOTAL").font = Font(bold=True)
        ws.cell(row=t, column=2).value = f"=SUM(B4:B{t-1})"
        ws.cell(row=t, column=3).value = f"=SUM(C4:C{t-1})"
        ws.cell(row=t, column=4).value = (
            f"=IF(B{t}=0,0,100*C{t}/B{t})")
        ws.cell(row=t, column=4).number_format = RATE_FMT
        for col, letter in ((5, "E"), (6, "F"), (7, "G")):
            ws.cell(row=t, column=col).value = (
                f"=SUM({letter}4:{letter}{t-1})")
        ws.cell(row=t, column=8).value = f"=F{t}+G{t}"
        ws.cell(row=t, column=9).value = f"=E{t}-F{t}"
        ws.cell(row=t, column=10).value = f"=E{t}-H{t}"
        ws.cell(row=t, column=11).value = (
            f"=IF(B{t}=0,0,E{t}/B{t})")
        ws.cell(row=t, column=11).number_format = AVG_FMT
        for col in (2, 3, 5, 6, 7, 8, 9, 10):
            ws.cell(row=t, column=col).number_format = PKR_FMT
        for col, w in zip("ABCDEFGHIJK",
                          [18, 9, 9, 14, 17, 13, 18, 14, 15, 16, 17]):
            ws.column_dimensions[col].width = w
        return ws

    group_sheet("By Category",
                "Return rate by category — live formulas over Returns",
                CATEGORIES, "D", len(CATEGORIES))
    group_sheet("By Channel",
                "Return rate by channel — live formulas over Returns",
                CHANNELS, "E", len(CHANNELS))
    mo = group_sheet("Monthly Trend",
                     "Monthly return rate — live formulas over Returns",
                     MONTHS, "C", len(MONTHS))
    # month-on-month return-rate change on the Monthly Trend sheet
    mo.cell(row=3, column=12, value="Return-rate change (pts)")
    style_header(mo, 12, row=3)
    mo.cell(row=4, column=12, value="—")
    for r in range(5, 10):
        mo.cell(row=r, column=12).value = f"=D{r}-D{r-1}"
        mo.cell(row=r, column=12).number_format = RATE_FMT
    mo.column_dimensions["L"].width = 22

    # ── By Reason ────────────────────────────────────────────────────
    rr = wb.create_sheet("By Reason")
    rr["A1"] = ("Why customers send it back — live formulas over "
                "Returns; money lost = refunded revenue + return "
                "shipping cost")
    rr["A1"].font = TITLE_FONT
    for c, h in enumerate(["Reason", "Returns", "Share of returns",
                           "Refunded (PKR)", "Return shipping (PKR)",
                           "Money lost (PKR)",
                           "Share of money lost"], start=1):
        rr.cell(row=3, column=c, value=h)
    style_header(rr, 7, row=3)
    rr.freeze_panes = "A4"
    for i, reason in enumerate(REASONS):
        r = 4 + i
        rr.cell(row=r, column=1, value=reason)
        rr.cell(row=r, column=2).value = (
            f"=COUNTIFS(Returns!$K$2:$K${last},$A{r})")
        rr.cell(row=r, column=2).number_format = PKR_FMT
        rr.cell(row=r, column=3).value = (
            f"=IF($B$10=0,0,B{r}/$B$10)")
        rr.cell(row=r, column=3).number_format = PCT_FMT
        rr.cell(row=r, column=4).value = (
            f"=SUMIFS(Returns!$M$2:$M${last},"
            f"Returns!$K$2:$K${last},$A{r})")
        rr.cell(row=r, column=5).value = (
            f"=SUMIFS(Returns!$N$2:$N${last},"
            f"Returns!$K$2:$K${last},$A{r})")
        rr.cell(row=r, column=6).value = f"=D{r}+E{r}"
        rr.cell(row=r, column=7).value = (
            f"=IF($F$10=0,0,F{r}/$F$10)")
        rr.cell(row=r, column=7).number_format = PCT_FMT
        for col in (4, 5, 6):
            rr.cell(row=r, column=col).number_format = PKR_FMT
    rr.cell(row=10, column=1, value="TOTAL").font = Font(bold=True)
    rr.cell(row=10, column=2).value = "=SUM(B4:B9)"
    rr.cell(row=10, column=2).number_format = PKR_FMT
    rr.cell(row=10, column=3).value = "=SUM(C4:C9)"
    rr.cell(row=10, column=3).number_format = PCT_FMT
    rr.cell(row=10, column=4).value = "=SUM(D4:D9)"
    rr.cell(row=10, column=5).value = "=SUM(E4:E9)"
    rr.cell(row=10, column=6).value = "=D10+E10"
    rr.cell(row=10, column=7).value = "=SUM(G4:G9)"
    rr.cell(row=10, column=7).number_format = PCT_FMT
    for col in (4, 5, 6):
        rr.cell(row=10, column=col).number_format = PKR_FMT
    for col, w in zip("ABCDEFG", [24, 9, 15, 15, 19, 15, 18]):
        rr.column_dimensions[col].width = w

    # ── Products ─────────────────────────────────────────────────────
    pr = wb.create_sheet("Products")
    pr["A1"] = ("Top problem products — live formulas over Returns; "
                "money lost = refunded revenue + return shipping cost")
    pr["A1"].font = TITLE_FONT
    for c, h in enumerate(["Product", "SKU", "Category", "Orders",
                           "Returns", "Return rate (%)",
                           "Gross revenue (PKR)", "Refunded (PKR)",
                           "Return shipping (PKR)", "Money lost (PKR)",
                           "Net revenue (PKR)",
                           "Revenue kept (PKR)", "Top reason"],
                          start=1):
        pr.cell(row=3, column=c, value=h)
    style_header(pr, 13, row=3)
    pr.freeze_panes = "A4"
    products = sorted({(r["product_name"], r["product_sku"],
                        r["category"]) for r in raw})
    for i, (name, sku, cat) in enumerate(products):
        r = 4 + i
        pr.cell(row=r, column=1, value=name)
        pr.cell(row=r, column=2, value=sku)
        pr.cell(row=r, column=3, value=cat)
        pr.cell(row=r, column=4).value = (
            f"=COUNTIFS(Returns!$G$2:$G${last},$B{r})")
        pr.cell(row=r, column=5).value = (
            f"=SUMIFS(Returns!$O$2:$O${last},"
            f"Returns!$G$2:$G${last},$B{r})")
        pr.cell(row=r, column=6).value = (
            f"=IF(D{r}=0,0,100*E{r}/D{r})")
        pr.cell(row=r, column=6).number_format = RATE_FMT
        pr.cell(row=r, column=7).value = (
            f"=SUMIFS(Returns!$I$2:$I${last},"
            f"Returns!$G$2:$G${last},$B{r})")
        pr.cell(row=r, column=8).value = (
            f"=SUMIFS(Returns!$M$2:$M${last},"
            f"Returns!$G$2:$G${last},$B{r})")
        pr.cell(row=r, column=9).value = (
            f"=SUMIFS(Returns!$N$2:$N${last},"
            f"Returns!$G$2:$G${last},$B{r})")
        pr.cell(row=r, column=10).value = f"=H{r}+I{r}"
        pr.cell(row=r, column=11).value = f"=G{r}-H{r}"
        pr.cell(row=r, column=12).value = f"=G{r}-J{r}"
        pr.cell(row=r, column=13, value=top_reason.get(sku, "—"))
        for col in (4, 5, 7, 8, 9, 10, 11, 12):
            pr.cell(row=r, column=col).number_format = PKR_FMT
    t = 24  # total row (20 products, rows 4..23)
    pr.cell(row=t, column=1, value="TOTAL").font = Font(bold=True)
    pr.cell(row=t, column=4).value = "=SUM(D4:D23)"
    pr.cell(row=t, column=5).value = "=SUM(E4:E23)"
    pr.cell(row=t, column=6).value = (
        "=IF(D24=0,0,100*E24/D24)")
    pr.cell(row=t, column=6).number_format = RATE_FMT
    for col, letter in ((7, "G"), (8, "H"), (9, "I")):
        pr.cell(row=t, column=col).value = (
            f"=SUM({letter}4:{letter}23)")
    pr.cell(row=t, column=10).value = "=H24+I24"
    pr.cell(row=t, column=11).value = "=G24-H24"
    pr.cell(row=t, column=12).value = "=G24-J24"
    for col in (4, 5, 7, 8, 9, 10, 11, 12):
        pr.cell(row=t, column=col).number_format = PKR_FMT
    for col, w in zip("ABCDEFGHIJKLM",
                      [28, 10, 15, 8, 8, 14, 17, 13, 18, 14, 15, 16,
                       22]):
        pr.column_dimensions[col].width = w

    # ── Dashboard ────────────────────────────────────────────────────
    dash = wb.create_sheet("Dashboard", 0)
    dash["A1"] = ("E-commerce Returns & Refunds Analysis — PakBazaar "
                  "(SAMPLE)")
    dash["A1"].font = TITLE_FONT
    dash["A2"] = ("6 months, 2026-01-01 → 2026-06-30 · every number "
                  "below is a live formula over the Returns sheet — "
                  "change the data and the report rebuilds itself. "
                  "Return rate = 100 x returned orders / total "
                  "orders; money lost = refunded revenue + return "
                  "shipping cost; see the Method sheet.")
    dash["A2"].font = Font(italic=True, size=10, color="595959")
    dash["A4"] = "Key metrics"
    dash["A4"].font = Font(bold=True, size=12, color="1F4E79")
    kpis = [
        ("Total orders", f"=COUNTA(Returns!A2:A{last})", COUNT_FMT),
        ("Returned orders", f"=SUM(Returns!O2:O{last})", COUNT_FMT),
        ("Return rate (%)", "=IF(B5=0,0,100*B6/B5)", RATE_FMT),
        ("Gross revenue (PKR)", f"=SUM(Returns!I2:I{last})", PKR_FMT),
        ("Refunded revenue (PKR)", f"=SUM(Returns!M2:M{last})",
         PKR_FMT),
        ("Return shipping cost (PKR)", f"=SUM(Returns!N2:N{last})",
         PKR_FMT),
        ("Money lost (PKR)", "=B9+B10", PKR_FMT),
        ("Net revenue (PKR)", "=B8-B9", PKR_FMT),
        ("Revenue kept after all return costs (PKR)", "=B8-B11",
         PKR_FMT),
        ("Refunded returns (count)",
         '=COUNTIFS(Returns!$L$2:$L$' + str(last) + ',"Refunded")',
         COUNT_FMT),
        ("Replaced returns (count)",
         '=COUNTIFS(Returns!$L$2:$L$' + str(last) + ',"Replaced")',
         COUNT_FMT),
        ("Refund share of returns", "=IF(B6=0,0,B14/B6)", PCT_FMT),
        ("Average order value (PKR)", "=IF(B5=0,0,B8/B5)", AVG_FMT),
        ("Refunded share of gross revenue", "=IF(B8=0,0,B9/B8)",
         PCT_FMT),
        ("Money lost share of gross revenue", "=IF(B8=0,0,B11/B8)",
         PCT_FMT),
    ]
    for i, (label, formula, fmt) in enumerate(kpis):
        r = 5 + i
        dash.cell(row=r, column=1, value=label)
        c = dash.cell(row=r, column=2, value=formula)
        if fmt:
            c.number_format = fmt
        c.font = Font(bold=True)
    # KPI rows: 5 orders, 6 returns, 7 rate, 8 gross, 9 refunded,
    # 10 shipping, 11 money lost, 12 net, 13 kept, 14 refunded count,
    # 15 replaced count, 16 refund share, 17 avg order, 18 refunded
    # share of gross, 19 money-lost share of gross
    dash.cell(row=7, column=3).value = (
        "Returned orders / total orders x 100.")
    dash.cell(row=11, column=3).value = (
        "Refunded revenue + return shipping cost.")
    dash.cell(row=12, column=3).value = "Gross revenue - refunded."
    dash.cell(row=13, column=3).value = (
        "Gross revenue - money lost.")

    def best_worst(row, label, sheet, label_range, value_range,
                   func):
        dash.cell(row=row, column=1, value=label).font = Font(bold=True)
        dash.cell(row=row, column=2).value = (
            f"=INDEX({sheet}!{label_range},MATCH({func}({sheet}!"
            f"{value_range}),{sheet}!{value_range},0))")
        dash.cell(row=row, column=2).font = Font(bold=True)
        dash.cell(row=row, column=3).value = (
            f"={func}({sheet}!{value_range})")
        dash.cell(row=row, column=3).number_format = RATE_FMT

    best_worst(21, "Worst category (return rate)", "'By Category'",
               "A4:A8", "D4:D8", "MAX")
    best_worst(22, "Best category (return rate)", "'By Category'",
               "A4:A8", "D4:D8", "MIN")
    best_worst(23, "Worst channel (return rate)", "'By Channel'",
               "A4:A7", "D4:D7", "MAX")
    best_worst(24, "Best channel (return rate)", "'By Channel'",
               "A4:A7", "D4:D7", "MIN")
    best_worst(25, "Top return reason (count)", "'By Reason'",
               "A4:A9", "B4:B9", "MAX")
    dash.cell(row=25, column=3).number_format = COUNT_FMT
    best_worst(26, "Worst product (return rate)", "Products",
               "A4:A23", "F4:F23", "MAX")
    dash.cell(row=27, column=1,
             value="Worst product (money lost)").font = Font(bold=True)
    dash.cell(row=27, column=2).value = (
        "=INDEX(Products!A4:A23,MATCH(MAX(Products!J4:J23),"
        "Products!J4:J23,0))")
    dash.cell(row=27, column=2).font = Font(bold=True)
    dash.cell(row=27, column=3).value = "=MAX(Products!J4:J23)"
    dash.cell(row=27, column=3).number_format = PKR_FMT
    best_worst(28, "Worst month (return rate)", "'Monthly Trend'",
               "A4:A9", "D4:D9", "MAX")

    # worst / best category return rate among categories with 50+
    # orders (MAXIFS/MINIFS written with the _xlfn. prefix)
    dash.cell(row=30, column=1,
             value="Worst category return rate (50+ orders)").font = (
                 Font(bold=True))
    dash.cell(row=30, column=2).value = (
        '=_xlfn.MAXIFS(\'By Category\'!$D$4:$D$8,'
        '\'By Category\'!$B$4:$B$8,">=50")')
    dash.cell(row=30, column=2).number_format = RATE_FMT
    dash.cell(row=30, column=3).value = (
        "=INDEX('By Category'!$A$4:$A$8,MATCH(B30,"
        "'By Category'!$D$4:$D$8,0))")
    dash.cell(row=31, column=1,
             value="Best category return rate (50+ orders)").font = (
                 Font(bold=True))
    dash.cell(row=31, column=2).value = (
        '=_xlfn.MINIFS(\'By Category\'!$D$4:$D$8,'
        '\'By Category\'!$B$4:$B$8,">=50")')
    dash.cell(row=31, column=2).number_format = RATE_FMT
    dash.cell(row=31, column=3).value = (
        "=INDEX('By Category'!$A$4:$A$8,MATCH(B31,"
        "'By Category'!$D$4:$D$8,0))")

    tr = 33  # monthly trend title row
    dash.cell(row=tr, column=1, value="Monthly return-rate trend").font = (
        Font(bold=True, size=12, color="1F4E79"))
    headers = ["Month", "Orders", "Returns", "Return rate (%)"]
    for c, h in enumerate(headers, start=1):
        dash.cell(row=tr + 1, column=c, value=h)
    style_header(dash, 4, row=tr + 1)
    for i, m in enumerate(MONTHS):
        r = tr + 2 + i  # rows 35..40
        dash.cell(row=r, column=1, value=m)
        dash.cell(row=r, column=2).value = (
            f"=COUNTIFS(Returns!$C$2:$C${last},$A{r})")
        dash.cell(row=r, column=3).value = (
            f"=SUMIFS(Returns!$O$2:$O${last},"
            f"Returns!$C$2:$C${last},$A{r})")
        dash.cell(row=r, column=4).value = (
            f"=IF(B{r}=0,0,100*C{r}/B{r})")
        dash.cell(row=r, column=4).number_format = RATE_FMT
    t = tr + 2 + len(MONTHS)  # total row = 41
    dash.cell(row=t, column=1, value="TOTAL (6 months)").font = Font(
        bold=True)
    dash.cell(row=t, column=2).value = f"=SUM(B{tr+2}:B{t-1})"
    dash.cell(row=t, column=3).value = f"=SUM(C{tr+2}:C{t-1})"
    dash.cell(row=t, column=4).value = (
        f"=IF(B{t}=0,0,100*C{t}/B{t})")
    dash.cell(row=t, column=4).number_format = RATE_FMT
    dash.column_dimensions["A"].width = 38
    dash.column_dimensions["B"].width = 26
    dash.column_dimensions["C"].width = 26

    chart = BarChart()
    chart.type = "col"
    chart.title = "Return rate by month (%)"
    chart.y_axis.title = "Return rate (%)"
    chart.x_axis.title = "Month"
    chart.data = Reference(dash, min_col=4, min_row=tr + 1, max_row=t - 1)
    chart.categories = Reference(dash, min_col=1, min_row=tr + 2,
                                 max_row=t - 1)
    chart.height, chart.width = 7.5, 15
    dash.add_chart(chart, f"A{t + 2}")

    # ── Method ─────────────────────────────────────────────────────────
    m = wb.create_sheet("Method")
    m["A1"] = "How this report is built"
    m["A1"].font = TITLE_FONT
    lines = [
        ("What this is", "A returns and refunds analysis for "
         "PakBazaar, a fictional Pakistani e-commerce store. 6 "
         "months, 2026-01-01 → 2026-06-30, 181 days, 1,200 orders, "
         "217 returns, 5 categories, 4 channels, 20 products, "
         "6 return reasons."),
        ("Data", f"SAMPLE: {n:,} orders. Generated by "
         "src/generate_data.py (seed 20261117). Demo data, never "
         "claimed as real client work. Each category was given a "
         "base return rate (Fashion highest, Beauty lowest), each "
         "product its own return propensity, each channel its own "
         "effect (Daraz highest), and March a sale spike — so the "
         "report has a real category story, a real channel story "
         "and a real monthly story to find."),
        ("Return rate", "Return rate = 100 x returned orders / "
         "total orders, in every group shown (category, channel, "
         "product, month). A group with few orders has a noisier "
         "rate — the workbook shows the order count next to every "
         "rate, and the Dashboard's MAXIFS / MINIFS rows only rank "
         "categories with 50+ orders."),
        ("Money lost", "Money lost = refunded revenue + return "
         "shipping cost. A refunded order gives the full order "
         "value back; a replaced order costs only the return "
         "shipping. Gross revenue - refunded revenue = net "
         "revenue; gross revenue - money lost = revenue kept after "
         "all return costs. Every sheet reconciles: the category, "
         "channel, product, reason and monthly totals all sum to "
         "the same headline numbers."),
        ("Refund vs replacement", "Of the 217 returns, 150 were "
         "refunded and 67 were replaced. Defective and wrong-item "
         "orders are usually replaced; changed-mind orders are "
         "usually refunded. The split is measured from the file, "
         "not assumed."),
        ("Return reasons", "Six reasons only: Wrong size, "
         "Defective/damaged, Wrong item sent, Late delivery, "
         "Quality not as expected, Changed mind. The reason mix "
         "differs by category in the demo data (Fashion returns "
         "are mostly wrong size, Electronics mostly defective), so "
         "the reason table and the top reason per product are "
         "measured, not typed."),
        ("Limits", "Return rates from 1,200 orders are an estimate "
         "of the store's behaviour, not a census — a product with "
         "34 orders has a wider margin of error than one with 96, "
         "and the workbook shows the order count next to every "
         "rate. Return reasons are the customer's stated reason; "
         "a 'quality not as expected' return can hide a sizing "
         "problem underneath. The fix-first actions rank where the "
         "money is lost; they do not promise how much comes back."),
        ("Live formulas", "Every KPI, category, channel, reason, "
         "product and month number is a live COUNTIFS / SUMIFS / "
         "INDEX-MATCH formula over the Returns sheet — paste a real "
         "orders export in the same columns and the report rebuilds "
         "itself. The Returns sheet's derived columns "
         "(is_returned, money_lost_pkr) are formulas too."),
        ("MAXIFS / MINIFS", "Worst and best category return rate "
         "among categories with 50+ orders on the Dashboard use "
         "MAXIFS and MINIFS, written with the _xlfn. prefix "
         "(=_xlfn.MAXIFS(...)) so LibreOffice and Excel calculate "
         "them — the plain form shows #NAME? in LibreOffice."),
        ("Check it", "Every number here is recomputed from the raw "
         "CSV by src/analyze.py and verified by src/checks.py — run "
         "it: python3 src/checks.py"),
        ("Same logic, two ways", "The headline KPIs, the category, "
         "channel, reason, product and monthly totals exist in two "
         "forms that must agree: the Python in src/analyze.py and "
         "this workbook's formulas. src/checks.py verifies both "
         "against the raw file, including a LibreOffice recalc."),
    ]
    r = 3
    for title, text in lines:
        m.cell(row=r, column=1, value=title).font = Font(bold=True)
        m.cell(row=r, column=2, value=text)
        m.cell(row=r, column=2).alignment = Alignment(wrap_text=True,
                                                      vertical="top")
        r += 1
    m.column_dimensions["A"].width = 22
    m.column_dimensions["B"].width = 110

    wb._sheets = [wb[n] for n in
                  ["Dashboard", "Returns", "By Category", "By Channel",
                   "By Reason", "Products", "Monthly Trend", "Method"]]
    wb.save(OUT)
    print(f"wrote {OUT.name}: sheets {wb.sheetnames}")


if __name__ == "__main__":
    main()
