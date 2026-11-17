#!/usr/bin/env python3
"""Build the P21 deliverables from the SAMPLE orders data.

Recomputes every number from outputs/summary.json (itself recomputed
from the raw CSV by src/analyze.py — nothing is typed in by hand)
and produces:
  * outputs/monthly_return_rate.png — monthly return rate (bars)
  * outputs/reason_money_lost.png — money lost by return reason
    (matplotlib)
  * outputs/returns_report.pdf — 2-page client report (headline,
    category and channel return rates, reasons ranked by money lost,
    refund vs replacement split, top problem products, monthly
    trend, 4 recommendations)

Run:  python3 src/build_report.py
"""
import json
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    Image,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

ROOT = Path(__file__).resolve().parent.parent
SUMMARY = ROOT / "outputs" / "summary.json"
CHART_MO = ROOT / "outputs" / "monthly_return_rate.png"
CHART_RS = ROOT / "outputs" / "reason_money_lost.png"
PDF = ROOT / "outputs" / "returns_report.pdf"

s = json.loads(SUMMARY.read_text(encoding="utf-8"))
TOT = s["total_orders"]
RET = s["total_returns"]
CAT = sorted(s["by_category"], key=lambda x: -x["return_rate"])
CH = sorted(s["by_channel"], key=lambda x: -x["return_rate"])
RS = sorted(s["by_reason"], key=lambda x: -x["money_lost"])
MO = s["months"]
PR = sorted(s["by_product"], key=lambda x: -x["return_rate"])[:5]
ACTS = s["actions"]

# monthly return rate: bars
months = [m["month"] for m in MO]
rates = [m["return_rate"] for m in MO]
fig, ax = plt.subplots(figsize=(8, 3.2))
ax.bar(months, rates, width=0.55, color="#1F4E79")
ax.set_ylabel("Return rate (%)")
ax.set_title("Monthly return rate — PakBazaar (SAMPLE, Jan–Jun 2026)")
for i, v in enumerate(rates):
    ax.text(i, v + 0.5, f"{v}%", ha="center", fontsize=9)
fig.tight_layout()
plt.savefig(CHART_MO, dpi=150)
plt.close()

# money lost by return reason, ranked (best first = top)
labels = [x["reason"] for x in RS]
lost = [x["money_lost"] for x in RS]
cols = ["#C62828" if i == 0 else ("#F9A825" if i <= 2 else "#2E7D32")
        for i in range(len(RS))]
plt.figure(figsize=(8, 3.6))
plt.barh(labels[::-1], lost[::-1], color=cols[::-1])
plt.xlabel("Money lost (PKR) — red = biggest leak, amber = next, "
           "green = smallest")
plt.title("Money lost by return reason — PakBazaar (SAMPLE)")
for i, v in enumerate(lost[::-1]):
    plt.text(v + 3000, i, f"PKR {v:,}", va="center", fontsize=9)
plt.tight_layout()
plt.savefig(CHART_RS, dpi=150)
plt.close()

styles = getSampleStyleSheet()
hdr = colors.HexColor("#1F4E79")


def table(rows, widths, size=8.5):
    t = Table(rows, colWidths=widths)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), hdr),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("FONTSIZE", (0, 0), (-1, -1), size),
    ]))
    return t


story = [
    Paragraph("PakBazaar — E-commerce Returns & Refunds Analysis",
              styles["Title"]),
    Paragraph("Orders analysed by category, channel, product, reason "
              "and month · 1 Jan – 30 Jun 2026 (6 months, 181 days) · "
              "SAMPLE data — fictional Pakistani online store, "
              "generated for this delivery (seed 20261117)",
              styles["Normal"]),
    Spacer(1, 0.1 * cm),
    Paragraph(
        f"Headline: <b>{TOT:,} orders</b>, <b>{RET} returned "
        f"({s['return_rate']}%)</b>. Gross revenue "
        f"<b>PKR {s['gross_revenue']:,}</b>; refunded "
        f"<b>PKR {s['refunded_revenue']:,}</b>; return shipping "
        f"<b>PKR {s['return_shipping_cost']:,}</b>; so returns cost "
        f"<b>PKR {s['money_lost']:,}</b> "
        f"({s['money_lost_share_of_gross']}% of gross). "
        f"{s['refunded_count']} returns were refunded "
        f"({s['refund_pct_of_returns']}%) and {s['replaced_count']} "
        f"replaced ({s['replaced_pct_of_returns']}%). Top reason: "
        f"<b>{s['top_reason']} ({s['top_reason_returns']} returns, "
        f"PKR {s['top_reason_money_lost']:,} lost)</b>. Worst "
        f"category: <b>{s['worst_category']} "
        f"({s['worst_category_return_rate']}%)</b>; worst channel: "
        f"<b>{s['worst_channel']} "
        f"({s['worst_channel_return_rate']}%)</b>; worst month: "
        f"<b>{s['worst_month']} "
        f"({s['worst_month_return_rate']}%)</b>.",
        styles["Normal"]),
    Spacer(1, 0.1 * cm),
    Paragraph("<b>Return rate by category — ranked</b>",
              styles["Heading2"]),
]
cat_rows = [["Category", "Orders", "Returns", "Return rate",
             "Gross (PKR)", "Refunded (PKR)", "Money lost (PKR)"]]
for c in CAT:
    cat_rows.append([c["category"], f"{c['orders']}",
                     f"{c['returns']}", f"{c['return_rate']}%",
                     f"{c['gross_revenue']:,}",
                     f"{c['refunded_revenue']:,}",
                     f"{c['money_lost']:,}"])
story.append(table(cat_rows,
                   [3.2 * cm, 1.6 * cm, 1.7 * cm, 2.2 * cm, 2.6 * cm,
                    2.6 * cm, 2.8 * cm]))
story.append(Spacer(1, 0.1 * cm))
story.append(Paragraph("<b>Return rate by channel</b>",
                       styles["Heading2"]))
ch_rows = [["Channel", "Orders", "Returns", "Return rate",
            "Gross (PKR)", "Refunded (PKR)", "Money lost (PKR)"]]
for c in CH:
    ch_rows.append([c["channel"], f"{c['orders']}",
                    f"{c['returns']}", f"{c['return_rate']}%",
                    f"{c['gross_revenue']:,}",
                    f"{c['refunded_revenue']:,}",
                    f"{c['money_lost']:,}"])
story.append(table(ch_rows,
                   [3.2 * cm, 1.6 * cm, 1.7 * cm, 2.2 * cm, 2.6 * cm,
                    2.6 * cm, 2.8 * cm]))
story.append(Spacer(1, 0.12 * cm))
story.append(Paragraph(
    f"Refund vs replacement: <b>{s['refunded_count']} refunded "
    f"({s['refund_pct_of_returns']}%)</b>, "
    f"<b>{s['replaced_count']} replaced "
    f"({s['replaced_pct_of_returns']}%)</b>. Refunds are where the "
    f"money goes — every refunded order gives the full order value "
    f"back, while a replacement only costs the return shipping.",
    styles["Normal"]))
story.append(Spacer(1, 0.1 * cm))
story.append(Paragraph(
    "<b>Why customers send it back (reasons, ranked by money lost)</b>",
    styles["Heading2"]))
rs_rows = [["Reason", "Returns", "Share of returns",
            "Refunded (PKR)", "Money lost (PKR)",
            "Share of money lost"]]
for x in RS:
    rs_rows.append([x["reason"], f"{x['returns']}",
                    f"{x['pct_of_returns']}%",
                    f"{x['refunded_revenue']:,}",
                    f"{x['money_lost']:,}",
                    f"{x['pct_of_money_lost']}%"])
story.append(table(rs_rows,
                   [4.0 * cm, 1.8 * cm, 2.6 * cm, 2.6 * cm, 2.8 * cm,
                    2.9 * cm]))
story.append(Spacer(1, 0.12 * cm))
story.append(Image(str(CHART_RS), width=15.5 * cm, height=4.1 * cm))
story.append(Spacer(1, 0.12 * cm))
story.append(Paragraph(
    f"Top reason by count and by money: <b>{s['top_reason']}</b> "
    f"({s['top_reason_returns']} returns, "
    f"PKR {s['top_reason_money_lost']:,} lost). Money lost = refunded "
    f"revenue + return shipping cost.",
    styles["Normal"]))
story.append(Spacer(1, 0.1 * cm))
story.append(Paragraph("<b>Top problem products (by return rate)</b>",
                       styles["Heading2"]))
pr_rows = [["Product", "Category", "Orders", "Returns",
            "Return rate", "Money lost (PKR)", "Top reason"]]
for p in PR:
    pr_rows.append([p["product_name"], p["category"],
                    f"{p['orders']}", f"{p['returns']}",
                    f"{p['return_rate']}%",
                    f"{p['money_lost']:,}", p["top_reason"]])
story.append(table(pr_rows,
                   [4.4 * cm, 2.4 * cm, 1.5 * cm, 1.6 * cm, 2.0 * cm,
                    2.6 * cm, 2.2 * cm]))
story.append(Spacer(1, 0.12 * cm))
story.append(Paragraph(
    f"Worst by return rate: <b>{s['worst_product_by_rate']} "
    f"({s['worst_product_by_rate_value']}%)</b>; worst by money "
    f"lost: <b>{s['worst_product_by_money']} "
    f"(PKR {s['worst_product_by_money_value']:,})</b>.",
    styles["Normal"]))
story.append(Spacer(1, 0.1 * cm))
story.append(Paragraph("<b>Month by month</b>", styles["Heading2"]))
mo_rows = [["Month", "Orders", "Returns", "Return rate",
            "Gross (PKR)", "Money lost (PKR)", "Rate change (pts)"]]
for m in MO:
    chg = "—" if m["return_rate_change_points"] is None else (
        f"{m['return_rate_change_points']:+.2f}")
    mo_rows.append([m["month"], f"{m['orders']}",
                    f"{m['returns']}", f"{m['return_rate']}%",
                    f"{m['gross_revenue']:,}",
                    f"{m['money_lost']:,}", chg])
story.append(table(mo_rows,
                   [2.2 * cm, 1.8 * cm, 1.8 * cm, 2.2 * cm, 2.6 * cm,
                    2.8 * cm, 2.3 * cm]))
story.append(Spacer(1, 0.12 * cm))
story.append(Image(str(CHART_MO), width=15.5 * cm, height=3.5 * cm))
story.append(Spacer(1, 0.12 * cm))
story.append(Paragraph(
    f"Worst month by rate: <b>{s['worst_month']} "
    f"({s['worst_month_return_rate']}%)</b> — the March sale spike. "
    f"Most money lost in <b>{s['worst_month_by_money']} "
    f"(PKR {s['worst_month_by_money_value']:,})</b>. Best month: "
    f"{s['best_month']} ({s['best_month_return_rate']}%).",
    styles["Normal"]))
story.append(Spacer(1, 0.1 * cm))
story.append(Paragraph("<b>4 recommendations — fix first, in this "
                       "order</b>", styles["Heading2"]))
for i, a in enumerate(ACTS, 1):
    story.append(Paragraph(
        f"{i}. <b>{a['title']}.</b> {a['detail']}", styles["Normal"]))
    story.append(Spacer(1, 0.1 * cm))
story.append(Spacer(1, 0.05 * cm))
story.append(Paragraph(
    "Method: every number above is recomputed from "
    "data/raw/orders_SAMPLE.csv by src/analyze.py and verified by "
    "src/checks.py. The workbook carries the same numbers as live "
    "formulas. Limits: return rates from 1,200 orders are an "
    "estimate, not a census — a product with few orders has a wider "
    "margin of error. Return reasons are the customer's stated "
    "reason; a 'quality not as expected' return can hide a sizing "
    "problem underneath.",
    styles["Normal"]))
SimpleDocTemplate(str(PDF), pagesize=A4,
                  title="PakBazaar — E-commerce Returns & Refunds "
                        "Analysis (SAMPLE)",
                  author="Muhammad Hashir",
                  topMargin=0.9 * cm, bottomMargin=0.9 * cm).build(story)
print(f"charts: {CHART_MO.name}, {CHART_RS.name}")
print(f"report: {PDF}")
