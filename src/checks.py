#!/usr/bin/env python3
"""Checks as code for P21. Recomputes every headline number from the
raw CSV and verifies the JSON, the processed CSVs, the workbook
formulas and the LibreOffice recalc against it — three independent
paths that must agree. Money lost = refunded revenue + return
shipping cost, so gross - refunded = net revenue and
gross - money lost = revenue kept after all return costs; every
group total (category, channel, reason, product, month) must sum to
the same headline numbers.

    python3 src/checks.py        (after the full pipeline + recalc:
    soffice --headless --calc --convert-to xlsx --outdir outputs/.recalc
    outputs/returns_report_SAMPLE.xlsx)

Exit code 0 when every check passes, 1 otherwise.
"""
import csv
import json
import re
import subprocess
import sys
from pathlib import Path

from openpyxl import load_workbook

PROJECT = Path(__file__).resolve().parent.parent
RAW = PROJECT / "data" / "raw"
SUMMARY = PROJECT / "outputs" / "summary.json"
WORKBOOK = PROJECT / "outputs" / "returns_report_SAMPLE.xlsx"
RECALC = PROJECT / "outputs" / ".recalc" / "returns_report_SAMPLE.xlsx"
PDF = PROJECT / "outputs" / "returns_report.pdf"
MD = PROJECT / "outputs" / "returns_report.md"

CATEGORIES = ["Fashion", "Electronics", "Home & Kitchen", "Beauty",
              "Sports"]
CHANNELS = ["Website", "Daraz", "Facebook", "WhatsApp"]
MONTHS = [f"2026-{m:02d}" for m in range(1, 7)]
REASONS = ["Wrong size", "Defective/damaged", "Wrong item sent",
           "Late delivery", "Quality not as expected", "Changed mind"]
PRICES = {
    "FAS-001": 2499, "FAS-002": 3999, "FAS-003": 4499, "FAS-004": 1799,
    "ELE-001": 2999, "ELE-002": 2499, "ELE-003": 5999, "ELE-004": 999,
    "HOM-001": 6999, "HOM-002": 2199, "HOM-003": 4999, "HOM-004": 1499,
    "BEA-001": 1899, "BEA-002": 1299, "BEA-003": 899, "BEA-004": 2999,
    "SPO-001": 3499, "SPO-002": 999, "SPO-003": 1999, "SPO-004": 1499,
}
SHIP_COST = {"Website": 250, "Daraz": 200, "Facebook": 300,
             "WhatsApp": 250}
fails = []
ran = []


def check(name, cond, detail=""):
    ran.append(name)
    print(("PASS " if cond else "FAIL ") + name +
          (f" — {detail}" if detail else ""))
    if not cond:
        fails.append(name)


def pct(n, d):
    return round(100.0 * n / d, 2) if d else 0.0


def group_stats(rows):
    n = len(rows)
    ret = [r for r in rows if r["is_returned"]]
    gross = sum(r["order_value_pkr"] for r in rows)
    refunded = sum(r["refund_amount_pkr"] for r in rows)
    ship = sum(r["return_shipping_cost_pkr"] for r in rows)
    lost = refunded + ship
    return {
        "orders": n, "returns": len(ret),
        "return_rate": pct(len(ret), n),
        "gross_revenue": gross, "refunded_revenue": refunded,
        "return_shipping_cost": ship, "money_lost": lost,
        "net_revenue": gross - refunded,
        "revenue_kept": gross - lost,
        "avg_order_value": round(gross / n, 2) if n else 0.0,
    }


raw = list(csv.DictReader(open(RAW / "orders_SAMPLE.csv",
                               encoding="utf-8")))
rows = []
for r in raw:
    is_ret = 1 if r["return_status"] == "Returned" else 0
    refund = int(r["refund_amount_pkr"])
    ship = int(r["return_shipping_cost_pkr"])
    rows.append({
        "order_id": r["order_id"], "order_date": r["order_date"],
        "month": r["month"], "category": r["category"],
        "channel": r["channel"], "product_name": r["product_name"],
        "product_sku": r["product_sku"],
        "quantity": int(r["quantity"]),
        "order_value_pkr": int(r["order_value_pkr"]),
        "return_status": r["return_status"],
        "return_reason": r["return_reason"],
        "resolution": r["resolution"],
        "refund_amount_pkr": refund,
        "return_shipping_cost_pkr": ship,
        "is_returned": is_ret,
        "money_lost_pkr": refund + ship if is_ret else 0,
    })
total = len(rows)
returned = [r for r in rows if r["is_returned"]]
ov = group_stats(rows)
by_category = {c: group_stats([r for r in rows if r["category"] == c])
               for c in CATEGORIES}
by_channel = {c: group_stats([r for r in rows if r["channel"] == c])
              for c in CHANNELS}
by_month = {m: group_stats([r for r in rows if r["month"] == m])
            for m in MONTHS}
worst_cat = max(CATEGORIES, key=lambda c: by_category[c]["return_rate"])
best_cat = min(CATEGORIES, key=lambda c: by_category[c]["return_rate"])
worst_chan = max(CHANNELS, key=lambda c: by_channel[c]["return_rate"])
best_chan = min(CHANNELS, key=lambda c: by_channel[c]["return_rate"])
worst_month = max(MONTHS, key=lambda m: by_month[m]["return_rate"])
best_month = min(MONTHS, key=lambda m: by_month[m]["return_rate"])
worst_month_money = max(MONTHS,
                        key=lambda m: by_month[m]["money_lost"])
refunded_n = sum(1 for r in returned if r["resolution"] == "Refunded")
replaced_n = sum(1 for r in returned if r["resolution"] == "Replaced")

reasons = {}
for reason in REASONS:
    rs = [r for r in returned if r["return_reason"] == reason]
    reasons[reason] = {
        "returns": len(rs),
        "pct_of_returns": pct(len(rs), len(returned)),
        "refunded_revenue": sum(r["refund_amount_pkr"] for r in rs),
        "return_shipping_cost": sum(
            r["return_shipping_cost_pkr"] for r in rs),
        "money_lost": sum(r["money_lost_pkr"] for r in rs),
        "pct_of_money_lost": pct(sum(r["money_lost_pkr"] for r in rs),
                                 ov["money_lost"]),
    }
top_reason = max(REASONS, key=lambda x: reasons[x]["returns"])
top_reason_money = max(REASONS, key=lambda x: reasons[x]["money_lost"])

products = {}
for sku in sorted(PRICES):
    rs = [r for r in rows if r["product_sku"] == sku]
    rs_ret = [r for r in rs if r["is_returned"]]
    counts = {reason: sum(1 for r in rs_ret
                          if r["return_reason"] == reason)
              for reason in REASONS}
    products[sku] = {
        "name": rs[0]["product_name"], "category": rs[0]["category"],
        **group_stats(rs),
        "top_reason": max(counts, key=lambda k: counts[k])
        if rs_ret else "—",
    }
worst_prod_rate = max(products,
                      key=lambda k: products[k]["return_rate"])
worst_prod_money = max(products,
                       key=lambda k: products[k]["money_lost"])

# ── raw file ───────────────────────────────────────────────────────────
check("orders = 1200", total == 1200, str(total))
check("dates span 2026-01-01..2026-06-30",
      min(r["order_date"] for r in rows) == "2026-01-01" and
      max(r["order_date"] for r in rows) == "2026-06-30")
check("6 distinct months",
      sorted({r["month"] for r in rows}) == MONTHS)
check("5 categories only",
      {r["category"] for r in rows} == set(CATEGORIES))
check("4 channels only",
      {r["channel"] for r in rows} == set(CHANNELS))
check("20 products only",
      {r["product_sku"] for r in rows} == set(PRICES))
check("order_ids unique",
      len({r["order_id"] for r in rows}) == total)
check("order_ids are ORD-####",
      all(re.fullmatch(r"ORD-\d{4}", r["order_id"]) for r in rows))
check("month matches the order-date prefix",
      all(r["month"] == r["order_date"][:7] for r in rows))
check("quantity is 1 or 2 throughout",
      all(r["quantity"] in (1, 2) for r in rows))
check("order value = price x quantity throughout",
      all(r["order_value_pkr"] == PRICES[r["product_sku"]]
          * r["quantity"] for r in rows))
check("return_status is Kept or Returned only",
      {r["return_status"] for r in rows} == {"Kept", "Returned"})
check("kept orders have no reason, no resolution, no refund, "
      "no shipping",
      all(r["return_reason"] == "" and r["resolution"] == ""
          and r["refund_amount_pkr"] == 0
          and r["return_shipping_cost_pkr"] == 0
          and r["money_lost_pkr"] == 0
          for r in rows if not r["is_returned"]))
check("returned orders have a reason from the 6 reasons",
      all(r["return_reason"] in REASONS for r in returned))
check("returned orders have resolution Refunded or Replaced",
      all(r["resolution"] in ("Refunded", "Replaced")
          for r in returned))
check("refunded orders refund the full order value",
      all(r["refund_amount_pkr"] == r["order_value_pkr"]
          for r in returned if r["resolution"] == "Refunded"))
check("replaced orders refund 0",
      all(r["refund_amount_pkr"] == 0
          for r in returned if r["resolution"] == "Replaced"))
check("return shipping = the channel cost on every return",
      all(r["return_shipping_cost_pkr"] == SHIP_COST[r["channel"]]
          for r in returned))
check("money_lost = refund + shipping on every return",
      all(r["money_lost_pkr"] == r["refund_amount_pkr"]
          + r["return_shipping_cost_pkr"] for r in returned))
check("every category has >= 100 orders",
      min(v["orders"] for v in by_category.values()) >= 100,
      str(min(v["orders"] for v in by_category.values())))
check("every channel has >= 100 orders",
      min(v["orders"] for v in by_channel.values()) >= 100,
      str(min(v["orders"] for v in by_channel.values())))
check("every month has >= 150 orders",
      min(v["orders"] for v in by_month.values()) >= 150,
      str(min(v["orders"] for v in by_month.values())))
check("every product has >= 30 orders",
      min(v["orders"] for v in products.values()) >= 30,
      str(min(v["orders"] for v in products.values())))

# ── recompute: headline money ─────────────────────────────────────────
check("returns = 217", len(returned) == 217, str(len(returned)))
check("return rate = 18.08", ov["return_rate"] == 18.08,
      str(ov["return_rate"]))
check("return rate recomputes from counts",
      ov["return_rate"] == round(100.0 * len(returned) / total, 2))
check("gross revenue = 4203195", ov["gross_revenue"] == 4203195,
      str(ov["gross_revenue"]))
check("refunded revenue = 564528", ov["refunded_revenue"] == 564528,
      str(ov["refunded_revenue"]))
check("return shipping = 51500",
      ov["return_shipping_cost"] == 51500,
      str(ov["return_shipping_cost"]))
check("money lost = 616028", ov["money_lost"] == 616028,
      str(ov["money_lost"]))
check("money lost = refunded + shipping",
      ov["money_lost"] == ov["refunded_revenue"]
      + ov["return_shipping_cost"])
check("net revenue = 3638667", ov["net_revenue"] == 3638667,
      str(ov["net_revenue"]))
check("net revenue = gross - refunded",
      ov["net_revenue"] == ov["gross_revenue"]
      - ov["refunded_revenue"])
check("revenue kept = 3587167", ov["revenue_kept"] == 3587167,
      str(ov["revenue_kept"]))
check("revenue kept = gross - money lost",
      ov["revenue_kept"] == ov["gross_revenue"] - ov["money_lost"])
check("avg order value = 3502.66",
      ov["avg_order_value"] == 3502.66, str(ov["avg_order_value"]))
check("refunded share of gross = 13.43",
      pct(ov["refunded_revenue"], ov["gross_revenue"]) == 13.43)
check("money lost share of gross = 14.66",
      pct(ov["money_lost"], ov["gross_revenue"]) == 14.66)
check("refunded returns = 150", refunded_n == 150, str(refunded_n))
check("replaced returns = 67", replaced_n == 67, str(replaced_n))
check("refunded + replaced = returns",
      refunded_n + replaced_n == len(returned))
check("refund share of returns = 69.12",
      pct(refunded_n, len(returned)) == 69.12)
check("replaced share of returns = 30.88",
      pct(replaced_n, len(returned)) == 30.88)
check("category orders sum to 1200",
      sum(v["orders"] for v in by_category.values()) == total)
check("category returns sum to 217",
      sum(v["returns"] for v in by_category.values()) == 217)
check("category gross sums to 4203195",
      sum(v["gross_revenue"] for v in by_category.values())
      == ov["gross_revenue"])
check("category money lost sums to 616028",
      sum(v["money_lost"] for v in by_category.values())
      == ov["money_lost"])
check("channel orders sum to 1200",
      sum(v["orders"] for v in by_channel.values()) == total)
check("channel returns sum to 217",
      sum(v["returns"] for v in by_channel.values()) == 217)
check("channel money lost sums to 616028",
      sum(v["money_lost"] for v in by_channel.values())
      == ov["money_lost"])
check("monthly orders sum to 1200",
      sum(v["orders"] for v in by_month.values()) == total)
check("monthly returns sum to 217",
      sum(v["returns"] for v in by_month.values()) == 217)
check("monthly money lost sums to 616028",
      sum(v["money_lost"] for v in by_month.values())
      == ov["money_lost"])
check("product orders sum to 1200",
      sum(v["orders"] for v in products.values()) == total)
check("product returns sum to 217",
      sum(v["returns"] for v in products.values()) == 217)
check("product money lost sums to 616028",
      sum(v["money_lost"] for v in products.values())
      == ov["money_lost"])
check("reason returns sum to 217",
      sum(v["returns"] for v in reasons.values()) == 217)
check("reason money lost sums to 616028",
      sum(v["money_lost"] for v in reasons.values())
      == ov["money_lost"])
check("reason refunded sums to 564528",
      sum(v["refunded_revenue"] for v in reasons.values())
      == ov["refunded_revenue"])
check("reason shipping sums to 51500",
      sum(v["return_shipping_cost"] for v in reasons.values())
      == ov["return_shipping_cost"])

# ── worst / best ─────────────────────────────────────────────────────
check("worst category = Fashion", worst_cat == "Fashion", worst_cat)
check("worst category return rate = 30.17",
      by_category[worst_cat]["return_rate"] == 30.17)
check("best category = Beauty", best_cat == "Beauty", best_cat)
check("best category return rate = 6.92",
      by_category[best_cat]["return_rate"] == 6.92)
check("Fashion money lost = 317659",
      by_category["Fashion"]["money_lost"] == 317659)
check("worst channel = Daraz", worst_chan == "Daraz", worst_chan)
check("worst channel return rate = 21.3",
      by_channel[worst_chan]["return_rate"] == 21.3)
check("best channel = Facebook", best_chan == "Facebook", best_chan)
check("best channel return rate = 14.67",
      by_channel[best_chan]["return_rate"] == 14.67)
check("worst month = 2026-03", worst_month == "2026-03", worst_month)
check("worst month return rate = 23.73",
      by_month[worst_month]["return_rate"] == 23.73)
check("best month = 2026-02", best_month == "2026-02", best_month)
check("best month return rate = 14.46",
      by_month[best_month]["return_rate"] == 14.46)
check("worst month by money = 2026-04",
      worst_month_money == "2026-04", worst_month_money)
check("worst month by money = 156761",
      by_month[worst_month_money]["money_lost"] == 156761)
check("top reason = Wrong size", top_reason == "Wrong size",
      top_reason)
check("top reason returns = 47",
      reasons[top_reason]["returns"] == 47)
check("top reason money lost = 157156",
      reasons[top_reason]["money_lost"] == 157156)
check("top reason by money = Wrong size too",
      top_reason_money == "Wrong size", top_reason_money)
check("worst product by rate = Women's Lawn 3pc Suit",
      products[worst_prod_rate]["name"] == "Women's Lawn 3pc Suit",
      products[worst_prod_rate]["name"])
check("worst product by rate = 33.33",
      products[worst_prod_rate]["return_rate"] == 33.33)
check("worst product by money = Running Sneakers",
      products[worst_prod_money]["name"] == "Running Sneakers",
      products[worst_prod_money]["name"])
check("worst product by money = 105378",
      products[worst_prod_money]["money_lost"] == 105378)
check("Smart Watch money lost = 89036",
      products["ELE-003"]["money_lost"] == 89036)
check("Bedsheet Set is the lowest-rate product (1.82)",
      min(products, key=lambda k: products[k]["return_rate"])
      == "HOM-002")
for c in CATEGORIES:
    check(f"category {c}: return rate recomputes from counts",
          by_category[c]["return_rate"] == round(
              100.0 * by_category[c]["returns"]
              / by_category[c]["orders"], 2),
          str(by_category[c]["return_rate"]))
for c in CHANNELS:
    check(f"channel {c}: return rate recomputes from counts",
          by_channel[c]["return_rate"] == round(
              100.0 * by_channel[c]["returns"]
              / by_channel[c]["orders"], 2),
          str(by_channel[c]["return_rate"]))

# ── summary JSON ─────────────────────────────────────────────────────
s = json.loads(SUMMARY.read_text(encoding="utf-8"))
check("JSON total orders = 1200", s["total_orders"] == total)
check("JSON total returns matches recompute",
      s["total_returns"] == len(returned))
check("JSON return rate matches recompute",
      s["return_rate"] == ov["return_rate"])
check("JSON gross matches recompute",
      s["gross_revenue"] == ov["gross_revenue"])
check("JSON refunded matches recompute",
      s["refunded_revenue"] == ov["refunded_revenue"])
check("JSON shipping matches recompute",
      s["return_shipping_cost"] == ov["return_shipping_cost"])
check("JSON money lost matches recompute",
      s["money_lost"] == ov["money_lost"])
check("JSON net matches recompute",
      s["net_revenue"] == ov["net_revenue"])
check("JSON kept matches recompute",
      s["revenue_kept"] == ov["revenue_kept"])
check("JSON refunded count matches recompute",
      s["refunded_count"] == refunded_n)
check("JSON replaced count matches recompute",
      s["replaced_count"] == replaced_n)
check("JSON worst category matches recompute",
      s["worst_category"] == worst_cat)
check("JSON worst channel matches recompute",
      s["worst_channel"] == worst_chan)
check("JSON top reason matches recompute",
      s["top_reason"] == top_reason)
check("JSON worst product by rate matches recompute",
      s["worst_product_by_rate"]
      == products[worst_prod_rate]["name"])
check("JSON worst product by money matches recompute",
      s["worst_product_by_money"]
      == products[worst_prod_money]["name"])
check("JSON worst month matches recompute",
      s["worst_month"] == worst_month)
check("JSON worst month by money matches recompute",
      s["worst_month_by_money"] == worst_month_money)
check("JSON categories match recompute",
      all(x["return_rate"] == by_category[x["category"]]["return_rate"]
          and x["money_lost"]
          == by_category[x["category"]]["money_lost"]
          for x in s["by_category"]))
check("JSON channels match recompute",
      all(x["return_rate"] == by_channel[x["channel"]]["return_rate"]
          and x["money_lost"]
          == by_channel[x["channel"]]["money_lost"]
          for x in s["by_channel"]))
check("JSON reasons match recompute",
      all(x["returns"] == reasons[x["reason"]]["returns"]
          and x["money_lost"] == reasons[x["reason"]]["money_lost"]
          for x in s["by_reason"]))
check("JSON products match recompute",
      all(p["return_rate"]
          == products[p["product_sku"]]["return_rate"]
          and p["money_lost"]
          == products[p["product_sku"]]["money_lost"]
          for p in s["by_product"]))
check("JSON months match recompute",
      all(m["return_rate"] == by_month[m["month"]]["return_rate"]
          and m["money_lost"] == by_month[m["month"]]["money_lost"]
          for m in s["months"]))
check("JSON has 4 actions", len(s["actions"]) == 4)
check("JSON action 1 names the top reason",
      top_reason.lower() in s["actions"][0]["title"].lower()
      or top_reason.lower() in s["actions"][0]["detail"].lower())
check("JSON action 2 names the worst channel",
      worst_chan in s["actions"][1]["title"])
check("JSON action 3 names Electronics",
      "Electronics" in s["actions"][2]["title"])
check("JSON action 4 names the worst month",
      worst_month in s["actions"][3]["detail"])


def read_csv(name):
    return list(csv.DictReader(open(
        PROJECT / "data" / "processed" / name, encoding="utf-8")))


enr = read_csv("orders_enriched.csv")
check("enriched CSV has 1200 rows", len(enr) == 1200)
check("enriched is_returned flags match return_status throughout",
      all(int(r["is_returned"]) == (1 if r["return_status"]
                                    == "Returned" else 0)
          for r in enr))
check("enriched money_lost = refund + shipping throughout",
      all(int(r["money_lost_pkr"])
          == (int(r["refund_amount_pkr"])
              + int(r["return_shipping_cost_pkr"]))
          if r["return_status"] == "Returned"
          else int(r["money_lost_pkr"]) == 0 for r in enr))
check("enriched first row is ORD-0001",
      enr[0]["order_id"] == "ORD-0001")
ccsv = read_csv("category_summary.csv")
check("category CSV has 5 rows", len(ccsv) == 5)
check("category CSV return rates match recompute",
      all(float(r["return_rate"])
          == by_category[r["category"]]["return_rate"] for r in ccsv))
check("category CSV money lost matches recompute",
      all(int(r["money_lost"])
          == by_category[r["category"]]["money_lost"] for r in ccsv))
chcsv = read_csv("channel_summary.csv")
check("channel CSV has 4 rows", len(chcsv) == 4)
check("channel CSV return rates match recompute",
      all(float(r["return_rate"])
          == by_channel[r["channel"]]["return_rate"] for r in chcsv))
rcsv = read_csv("reason_summary.csv")
check("reason CSV has 6 rows", len(rcsv) == 6)
check("reason CSV first row is the top reason",
      rcsv[0]["reason"] == "Wrong size", rcsv[0]["reason"])
check("reason CSV money lost matches recompute",
      all(int(r["money_lost"])
          == reasons[r["reason"]]["money_lost"] for r in rcsv))
pcsv = read_csv("product_summary.csv")
check("product CSV has 20 rows", len(pcsv) == 20)
check("product CSV return rates match recompute",
      all(float(r["return_rate"])
          == products[r["product_sku"]]["return_rate"] for r in pcsv))
check("product CSV money lost matches recompute",
      all(int(r["money_lost"])
          == products[r["product_sku"]]["money_lost"] for r in pcsv))
check("product CSV worst by rate is Women's Lawn 3pc Suit",
      max(pcsv, key=lambda r: float(r["return_rate"]))["product_name"]
      == "Women's Lawn 3pc Suit")
mcsv = read_csv("monthly_summary.csv")
check("monthly CSV has 6 rows", len(mcsv) == 6)
check("monthly CSV return rates match recompute",
      all(float(r["return_rate"])
          == by_month[r["month"]]["return_rate"] for r in mcsv))
check("monthly CSV January rate change is blank",
      mcsv[0]["return_rate_change_points"] in ("", None))
check("monthly CSV March rate change = +9.27 (Mar minus Feb)",
      float(mcsv[2]["return_rate_change_points"]) == 9.27)

# ── workbook (formulas) ───────────────────────────────────────────────
wb = load_workbook(WORKBOOK)
check("sheets = Dashboard, Returns, By Category, By Channel, "
      "By Reason, Products, Monthly Trend, Method",
      wb.sheetnames == ["Dashboard", "Returns", "By Category",
                        "By Channel", "By Reason", "Products",
                        "Monthly Trend", "Method"],
      str(wb.sheetnames))
rsw = wb["Returns"]
check("Returns rows = 1201 incl header", rsw.max_row == 1201,
      str(rsw.max_row))
check("Returns header has 16 columns", rsw.max_column == 16,
      str(rsw.max_column))
check("Returns A2 = ORD-0001", rsw["A2"].value == "ORD-0001",
      str(rsw["A2"].value))
check("Returns O2 is the is_returned IF formula",
      rsw["O2"].value == '=IF(J2="Returned",1,0)',
      str(rsw["O2"].value))
check("Returns P2 is the money-lost IF formula",
      rsw["P2"].value == "=IF(O2=1,M2+N2,0)", str(rsw["P2"].value))
dash = wb["Dashboard"]
check("Dashboard B5 is a COUNTA formula over Returns",
      dash["B5"].value == f"=COUNTA(Returns!A2:A{total + 1})",
      str(dash["B5"].value))
check("Dashboard B6 sums is_returned flags",
      dash["B6"].value == f"=SUM(Returns!O2:O{total + 1})",
      str(dash["B6"].value))
check("Dashboard B7 is the return-rate formula",
      dash["B7"].value == "=IF(B5=0,0,100*B6/B5)",
      str(dash["B7"].value))
check("Dashboard B8 sums order values",
      dash["B8"].value == f"=SUM(Returns!I2:I{total + 1})",
      str(dash["B8"].value))
check("Dashboard B9 sums refunds",
      dash["B9"].value == f"=SUM(Returns!M2:M{total + 1})",
      str(dash["B9"].value))
check("Dashboard B10 sums return shipping",
      dash["B10"].value == f"=SUM(Returns!N2:N{total + 1})",
      str(dash["B10"].value))
check("Dashboard B11 is money lost = refunded + shipping",
      dash["B11"].value == "=B9+B10", str(dash["B11"].value))
check("Dashboard B12 is net = gross - refunded",
      dash["B12"].value == "=B8-B9", str(dash["B12"].value))
check("Dashboard B13 is kept = gross - money lost",
      dash["B13"].value == "=B8-B11", str(dash["B13"].value))
check("Dashboard B14 counts refunded returns",
      isinstance(dash["B14"].value, str)
      and dash["B14"].value.startswith("=COUNTIFS(Returns!$L$")
      and '"Refunded"' in dash["B14"].value, str(dash["B14"].value))
check("Dashboard B15 counts replaced returns",
      isinstance(dash["B15"].value, str)
      and dash["B15"].value.startswith("=COUNTIFS(Returns!$L$")
      and '"Replaced"' in dash["B15"].value, str(dash["B15"].value))
check("Dashboard B16 is the refund share formula",
      dash["B16"].value == "=IF(B6=0,0,B14/B6)",
      str(dash["B16"].value))
check("Dashboard B17 is the avg order value formula",
      dash["B17"].value == "=IF(B5=0,0,B8/B5)",
      str(dash["B17"].value))
check("Dashboard B21 is the worst-category INDEX/MATCH",
      isinstance(dash["B21"].value, str)
      and dash["B21"].value.startswith(
          "=INDEX('By Category'!A4:A8,MATCH(MAX("),
      str(dash["B21"].value))
check("Dashboard B22 is the best-category INDEX/MATCH",
      isinstance(dash["B22"].value, str)
      and dash["B22"].value.startswith(
          "=INDEX('By Category'!A4:A8,MATCH(MIN("),
      str(dash["B22"].value))
check("Dashboard B23 is the worst-channel INDEX/MATCH",
      isinstance(dash["B23"].value, str)
      and dash["B23"].value.startswith(
          "=INDEX('By Channel'!A4:A7,MATCH(MAX("),
      str(dash["B23"].value))
check("Dashboard B25 is the top-reason INDEX/MATCH",
      isinstance(dash["B25"].value, str)
      and dash["B25"].value.startswith(
          "=INDEX('By Reason'!A4:A9,MATCH(MAX("),
      str(dash["B25"].value))
check("Dashboard B26 is the worst-product INDEX/MATCH",
      isinstance(dash["B26"].value, str)
      and dash["B26"].value.startswith(
          "=INDEX(Products!A4:A23,MATCH(MAX("),
      str(dash["B26"].value))
check("Dashboard B27 is the worst-product-by-money INDEX/MATCH",
      isinstance(dash["B27"].value, str)
      and dash["B27"].value.startswith(
          "=INDEX(Products!A4:A23,MATCH(MAX(Products!J4:J23)"),
      str(dash["B27"].value))
check("Dashboard B30 uses _xlfn.MAXIFS for worst category (50+)",
      isinstance(dash["B30"].value, str)
      and dash["B30"].value.startswith("=_xlfn.MAXIFS("),
      str(dash["B30"].value))
check("Dashboard B31 uses _xlfn.MINIFS for best category (50+)",
      isinstance(dash["B31"].value, str)
      and dash["B31"].value.startswith("=_xlfn.MINIFS("),
      str(dash["B31"].value))
check("Dashboard monthly B35 is a COUNTIFS over Returns month",
      isinstance(dash["B35"].value, str)
      and dash["B35"].value.startswith("=COUNTIFS(Returns!$C$"),
      str(dash["B35"].value))
check("Dashboard monthly D35 is the return-rate formula",
      isinstance(dash["D35"].value, str)
      and dash["D35"].value.startswith("=IF(B35=0,0,100*C35/B35)"),
      str(dash["D35"].value))
check("Dashboard monthly total row uses SUM",
      dash["B41"].value == "=SUM(B35:B40)", str(dash["B41"].value))
check("Dashboard has a chart", len(dash._charts) >= 1)
catw = wb["By Category"]
check("By Category rows = 9 incl header + total", catw.max_row == 9,
      str(catw.max_row))
check("By Category B4 is a COUNTIFS over Returns category",
      isinstance(catw["B4"].value, str)
      and catw["B4"].value.startswith("=COUNTIFS(Returns!$D$"),
      str(catw["B4"].value))
check("By Category D4 is the return-rate formula",
      catw["D4"].value == "=IF(B4=0,0,100*C4/B4)",
      str(catw["D4"].value))
check("By Category H4 is money lost = refunded + shipping",
      catw["H4"].value == "=F4+G4", str(catw["H4"].value))
check("By Category I4 is net = gross - refunded",
      catw["I4"].value == "=E4-F4", str(catw["I4"].value))
check("By Category J4 is kept = gross - money lost",
      catw["J4"].value == "=E4-H4", str(catw["J4"].value))
check("By Category total row B9 sums orders",
      catw["B9"].value == "=SUM(B4:B8)", str(catw["B9"].value))
chw = wb["By Channel"]
check("By Channel B4 is a COUNTIFS over Returns channel",
      isinstance(chw["B4"].value, str)
      and chw["B4"].value.startswith("=COUNTIFS(Returns!$E$"),
      str(chw["B4"].value))
check("By Channel total row B8 sums orders",
      chw["B8"].value == "=SUM(B4:B7)", str(chw["B8"].value))
rsw2 = wb["By Reason"]
check("By Reason rows = 10 incl header + total", rsw2.max_row == 10,
      str(rsw2.max_row))
check("By Reason B4 counts returns by reason",
      isinstance(rsw2["B4"].value, str)
      and rsw2["B4"].value.startswith("=COUNTIFS(Returns!$K$"),
      str(rsw2["B4"].value))
check("By Reason F4 is money lost = refunded + shipping",
      rsw2["F4"].value == "=D4+E4", str(rsw2["F4"].value))
check("By Reason total row B10 sums returns",
      rsw2["B10"].value == "=SUM(B4:B9)", str(rsw2["B10"].value))
prw = wb["Products"]
check("Products rows = 24 incl header + total", prw.max_row == 24,
      str(prw.max_row))
check("Products D4 counts orders by SKU",
      isinstance(prw["D4"].value, str)
      and prw["D4"].value.startswith("=COUNTIFS(Returns!$G$"),
      str(prw["D4"].value))
check("Products F4 is the return-rate formula",
      prw["F4"].value == "=IF(D4=0,0,100*E4/D4)",
      str(prw["F4"].value))
check("Products J4 is money lost = refunded + shipping",
      prw["J4"].value == "=H4+I4", str(prw["J4"].value))
check("Products total row D24 sums orders",
      prw["D24"].value == "=SUM(D4:D23)", str(prw["D24"].value))
mow = wb["Monthly Trend"]
check("Monthly B4 is a COUNTIFS over Returns month",
      isinstance(mow["B4"].value, str)
      and mow["B4"].value.startswith("=COUNTIFS(Returns!$C$"),
      str(mow["B4"].value))
check("Monthly L4 is the em-dash for January",
      mow["L4"].value == "—", str(mow["L4"].value))
check("Monthly L5 is the rate-change formula",
      mow["L5"].value == "=D5-D4", str(mow["L5"].value))
check("Monthly total row B10 sums orders",
      mow["B10"].value == "=SUM(B4:B9)", str(mow["B10"].value))
check("Method sheet mentions _xlfn.",
      any("_xlfn." in (wb["Method"].cell(row=r, column=2).value or "")
          for r in range(3, 15)))
check("Method sheet states the return-rate formula",
      any("Return rate = 100" in
          (wb["Method"].cell(row=r, column=2).value or "")
          for r in range(3, 15)))
check("Method sheet states the money-lost formula",
      any("Money lost = refunded" in
          (wb["Method"].cell(row=r, column=2).value or "")
          for r in range(3, 15)))
bad = []
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith("=") \
                    and "{" in cell.value:
                bad.append(f"{ws.title}!{cell.coordinate}={cell.value}")
check("no formula contains an unexpanded {placeholder}",
      not bad, "; ".join(bad[:3]))

# ── LibreOffice recalculated values ────────────────────────────────────
if RECALC.exists():
    wb2 = load_workbook(RECALC, data_only=True)
    d2 = wb2["Dashboard"]
    check("recalc Dashboard total orders = 1200",
          d2["B5"].value == 1200, str(d2["B5"].value))
    check("recalc Dashboard returns = 217",
          d2["B6"].value == 217, str(d2["B6"].value))
    check("recalc Dashboard return rate matches recompute",
          abs(d2["B7"].value - ov["return_rate"]) < 0.01,
          str(d2["B7"].value))
    check("recalc Dashboard gross = 4203195",
          d2["B8"].value == 4203195, str(d2["B8"].value))
    check("recalc Dashboard refunded = 564528",
          d2["B9"].value == 564528, str(d2["B9"].value))
    check("recalc Dashboard shipping = 51500",
          d2["B10"].value == 51500, str(d2["B10"].value))
    check("recalc Dashboard money lost = 616028",
          d2["B11"].value == 616028, str(d2["B11"].value))
    check("recalc Dashboard net = 3638667",
          d2["B12"].value == 3638667, str(d2["B12"].value))
    check("recalc Dashboard kept = 3587167",
          d2["B13"].value == 3587167, str(d2["B13"].value))
    check("recalc Dashboard refunded count = 150",
          d2["B14"].value == 150, str(d2["B14"].value))
    check("recalc Dashboard replaced count = 67",
          d2["B15"].value == 67, str(d2["B15"].value))
    check("recalc Dashboard refund share matches recompute",
          abs(d2["B16"].value - 150 / 217) < 0.0005,
          str(d2["B16"].value))
    check("recalc Dashboard avg order value matches recompute",
          abs(d2["B17"].value - ov["avg_order_value"]) < 0.01,
          str(d2["B17"].value))
    check("recalc Dashboard money-lost share matches recompute",
          abs(d2["B19"].value - 0.1466) < 0.0005,
          str(d2["B19"].value))
    check("recalc Dashboard worst category = Fashion",
          d2["B21"].value == "Fashion", str(d2["B21"].value))
    check("recalc Dashboard worst category rate matches recompute",
          abs(d2["C21"].value
              - by_category[worst_cat]["return_rate"]) < 0.01,
          str(d2["C21"].value))
    check("recalc Dashboard best category = Beauty",
          d2["B22"].value == "Beauty", str(d2["B22"].value))
    check("recalc Dashboard worst channel = Daraz",
          d2["B23"].value == "Daraz", str(d2["B23"].value))
    check("recalc Dashboard best channel = Facebook",
          d2["B24"].value == "Facebook", str(d2["B24"].value))
    check("recalc Dashboard top reason = Wrong size",
          d2["B25"].value == "Wrong size", str(d2["B25"].value))
    check("recalc Dashboard top reason count = 47",
          d2["C25"].value == 47, str(d2["C25"].value))
    check("recalc Dashboard worst product = Women's Lawn 3pc Suit",
          d2["B26"].value == "Women's Lawn 3pc Suit",
          str(d2["B26"].value))
    check("recalc Dashboard worst product by money = Running Sneakers",
          d2["B27"].value == "Running Sneakers", str(d2["B27"].value))
    check("recalc Dashboard worst product money = 105378",
          d2["C27"].value == 105378, str(d2["C27"].value))
    check("recalc Dashboard worst month = 2026-03",
          d2["B28"].value == "2026-03", str(d2["B28"].value))
    check("recalc worst category (50+) = 30.17",
          abs(d2["B30"].value - 30.17) < 0.01, str(d2["B30"].value))
    check("recalc worst category (50+) name = Fashion",
          d2["C30"].value == "Fashion", str(d2["C30"].value))
    check("recalc best category (50+) = 6.92",
          abs(d2["B31"].value - 6.92) < 0.01, str(d2["B31"].value))
    check("recalc best category (50+) name = Beauty",
          d2["C31"].value == "Beauty", str(d2["C31"].value))
    check("recalc monthly Jan return rate matches recompute",
          abs(d2["D35"].value - by_month["2026-01"]["return_rate"])
          < 0.01, str(d2["D35"].value))
    check("recalc monthly Mar return rate matches recompute",
          abs(d2["D37"].value - by_month["2026-03"]["return_rate"])
          < 0.01, str(d2["D37"].value))
    check("recalc monthly TOTAL orders = 1200",
          d2["B41"].value == 1200, str(d2["B41"].value))
    check("recalc monthly TOTAL returns = 217",
          d2["C41"].value == 217, str(d2["C41"].value))
    check("recalc monthly TOTAL return rate matches recompute",
          abs(d2["D41"].value - ov["return_rate"]) < 0.01,
          str(d2["D41"].value))
    c2 = wb2["By Category"]
    for name in CATEGORIES:
        row = next(r for r in range(4, 9)
                   if c2.cell(row=r, column=1).value == name)
        v = by_category[name]
        check(f"recalc By Category {name}: return rate matches",
              abs(c2.cell(row=row, column=4).value
                  - v["return_rate"]) < 0.01,
              str(c2.cell(row=row, column=4).value))
        check(f"recalc By Category {name}: money lost matches",
              c2.cell(row=row, column=8).value == v["money_lost"],
              str(c2.cell(row=row, column=8).value))
    check("recalc By Category TOTAL orders = 1200",
          c2["B9"].value == 1200, str(c2["B9"].value))
    check("recalc By Category TOTAL money lost = 616028",
          c2["H9"].value == 616028, str(c2["H9"].value))
    h2 = wb2["By Channel"]
    for name in CHANNELS:
        row = next(r for r in range(4, 8)
                   if h2.cell(row=r, column=1).value == name)
        check(f"recalc By Channel {name}: return rate matches",
              abs(h2.cell(row=row, column=4).value
                  - by_channel[name]["return_rate"]) < 0.01,
              str(h2.cell(row=row, column=4).value))
    check("recalc By Channel TOTAL orders = 1200",
          h2["B8"].value == 1200, str(h2["B8"].value))
    r2 = wb2["By Reason"]
    for reason in REASONS:
        row = next(r for r in range(4, 10)
                   if r2.cell(row=r, column=1).value == reason)
        v = reasons[reason]
        check(f"recalc By Reason {reason}: returns match",
              r2.cell(row=row, column=2).value == v["returns"],
              str(r2.cell(row=row, column=2).value))
        check(f"recalc By Reason {reason}: money lost matches",
              r2.cell(row=row, column=6).value == v["money_lost"],
              str(r2.cell(row=row, column=6).value))
    check("recalc By Reason TOTAL returns = 217",
          r2["B10"].value == 217, str(r2["B10"].value))
    check("recalc By Reason TOTAL money lost = 616028",
          r2["F10"].value == 616028, str(r2["F10"].value))
    p2 = wb2["Products"]
    for sku, v in products.items():
        row = next(r for r in range(4, 24)
                   if p2.cell(row=r, column=2).value == sku)
        check(f"recalc Products {v['name']}: return rate matches",
              abs(p2.cell(row=row, column=6).value
                  - v["return_rate"]) < 0.01,
              str(p2.cell(row=row, column=6).value))
        check(f"recalc Products {v['name']}: money lost matches",
              p2.cell(row=row, column=10).value == v["money_lost"],
              str(p2.cell(row=row, column=10).value))
    check("recalc Products TOTAL orders = 1200",
          p2["D24"].value == 1200, str(p2["D24"].value))
    check("recalc Products TOTAL money lost = 616028",
          p2["J24"].value == 616028, str(p2["J24"].value))
    m2 = wb2["Monthly Trend"]
    check("recalc Monthly TOTAL orders = 1200",
          m2["B10"].value == 1200, str(m2["B10"].value))
    check("recalc Monthly Mar rate change = +9.27 (Mar minus Feb)",
          abs(m2["L6"].value - 9.27) < 0.01, str(m2["L6"].value))
    o2 = wb2["Returns"]
    check("recalc Returns is_returned flags sum to 217",
          sum(o2.cell(row=r, column=15).value or 0
              for r in range(2, 1202)) == 217)
    check("recalc Returns money_lost sums to 616028",
          sum(o2.cell(row=r, column=16).value or 0
              for r in range(2, 1202)) == 616028)
    errors = []
    for ws in wb2.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("#"):
                    errors.append(f"{ws.title}!{cell.coordinate}={cell.value}")
    check("recalc workbook has no error values (#NAME?, #VALUE!, ...)",
          not errors, "; ".join(errors[:5]))
else:
    print(f"  NOTE: no LibreOffice recalc copy found at {RECALC} — "
          "recalc checks skipped")

# ── PDF + markdown ─────────────────────────────────────────────────────
check("returns_report.pdf exists", PDF.exists())
if PDF.exists():
    info = subprocess.run(["pdfinfo", str(PDF)], capture_output=True,
                          text=True)
    pages = re.search(r"Pages:\s+(\d+)", info.stdout)
    check("PDF has exactly 2 pages",
          pages is not None and int(pages.group(1)) == 2,
          info.stdout.splitlines()[0] if info.stdout else "pdfinfo failed")
    text = subprocess.run(["pdftotext", str(PDF), "-"], capture_output=True,
                          text=True).stdout
    check("PDF text contains 1,200", "1,200" in text)
    check("PDF text contains return rate 18.08", "18.08" in text)
    check("PDF text contains money lost 616,028", "616,028" in text)
    check("PDF text contains Fashion", "Fashion" in text)
    check("PDF text contains Daraz", "Daraz" in text)
    check("PDF text contains Wrong size", "Wrong size" in text)
    check("PDF text contains Women's Lawn", "Women's Lawn" in text)
    check("PDF text contains Running Sneakers", "Running Sneakers" in text)
check("returns_report.md exists", MD.exists())
if MD.exists():
    md = MD.read_text(encoding="utf-8")
    check("MD contains return rate 18.08", "18.08" in md)
    check("MD contains money lost 616,028", "616,028" in md)
    check("MD contains top reason Wrong size", "Wrong size" in md)
    check("MD contains 4 recommendations",
          md.count("**Fix sizing") + md.count("**Tighten the Daraz")
          + md.count("**Quality-check") + md.count("**Plan for the sale")
          == 4)

print(f"\n{len(ran) - len(fails)}/{len(ran)} checks passed")
if fails:
    print("FAILURES:", fails)
    sys.exit(1)
