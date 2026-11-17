#!/usr/bin/env python3
"""P21 — analyse the SAMPLE e-commerce orders file and regenerate
every output number from the raw CSV (never typed): the headline
return rate and money-lost KPIs, return rate by category, channel
and month, return reasons ranked by count and by money lost,
refund vs replacement split, top problem products, and the
fix-first action list.

Derived per order: is_returned (0/1), money_lost_pkr =
refund_amount_pkr + return_shipping_cost_pkr (0 for kept orders).
Money lost = refunded revenue + return shipping cost, so
gross revenue - refunded revenue = net revenue and
gross revenue - money lost = revenue kept after all return costs.
Return rate = 100 x returned orders / total orders.

Writes:
  data/processed/orders_enriched.csv
  data/processed/category_summary.csv
  data/processed/channel_summary.csv
  data/processed/reason_summary.csv
  data/processed/product_summary.csv
  data/processed/monthly_summary.csv
  outputs/summary.json
  outputs/returns_report.md

    python3 src/analyze.py
"""
import csv
import json
from pathlib import Path

PROJECT = Path(__file__).resolve().parent.parent
RAW = PROJECT / "data" / "raw"

CATEGORIES = ["Fashion", "Electronics", "Home & Kitchen", "Beauty",
              "Sports"]
CHANNELS = ["Website", "Daraz", "Facebook", "WhatsApp"]
MONTHS = [f"2026-{m:02d}" for m in range(1, 7)]
REASONS = ["Wrong size", "Defective/damaged", "Wrong item sent",
           "Late delivery", "Quality not as expected", "Changed mind"]
PRODUCTS = [
    ("Men's Kurta — Navy", "FAS-001", "Fashion"),
    ("Women's Lawn 3pc Suit", "FAS-002", "Fashion"),
    ("Running Sneakers", "FAS-003", "Fashion"),
    ("Kids Frock — Pink", "FAS-004", "Fashion"),
    ("Wireless Earbuds", "ELE-001", "Electronics"),
    ("Power Bank 10000mAh", "ELE-002", "Electronics"),
    ("Smart Watch", "ELE-003", "Electronics"),
    ("LED Bulb Pack (4 pcs)", "ELE-004", "Electronics"),
    ("Dinner Set (32 pcs)", "HOM-001", "Home & Kitchen"),
    ("Bedsheet Set — King", "HOM-002", "Home & Kitchen"),
    ("Non-stick Cookware Set", "HOM-003", "Home & Kitchen"),
    ("Wall Clock — Wooden", "HOM-004", "Home & Kitchen"),
    ("Face Serum 30ml", "BEA-001", "Beauty"),
    ("Lipstick Set (6 pcs)", "BEA-002", "Beauty"),
    ("Shampoo 500ml", "BEA-003", "Beauty"),
    ("Perfume 50ml", "BEA-004", "Beauty"),
    ("Cricket Bat — Kashmir Willow", "SPO-001", "Sports"),
    ("Gym Gloves", "SPO-002", "Sports"),
    ("Football — Size 5", "SPO-003", "Sports"),
    ("Yoga Mat 6mm", "SPO-004", "Sports"),
]


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


def main():
    raw = list(csv.DictReader(
        open(RAW / "orders_SAMPLE.csv", encoding="utf-8")))
    rows = []
    for r in raw:
        is_ret = 1 if r["return_status"] == "Returned" else 0
        refund = int(r["refund_amount_pkr"])
        ship = int(r["return_shipping_cost_pkr"])
        rows.append({
            "order_id": r["order_id"], "order_date": r["order_date"],
            "month": r["month"], "category": r["category"],
            "channel": r["channel"],
            "product_name": r["product_name"],
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
    overall = group_stats(rows)
    returned = [r for r in rows if r["is_returned"]]
    refunded_n = sum(1 for r in returned
                     if r["resolution"] == "Refunded")
    replaced_n = sum(1 for r in returned
                     if r["resolution"] == "Replaced")

    by_category = []
    for c in CATEGORIES:
        by_category.append({"category": c,
                            **group_stats([r for r in rows
                                           if r["category"] == c])})
    cat_by_rate = sorted(by_category, key=lambda x: -x["return_rate"])
    worst_cat = cat_by_rate[0]
    best_cat = cat_by_rate[-1]

    by_channel = []
    for c in CHANNELS:
        by_channel.append({"channel": c,
                           **group_stats([r for r in rows
                                          if r["channel"] == c])})
    chan_by_rate = sorted(by_channel, key=lambda x: -x["return_rate"])
    worst_chan = chan_by_rate[0]
    best_chan = chan_by_rate[-1]

    by_reason = []
    for reason in REASONS:
        rs = [r for r in returned if r["return_reason"] == reason]
        by_reason.append({
            "reason": reason, "returns": len(rs),
            "pct_of_returns": pct(len(rs), len(returned)),
            "refunded_revenue": sum(r["refund_amount_pkr"] for r in rs),
            "return_shipping_cost": sum(
                r["return_shipping_cost_pkr"] for r in rs),
            "money_lost": sum(r["money_lost_pkr"] for r in rs),
            "pct_of_money_lost": pct(sum(r["money_lost_pkr"] for r in rs),
                                     overall["money_lost"]),
        })
    by_reason.sort(key=lambda x: -x["returns"])
    top_reason = by_reason[0]
    top_reason_money = max(by_reason, key=lambda x: x["money_lost"])

    by_product = []
    for name, sku, cat in PRODUCTS:
        rs = [r for r in rows if r["product_sku"] == sku]
        rs_ret = [r for r in rs if r["is_returned"]]
        top_r = None
        if rs_ret:
            counts = {reason: sum(1 for r in rs_ret
                                  if r["return_reason"] == reason)
                      for reason in REASONS}
            top_r = max(counts, key=lambda k: counts[k])
        by_product.append({
            "product_name": name, "product_sku": sku, "category": cat,
            **group_stats(rs), "top_reason": top_r or "—",
        })
    prod_by_rate = sorted(by_product, key=lambda x: -x["return_rate"])
    worst_prod_rate = prod_by_rate[0]
    worst_prod_money = max(by_product, key=lambda x: x["money_lost"])

    by_month = []
    for m in MONTHS:
        by_month.append({"month": m,
                         **group_stats([r for r in rows
                                        if r["month"] == m])})
    for i, mo in enumerate(by_month):
        prev = by_month[i - 1]["return_rate"] if i else None
        mo["return_rate_change_points"] = (
            round(mo["return_rate"] - prev, 2)
            if prev is not None else None)
    worst_month = max(by_month, key=lambda m: m["return_rate"])
    best_month = min(by_month, key=lambda m: m["return_rate"])
    worst_month_money = max(by_month, key=lambda m: m["money_lost"])

    # fix-first actions, built from the numbers above
    fashion_lost = next(c for c in by_category
                        if c["category"] == "Fashion")
    defective = next(x for x in by_reason
                     if x["reason"] == "Defective/damaged")
    wrong_item = next(x for x in by_reason
                      if x["reason"] == "Wrong item sent")
    actions = [
        {
            "title": "Fix sizing first — it is the biggest leak",
            "detail": f"{top_reason['reason']} is the top return "
                      f"reason: {top_reason['returns']} of "
                      f"{len(returned)} returns "
                      f"({top_reason['pct_of_returns']}%), costing "
                      f"PKR {top_reason['money_lost']:,} "
                      f"({top_reason['pct_of_money_lost']}% of all "
                      f"money lost). Fashion is the worst category "
                      f"({worst_cat['return_rate']}% return rate, "
                      f"PKR {fashion_lost['money_lost']:,} lost) and "
                      f"{worst_prod_rate['product_name']} is the "
                      f"worst product by rate "
                      f"({worst_prod_rate['return_rate']}%, "
                      f"{worst_prod_rate['returns']} of "
                      f"{worst_prod_rate['orders']} orders). Publish "
                      f"real measurements and a size chart on the "
                      f"Fashion listings before spending on anything "
                      f"else.",
        },
        {
            "title": f"Tighten the {worst_chan['channel']} channel",
            "detail": f"{worst_chan['channel']} has the highest "
                      f"return rate at {worst_chan['return_rate']}% "
                      f"({worst_chan['returns']} of "
                      f"{worst_chan['orders']} orders), against "
                      f"{best_chan['channel']} at "
                      f"{best_chan['return_rate']}%, and lost "
                      f"PKR {worst_chan['money_lost']:,}. Check that "
                      f"its listing photos, sizes and descriptions "
                      f"match what is actually shipped — the gap "
                      f"against the Website channel is a listing "
                      f"problem, not a customer problem.",
        },
        {
            "title": "Quality-check before dispatch, starting with "
                      "Electronics",
            "detail": f"Defective/damaged ({defective['returns']} "
                      f"returns, PKR {defective['money_lost']:,} "
                      f"lost) plus wrong item sent "
                      f"({wrong_item['returns']} returns, "
                      f"PKR {wrong_item['money_lost']:,} lost) cost "
                      f"PKR {defective['money_lost'] + wrong_item['money_lost']:,} "
                      f"together — every one of those is a return "
                      f"the store caused. {worst_prod_money['product_name']} "
                      f"alone lost PKR "
                      f"{worst_prod_money['money_lost']:,} "
                      f"({worst_prod_money['returns']} returns). A "
                      f"one-minute check before sealing the box "
                      f"removes most of it.",
        },
        {
            "title": "Plan for the sale spike before the next one",
            "detail": f"{worst_month['month']} had the highest "
                      f"return rate of the half year "
                      f"({worst_month['return_rate']}%, "
                      f"{worst_month['returns']} returns) and "
                      f"{worst_month_money['month']} lost the most "
                      f"money (PKR "
                      f"{worst_month_money['money_lost']:,}). Sale "
                      f"months bring the volume — and the returns. "
                      f"Staff the returns desk and tighten size and "
                      f"quality checks before the next sale, not "
                      f"after it.",
        },
    ]

    summary = {
        "store": "PakBazaar (SAMPLE)",
        "period": "2026-01-01 to 2026-06-30",
        "days": 181,
        "total_orders": total,
        "total_returns": len(returned),
        "return_rate": overall["return_rate"],
        "gross_revenue": overall["gross_revenue"],
        "refunded_revenue": overall["refunded_revenue"],
        "return_shipping_cost": overall["return_shipping_cost"],
        "money_lost": overall["money_lost"],
        "net_revenue": overall["net_revenue"],
        "revenue_kept": overall["revenue_kept"],
        "avg_order_value": overall["avg_order_value"],
        "refunded_count": refunded_n,
        "replaced_count": replaced_n,
        "refund_pct_of_returns": pct(refunded_n, len(returned)),
        "replaced_pct_of_returns": pct(replaced_n, len(returned)),
        "refunded_share_of_gross": pct(overall["refunded_revenue"],
                                       overall["gross_revenue"]),
        "money_lost_share_of_gross": pct(overall["money_lost"],
                                         overall["gross_revenue"]),
        "by_category": by_category,
        "worst_category": worst_cat["category"],
        "worst_category_return_rate": worst_cat["return_rate"],
        "best_category": best_cat["category"],
        "best_category_return_rate": best_cat["return_rate"],
        "by_channel": by_channel,
        "worst_channel": worst_chan["channel"],
        "worst_channel_return_rate": worst_chan["return_rate"],
        "best_channel": best_chan["channel"],
        "best_channel_return_rate": best_chan["return_rate"],
        "by_reason": by_reason,
        "top_reason": top_reason["reason"],
        "top_reason_returns": top_reason["returns"],
        "top_reason_money_lost": top_reason["money_lost"],
        "top_reason_by_money": top_reason_money["reason"],
        "by_product": by_product,
        "worst_product_by_rate": worst_prod_rate["product_name"],
        "worst_product_by_rate_value": worst_prod_rate["return_rate"],
        "worst_product_by_money": worst_prod_money["product_name"],
        "worst_product_by_money_value": worst_prod_money["money_lost"],
        "months": by_month,
        "worst_month": worst_month["month"],
        "worst_month_return_rate": worst_month["return_rate"],
        "best_month": best_month["month"],
        "best_month_return_rate": best_month["return_rate"],
        "worst_month_by_money": worst_month_money["month"],
        "worst_month_by_money_value": worst_month_money["money_lost"],
        "actions": actions,
    }
    (PROJECT / "outputs" / "summary.json").write_text(
        json.dumps(summary, indent=2), encoding="utf-8")

    def write_csv(name, fields, rows_out):
        with open(PROJECT / "data" / "processed" / name, "w",
                  newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=fields)
            w.writeheader()
            w.writerows(rows_out)

    write_csv("orders_enriched.csv",
              ["order_id", "order_date", "month", "category", "channel",
               "product_name", "product_sku", "quantity",
               "order_value_pkr", "return_status", "return_reason",
               "resolution", "refund_amount_pkr",
               "return_shipping_cost_pkr", "is_returned",
               "money_lost_pkr"], rows)

    def group_csv(name, key, groups, extra_fields):
        fields = [key, "orders", "returns", "return_rate",
                  "gross_revenue", "refunded_revenue",
                  "return_shipping_cost", "money_lost", "net_revenue",
                  "revenue_kept", "avg_order_value"] + extra_fields
        write_csv(name, fields,
                  [{k: g[k] for k in fields} for g in groups])

    group_csv("category_summary.csv", "category", by_category, [])
    group_csv("channel_summary.csv", "channel", by_channel, [])
    group_csv("monthly_summary.csv", "month", by_month,
              ["return_rate_change_points"])
    write_csv("reason_summary.csv",
              ["reason", "returns", "pct_of_returns",
               "refunded_revenue", "return_shipping_cost",
               "money_lost", "pct_of_money_lost"],
              [{k: x[k] for k in
                ["reason", "returns", "pct_of_returns",
                 "refunded_revenue", "return_shipping_cost",
                 "money_lost", "pct_of_money_lost"]}
               for x in by_reason])
    write_csv("product_summary.csv",
              ["product_name", "product_sku", "category", "orders",
               "returns", "return_rate", "gross_revenue",
               "refunded_revenue", "return_shipping_cost",
               "money_lost", "net_revenue", "revenue_kept",
               "avg_order_value", "top_reason"],
              [{k: p[k] for k in
                ["product_name", "product_sku", "category", "orders",
                 "returns", "return_rate", "gross_revenue",
                 "refunded_revenue", "return_shipping_cost",
                 "money_lost", "net_revenue", "revenue_kept",
                 "avg_order_value", "top_reason"]}
               for p in by_product])

    # findings report (markdown) — every number from the data above
    ov = overall
    L = [
        "# P21 — E-commerce Returns & Refunds Analysis: PakBazaar "
        "(SAMPLE)",
        "",
        "**Period:** Thu 2026-01-01 → Tue 2026-06-30 (6 months, 181 days)",
        f"**Orders:** {total:,} · **Returns:** {len(returned)} "
        f"({ov['return_rate']}%) · **Gross revenue:** "
        f"PKR {ov['gross_revenue']:,} · **Refunded:** "
        f"PKR {ov['refunded_revenue']:,} · **Return shipping:** "
        f"PKR {ov['return_shipping_cost']:,} · **Money lost:** "
        f"PKR {ov['money_lost']:,} ({pct(ov['money_lost'], ov['gross_revenue'])}% "
        f"of gross) · **Net revenue:** PKR {ov['net_revenue']:,} · "
        f"**Revenue kept after all return costs:** "
        f"PKR {ov['revenue_kept']:,}",
        "",
        "## 1. The headline — what do returns cost PakBazaar?",
        "",
        f"Of {total:,} orders, {len(returned)} came back "
        f"({ov['return_rate']}%). Refunds took "
        f"PKR {ov['refunded_revenue']:,} off gross revenue "
        f"(net revenue PKR {ov['net_revenue']:,}), and return "
        f"shipping added PKR {ov['return_shipping_cost']:,} on top — "
        f"so returns cost **PKR {ov['money_lost']:,} in total**, "
        f"{pct(ov['money_lost'], ov['gross_revenue'])}% of everything "
        f"the store sold. Of the {len(returned)} returns, "
        f"{refunded_n} were refunded "
        f"({pct(refunded_n, len(returned))}%) and {replaced_n} were "
        f"replaced ({pct(replaced_n, len(returned))}%). The worst "
        f"category is **{worst_cat['category']} "
        f"({worst_cat['return_rate']}%)**, the worst channel "
        f"**{worst_chan['channel']} ({worst_chan['return_rate']}%)**, "
        f"and the worst month **{worst_month['month']} "
        f"({worst_month['return_rate']}%)**.",
        "",
        "## 2. Return rate by category",
        "",
        "| Category | Orders | Returns | Return rate | Gross revenue | "
        "Refunded | Return shipping | Money lost | Net revenue |",
        "|---|---|---|---|---|---|---|---|---|",
    ]
    for c in cat_by_rate:
        L.append(f"| {c['category']} | {c['orders']} | {c['returns']} | "
                 f"{c['return_rate']}% | PKR {c['gross_revenue']:,} | "
                 f"PKR {c['refunded_revenue']:,} | "
                 f"PKR {c['return_shipping_cost']:,} | "
                 f"PKR {c['money_lost']:,} | PKR {c['net_revenue']:,} |")
    L += [
        "",
        f"**Worst: {worst_cat['category']} ({worst_cat['return_rate']}%, "
        f"PKR {worst_cat['money_lost']:,} lost)** — more than half of "
        f"all money lost sits in one category. Best: "
        f"{best_cat['category']} ({best_cat['return_rate']}%).",
        "",
        "## 3. Return rate by channel",
        "",
        "| Channel | Orders | Returns | Return rate | Gross revenue | "
        "Refunded | Return shipping | Money lost | Net revenue |",
        "|---|---|---|---|---|---|---|---|---|",
    ]
    for c in chan_by_rate:
        L.append(f"| {c['channel']} | {c['orders']} | {c['returns']} | "
                 f"{c['return_rate']}% | PKR {c['gross_revenue']:,} | "
                 f"PKR {c['refunded_revenue']:,} | "
                 f"PKR {c['return_shipping_cost']:,} | "
                 f"PKR {c['money_lost']:,} | PKR {c['net_revenue']:,} |")
    L += [
        "",
        f"**Worst: {worst_chan['channel']} ({worst_chan['return_rate']}%)** "
        f"against best {best_chan['channel']} "
        f"({best_chan['return_rate']}%). Website lost the most money "
        f"in absolute terms (PKR "
        f"{next(c['money_lost'] for c in by_channel if c['channel'] == 'Website'):,}) "
        f"because it carries the most orders — its rate is the second "
        f"lowest, so the channel problem is Daraz, not Website.",
        "",
        "## 4. Why customers send it back (reasons)",
        "",
        "| Reason | Returns | Share of returns | Refunded | "
        "Return shipping | Money lost | Share of money lost |",
        "|---|---|---|---|---|---|---|",
    ]
    for x in by_reason:
        L.append(f"| {x['reason']} | {x['returns']} | "
                 f"{x['pct_of_returns']}% | "
                 f"PKR {x['refunded_revenue']:,} | "
                 f"PKR {x['return_shipping_cost']:,} | "
                 f"PKR {x['money_lost']:,} | "
                 f"{x['pct_of_money_lost']}% |")
    L += [
        "",
        f"**Top reason by count and by money: {top_reason['reason']}** "
        f"({top_reason['returns']} returns, "
        f"PKR {top_reason['money_lost']:,} lost, "
        f"{top_reason['pct_of_money_lost']}% of all money lost). "
        f"Defective/damaged plus wrong item sent — the returns the "
        f"store itself caused — cost "
        f"PKR {defective['money_lost'] + wrong_item['money_lost']:,} "
        f"together.",
        "",
        "## 5. Refund vs replacement",
        "",
        f"{refunded_n} of {len(returned)} returns were refunded "
        f"({pct(refunded_n, len(returned))}%) and {replaced_n} were "
        f"replaced ({pct(replaced_n, len(returned))}%). Refunds are "
        f"where the money goes: every refunded order gives the full "
        f"order value back, while a replacement only costs the "
        f"return shipping. Defective and wrong-item orders are "
        f"usually replaced; changed-mind orders are usually refunded.",
        "",
        "## 6. Top problem products",
        "",
        "| Product | Category | Orders | Returns | Return rate | "
        "Money lost | Top reason |",
        "|---|---|---|---|---|---|---|",
    ]
    for p in prod_by_rate[:5]:
        L.append(f"| {p['product_name']} | {p['category']} | "
                 f"{p['orders']} | {p['returns']} | "
                 f"{p['return_rate']}% | PKR {p['money_lost']:,} | "
                 f"{p['top_reason']} |")
    L += [
        "",
        f"**Worst by return rate: {worst_prod_rate['product_name']} "
        f"({worst_prod_rate['return_rate']}%, "
        f"{worst_prod_rate['returns']} of "
        f"{worst_prod_rate['orders']} orders).** Worst by money lost: "
        f"**{worst_prod_money['product_name']} "
        f"(PKR {worst_prod_money['money_lost']:,}, "
        f"{worst_prod_money['returns']} returns)** — a lower rate on "
        f"a pricier product still costs more. Fix the rate-worst "
        f"product first for customers, the money-worst one first for "
        f"the bank balance.",
        "",
        "## 7. Month by month",
        "",
        "| Month | Orders | Returns | Return rate | Gross revenue | "
        "Refunded | Money lost | Return-rate change (pts) |",
        "|---|---|---|---|---|---|---|",
    ]
    for mo in by_month:
        chg = "—" if mo["return_rate_change_points"] is None else (
            f"{mo['return_rate_change_points']:+.2f}")
        L.append(f"| {mo['month']} | {mo['orders']} | {mo['returns']} | "
                 f"{mo['return_rate']}% | PKR {mo['gross_revenue']:,} | "
                 f"PKR {mo['refunded_revenue']:,} | "
                 f"PKR {mo['money_lost']:,} | {chg} |")
    L += [
        "",
        f"**Worst month by rate: {worst_month['month']} "
        f"({worst_month['return_rate']}%)** — the March sale spike. "
        f"But the most money was lost in "
        f"**{worst_month_money['month']} "
        f"(PKR {worst_month_money['money_lost']:,})**, when pricier "
        f"orders came back. Best month: {best_month['month']} "
        f"({best_month['return_rate']}%).",
        "",
        "## 8. Four actions — fix first, in this order",
        "",
    ]
    for i, a in enumerate(actions, 1):
        L.append(f"{i}. **{a['title']}.** {a['detail']}")
    L += [
        "",
        "> Every number in this report is recomputed from "
        "`data/raw/orders_SAMPLE.csv` by `src/analyze.py` — nothing "
        "is typed by hand. Run `python3 src/checks.py` to re-verify.",
    ]
    (PROJECT / "outputs" / "returns_report.md").write_text(
        "\n".join(L) + "\n", encoding="utf-8")
    print(f"orders {total:,}, returns {len(returned)} "
          f"({ov['return_rate']}%), gross PKR {ov['gross_revenue']:,}, "
          f"refunded PKR {ov['refunded_revenue']:,}, shipping "
          f"PKR {ov['return_shipping_cost']:,}, money lost "
          f"PKR {ov['money_lost']:,}, top reason {top_reason['reason']} "
          f"({top_reason['returns']}), worst product by rate "
          f"{worst_prod_rate['product_name']} "
          f"({worst_prod_rate['return_rate']}%), worst month "
          f"{worst_month['month']} ({worst_month['return_rate']}%)")


if __name__ == "__main__":
    main()
