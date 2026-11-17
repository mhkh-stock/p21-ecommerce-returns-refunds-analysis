#!/usr/bin/env python3
"""P21 — generate the SAMPLE e-commerce orders file for PakBazaar
(SAMPLE).

Fictional Pakistani online store, 6 months of orders:
Thu 2026-01-01 -> Tue 2026-06-30 (181 days). Seeded, so the file is
reproducible. One file, one row per order:

  data/raw/orders_SAMPLE.csv — grain: one order
      (order_id, order_date, month, category, channel, product_name,
       product_sku, quantity, order_value_pkr, return_status,
       return_reason, resolution, refund_amount_pkr,
       return_shipping_cost_pkr)

How the demo data is built (and why the report can still find real
things in it): each category has a base return rate (Fashion returns
most, Beauty least) and each product its own return propensity, so
one product is clearly the worst; each channel has its own effect
(Daraz returns most); March carries a sale spike, so returns spike
there too. A return's reason is drawn from a per-category reason
mix (Fashion returns are mostly wrong size, Electronics mostly
defective), and the resolution follows the reason (defective or
wrong-item orders are usually replaced, changed-mind orders are
usually refunded). Refund = the full order value when refunded,
0 when replaced; return shipping is a fixed per-channel courier
cost on every return, 0 otherwise. Money lost = refunded revenue
+ return shipping cost, so gross - refunded = net revenue and
gross - money lost = revenue kept after all return costs.

    python3 src/generate_data.py
"""
import csv
import random
from datetime import date, timedelta
from pathlib import Path

PROJECT = Path(__file__).resolve().parent.parent
RAW = PROJECT / "data" / "raw"

SEED = 20261117
START = date(2026, 1, 1)
END = date(2026, 6, 30)
N_ORDERS = 1200

# category: (base return rate, volume weight)
CATEGORIES = {
    "Fashion": (0.22, 0.30),
    "Electronics": (0.12, 0.25),
    "Home & Kitchen": (0.10, 0.20),
    "Beauty": (0.08, 0.13),
    "Sports": (0.14, 0.12),
}
# product: (category, base price PKR, return propensity multiplier)
PRODUCTS = [
    ("Men's Kurta — Navy", "FAS-001", "Fashion", 2499, 1.60),
    ("Women's Lawn 3pc Suit", "FAS-002", "Fashion", 3999, 1.20),
    ("Running Sneakers", "FAS-003", "Fashion", 4499, 1.30),
    ("Kids Frock — Pink", "FAS-004", "Fashion", 1799, 0.90),
    ("Wireless Earbuds", "ELE-001", "Electronics", 2999, 1.40),
    ("Power Bank 10000mAh", "ELE-002", "Electronics", 2499, 0.80),
    ("Smart Watch", "ELE-003", "Electronics", 5999, 1.50),
    ("LED Bulb Pack (4 pcs)", "ELE-004", "Electronics", 999, 0.60),
    ("Dinner Set (32 pcs)", "HOM-001", "Home & Kitchen", 6999, 0.70),
    ("Bedsheet Set — King", "HOM-002", "Home & Kitchen", 2199, 0.90),
    ("Non-stick Cookware Set", "HOM-003", "Home & Kitchen", 4999, 1.10),
    ("Wall Clock — Wooden", "HOM-004", "Home & Kitchen", 1499, 0.80),
    ("Face Serum 30ml", "BEA-001", "Beauty", 1899, 0.70),
    ("Lipstick Set (6 pcs)", "BEA-002", "Beauty", 1299, 0.80),
    ("Shampoo 500ml", "BEA-003", "Beauty", 899, 0.60),
    ("Perfume 50ml", "BEA-004", "Beauty", 2999, 1.00),
    ("Cricket Bat — Kashmir Willow", "SPO-001", "Sports", 3499, 1.10),
    ("Gym Gloves", "SPO-002", "Sports", 999, 0.90),
    ("Football — Size 5", "SPO-003", "Sports", 1999, 1.00),
    ("Yoga Mat 6mm", "SPO-004", "Sports", 1499, 0.80),
]
# channel: (return-rate multiplier, return shipping cost PKR,
#           volume weight)
CHANNELS = {
    "Website": (1.00, 250, 0.45),
    "Daraz": (1.25, 200, 0.30),
    "Facebook": (1.15, 300, 0.15),
    "WhatsApp": (0.90, 250, 0.10),
}
# month: (return-rate multiplier, volume weight) — March carries the
# sale spike, so returns spike there too
MONTHS = [f"2026-{m:02d}" for m in range(1, 7)]
MONTH_EFFECT = {"2026-01": (0.90, 0.16), "2026-02": (0.95, 0.15),
                "2026-03": (1.35, 0.19), "2026-04": (1.05, 0.17),
                "2026-05": (1.00, 0.16), "2026-06": (0.85, 0.17)}
REASONS = ["Wrong size", "Defective/damaged", "Wrong item sent",
           "Late delivery", "Quality not as expected", "Changed mind"]
# per-category reason mix (weights in REASONS order)
REASON_MIX = {
    "Fashion": [0.38, 0.08, 0.08, 0.12, 0.18, 0.16],
    "Electronics": [0.10, 0.34, 0.14, 0.14, 0.16, 0.12],
    "Home & Kitchen": [0.12, 0.24, 0.12, 0.14, 0.22, 0.16],
    "Beauty": [0.12, 0.12, 0.14, 0.16, 0.26, 0.20],
    "Sports": [0.28, 0.16, 0.10, 0.12, 0.20, 0.14],
}
# refund probability by reason (the rest are replaced)
REFUND_PROB = {"Wrong size": 0.80, "Defective/damaged": 0.45,
               "Wrong item sent": 0.45, "Late delivery": 0.70,
               "Quality not as expected": 0.75, "Changed mind": 0.90}


def main():
    rng = random.Random(SEED)
    month_days = {}
    for m in range(1, 7):
        first = date(2026, m, 1)
        nxt = date(2026, m + 1, 1) if m < 6 else date(2026, 7, 1)
        month_days[f"2026-{m:02d}"] = (first, (nxt - first).days)

    cat_names = list(CATEGORIES)
    cat_weights = [CATEGORIES[c][1] for c in cat_names]
    chan_names = list(CHANNELS)
    chan_weights = [CHANNELS[c][2] for c in chan_names]
    month_weights = [MONTH_EFFECT[m][1] for m in MONTHS]

    rows = []
    for _ in range(N_ORDERS):
        mi = rng.choices(range(6), weights=month_weights)[0]
        mkey = MONTHS[mi]
        first, ndays = month_days[mkey]
        day = first + timedelta(days=rng.randrange(ndays))
        category = rng.choices(cat_names, weights=cat_weights)[0]
        channel = rng.choices(chan_names, weights=chan_weights)[0]
        prods = [p for p in PRODUCTS if p[2] == category]
        name, sku, _cat, price, mult = rng.choice(prods)
        qty = 2 if rng.random() < 0.15 else 1
        value = price * qty
        p_return = min(0.45, max(0.02,
                                 CATEGORIES[category][0] * mult
                                 * CHANNELS[channel][0]
                                 * MONTH_EFFECT[mkey][0]))
        returned = rng.random() < p_return
        if returned:
            reason = rng.choices(REASONS,
                                 weights=REASON_MIX[category])[0]
            resolution = ("Refunded" if rng.random()
                          < REFUND_PROB[reason] else "Replaced")
            refund = value if resolution == "Refunded" else 0
            shipping = CHANNELS[channel][1]
            status = "Returned"
        else:
            reason, resolution, refund, shipping = "", "", 0, 0
            status = "Kept"
        rows.append({
            "order_id": "", "order_date": day.isoformat(),
            "month": mkey, "category": category, "channel": channel,
            "product_name": name, "product_sku": sku,
            "quantity": qty, "order_value_pkr": value,
            "return_status": status, "return_reason": reason,
            "resolution": resolution, "refund_amount_pkr": refund,
            "return_shipping_cost_pkr": shipping,
        })

    rows.sort(key=lambda r: (r["order_date"], r["product_sku"]))
    for i, r in enumerate(rows):
        r["order_id"] = f"ORD-{i + 1:04d}"
    with open(RAW / "orders_SAMPLE.csv", "w", newline="",
              encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=[
            "order_id", "order_date", "month", "category", "channel",
            "product_name", "product_sku", "quantity",
            "order_value_pkr", "return_status", "return_reason",
            "resolution", "refund_amount_pkr",
            "return_shipping_cost_pkr"])
        w.writeheader()
        w.writerows(rows)

    n_ret = sum(1 for r in rows if r["return_status"] == "Returned")
    gross = sum(r["order_value_pkr"] for r in rows)
    refunded = sum(r["refund_amount_pkr"] for r in rows)
    ship = sum(r["return_shipping_cost_pkr"] for r in rows)
    print(f"wrote {len(rows)} orders "
          f"({rows[0]['order_date']}..{rows[-1]['order_date']}), "
          f"returns {n_ret} ({100.0 * n_ret / len(rows):.2f}%), "
          f"gross PKR {gross:,}, refunded PKR {refunded:,}, "
          f"return shipping PKR {ship:,}, money lost PKR "
          f"{refunded + ship:,}")


if __name__ == "__main__":
    main()
