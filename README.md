# P21 — E-commerce Returns & Refunds Analysis

**Client (SAMPLE):** PakBazaar — fictional Pakistani online store,
6 months (Thu 2026-01-01 → Tue 2026-06-30, 181 days) ·
**Status: SAMPLE delivery on DEMO data**
**Repo:** https://github.com/mhkh-stock/p21-ecommerce-returns-refunds-analysis
**Tools:** Python · Excel (live formulas) · reportlab
**Deliverable:** `outputs/returns_report.pdf` (2 pages) +
`outputs/returns_report_SAMPLE.xlsx`

> ⚠️ **What this is and is not.** The orders file is a fictional
> store's demo data, generated to show the method. No real store,
> product or customer is involved. What is real is the analysis:
> the same return-rate, reason and money-lost breakdowns and
> workbook run on a real client's orders export the same way, and
> every number in the report is recomputed from the data, never
> typed.

## Business question it answers

> "Why are customers sending it back — and what is it costing us?"

A store that only watches total sales cannot tell whether its
returns are a sizing problem, a quality problem, a listing
problem or a channel problem — or how much money each one is
actually eating. This delivery answers it from the orders: the
return rate overall, by category, channel, product and month,
why customers send items back (ranked by count and by money),
what refunds plus return shipping cost in total, how refunds
split against replacements, and the fix-first action list — with
the biggest leak named, not guessed.

## The method (plainly stated)

**Return rate = 100 × returned orders / total orders**, in every
group shown (category, channel, product, month). A group with few
orders has a noisier rate — the workbook shows the order count
next to every rate, and the Dashboard's MAXIFS / MINIFS rows only
rank categories with 50+ orders.

**Money lost = refunded revenue + return shipping cost.** A
refunded order gives the full order value back; a replaced order
costs only the return shipping (a fixed per-channel courier cost:
Website PKR 250, Daraz PKR 200, Facebook PKR 300, WhatsApp
PKR 250). So gross revenue − refunded revenue = net revenue, and
gross revenue − money lost = revenue kept after all return costs.
Every group total (category, channel, reason, product, month)
sums to the same headline numbers.

**Reasons:** six reasons only — Wrong size, Defective/damaged,
Wrong item sent, Late delivery, Quality not as expected, Changed
mind. The reason mix differs by category in the demo data
(Fashion returns are mostly wrong size, Electronics mostly
defective), so the reason table and the top reason per product are
measured, not typed.

**Refund vs replacement:** defective and wrong-item orders are
usually replaced; changed-mind orders are usually refunded. The
split is measured from the file, not assumed.

**Limits, stated upfront:** return rates from 1,200 orders are an
estimate of the store's behaviour, not a census — a product with
34 orders has a wider margin of error than one with 96, and the
workbook shows the order count next to every rate. Return reasons
are the customer's stated reason; a "quality not as expected"
return can hide a sizing problem underneath. The fix-first actions
rank where the money is lost; they do not promise how much comes
back.

## What is inside

| File | What it does |
|---|---|
| `outputs/returns_report.pdf` | The deliverable: 2-page client report — headline, category and channel return rates, reasons ranked by money lost, refund vs replacement split, top problem products, monthly trend, 4 recommendations |
| `outputs/returns_report_SAMPLE.xlsx` | The working report: Dashboard (KPIs + monthly return-rate trend, live formulas) + Returns + By Category + By Channel + By Reason + Products + Monthly Trend + Method sheets |
| `outputs/returns_report.md` | Findings report: headline numbers, all tables, the reasoning |
| `outputs/summary.json` | Machine-readable summary (overall KPIs + categories + channels + reasons + products + months + actions) |
| `outputs/monthly_return_rate.png` | Monthly return-rate trend chart (render of the same data) |
| `outputs/reason_money_lost.png` | Money lost by return reason, ranked (render of the same data) |
| `data/raw/orders_SAMPLE.csv` | The orders file: 1,200 orders, category + channel + product + quantity + value + return status + reason + resolution + refund + shipping |
| `data/processed/orders_enriched.csv` | All 1,200 orders with the derived columns attached (is_returned, money_lost_pkr) |
| `data/processed/category_summary.csv` | 5-row category summary produced by `src/analyze.py` |
| `data/processed/channel_summary.csv` | 4-row channel summary produced by `src/analyze.py` |
| `data/processed/reason_summary.csv` | 6-row reason summary produced by `src/analyze.py` |
| `data/processed/product_summary.csv` | 20-row product summary produced by `src/analyze.py` |
| `data/processed/monthly_summary.csv` | 6-month summary produced by `src/analyze.py` |
| `data/README.md` | Data dictionary + known quirks |
| `src/generate_data.py` | Makes the demo file (seeded; NOT run on a real client file) |
| `src/analyze.py` | Recomputes every headline number from the raw CSV |
| `src/build_workbook.py` | Builds the workbook: 8 sheets, live formulas |
| `src/build_report.py` | Builds the 2-page PDF + charts from `outputs/summary.json` |
| `src/checks.py` | Data-quality + reconciliation checks as code — 326 checks, all passing |
| `requirements.txt` | Pinned: `openpyxl==3.1.2`, `reportlab==5.0.1`, `matplotlib==3.6.3` (Python 3.12) |

## How it fits together

```
Client orders export (same columns) ──▶ src/analyze.py ──▶ summary.json
        (1,200 rows in the demo)                + returns_report.md
                                                 ▼
                                       src/build_workbook.py ──▶ returns_report_SAMPLE.xlsx
                                                 │              (8 sheets, live formulas)
                                                 ▼
                                       src/build_report.py ──▶ returns_report.pdf (2 pages)
                                                 ▲ src/checks.py verifies all of it (326 checks)
```

## Headline numbers (SAMPLE — PakBazaar, 6 months)

- **1,200 orders** · **217 returns (18.08%)** · **gross revenue
  PKR 4,203,195** · **refunded PKR 564,528 (13.43% of gross)** ·
  **return shipping PKR 51,500** · **money lost PKR 616,028
  (14.66% of gross)** · **net revenue PKR 3,638,667** ·
  **revenue kept after all return costs PKR 3,587,167**
- **Refunded: 150 returns (69.12%)**; **replaced: 67 (30.88%)**
- **Top reason: Wrong size — 47 returns (21.66%), PKR 157,156
  lost (25.51% of all money lost)**
- **Worst category: Fashion (30.17%, PKR 317,659 lost — 51.57%
  of all money lost)**; best: **Beauty (6.92%)**
- **Worst channel: Daraz (21.30%)**; best: **Facebook (14.67%)**
- **Worst product by return rate: Women's Lawn 3pc Suit
  (33.33%, 32 of 96 orders)**; worst by money lost: **Running
  Sneakers (PKR 105,378, 27 returns)**
- **Worst month by rate: March 2026 (23.73%)** — the sale spike;
  most money lost in **April (PKR 156,761)**

## Return rate by category

| Category | Orders | Returns | Return rate | Gross revenue | Refunded | Return shipping | Money lost | Net revenue |
|---|---|---|---|---|---|---|---|---|
| Fashion | 358 | 108 | 30.17% | PKR 1,319,085 | PKR 291,909 | PKR 25,750 | PKR 317,659 | PKR 1,027,176 |
| Electronics | 301 | 52 | 17.28% | PKR 1,132,142 | PKR 135,466 | PKR 12,300 | PKR 147,766 | PKR 996,676 |
| Home & Kitchen | 224 | 27 | 12.05% | PKR 1,053,536 | PKR 89,680 | PKR 6,250 | PKR 95,930 | PKR 963,856 |
| Sports | 158 | 19 | 12.03% | PKR 364,817 | PKR 26,486 | PKR 4,600 | PKR 31,086 | PKR 338,331 |
| Beauty | 159 | 11 | 6.92% | PKR 333,615 | PKR 20,987 | PKR 2,600 | PKR 23,587 | PKR 312,628 |

Fashion is the story: 108 of the 217 returns and over half of
all money lost sit in one category. Beauty barely returns at all
(11 of 159 orders).

## Return rate by channel

| Channel | Orders | Returns | Return rate | Gross revenue | Refunded | Money lost |
|---|---|---|---|---|---|---|
| Daraz | 385 | 82 | 21.30% | PKR 1,292,346 | PKR 195,639 | PKR 212,039 |
| WhatsApp | 123 | 22 | 17.89% | PKR 432,163 | PKR 44,585 | PKR 50,085 |
| Website | 508 | 86 | 16.93% | PKR 1,855,499 | PKR 250,927 | PKR 272,427 |
| Facebook | 184 | 27 | 14.67% | PKR 623,187 | PKR 73,377 | PKR 81,477 |

Daraz has the worst rate (21.30%) against Facebook at 14.67%.
Website lost the most money in absolute terms (PKR 272,427)
because it carries the most orders — its rate is the second
lowest, so the channel problem is Daraz, not Website.

## Why customers send it back (reasons)

| Reason | Returns | Share of returns | Refunded | Return shipping | Money lost | Share of money lost |
|---|---|---|---|---|---|---|
| Wrong size | 47 | 21.66% | PKR 145,956 | PKR 11,200 | PKR 157,156 | 25.51% |
| Quality not as expected | 40 | 18.43% | PKR 145,163 | PKR 9,350 | PKR 154,513 | 25.08% |
| Changed mind | 40 | 18.43% | PKR 127,256 | PKR 9,950 | PKR 137,206 | 22.27% |
| Late delivery | 33 | 15.21% | PKR 70,776 | PKR 7,450 | PKR 78,226 | 12.70% |
| Defective/damaged | 32 | 14.75% | PKR 42,887 | PKR 7,450 | PKR 50,337 | 8.17% |
| Wrong item sent | 25 | 11.52% | PKR 32,490 | PKR 6,100 | PKR 38,590 | 6.26% |

Wrong size leads on both count and money. Defective/damaged plus
wrong item sent — the returns the store itself caused — cost
PKR 88,927 together, and most of those were replaced, not
refunded, which is why their money lost is small next to their
count.

## Refund vs replacement

150 of 217 returns were refunded (69.12%) and 67 were replaced
(30.88%). Refunds are where the money goes: every refunded order
gives the full order value back, while a replacement only costs
the return shipping. Defective and wrong-item orders are usually
replaced; changed-mind orders are usually refunded.

## Top problem products (by return rate)

| Product | Category | Orders | Returns | Return rate | Money lost | Top reason |
|---|---|---|---|---|---|---|
| Women's Lawn 3pc Suit | Fashion | 96 | 32 | 33.33% | PKR 99,677 | Wrong size |
| Running Sneakers | Fashion | 83 | 27 | 32.53% | PKR 105,378 | Wrong size |
| Men's Kurta — Navy | Fashion | 88 | 26 | 29.55% | PKR 71,174 | Wrong size |
| Smart Watch | Electronics | 77 | 21 | 27.27% | PKR 89,036 | Defective/damaged |
| Kids Frock — Pink | Fashion | 91 | 23 | 25.27% | PKR 41,430 | Wrong size |

The worst product by rate is not the worst by money: Women's
Lawn 3pc Suit returns most often (33.33%), but Running Sneakers
lost the most money (PKR 105,378) — a lower rate on a pricier
product still costs more. Four of the top five are Fashion, and
four of the top five come back for the same reason: wrong size.

## Month by month

| Month | Orders | Returns | Return rate | Gross revenue | Refunded | Money lost | Return-rate change (pts) |
|---|---|---|---|---|---|---|---|
| 2026-01 | 204 | 36 | 17.65% | PKR 699,665 | PKR 89,574 | PKR 98,324 | — |
| 2026-02 | 166 | 24 | 14.46% | PKR 617,802 | PKR 68,780 | PKR 74,530 | -3.19 |
| 2026-03 | 236 | 56 | 23.73% | PKR 755,926 | PKR 114,757 | PKR 127,907 | +9.27 |
| 2026-04 | 201 | 41 | 20.40% | PKR 725,061 | PKR 146,961 | PKR 156,761 | -3.33 |
| 2026-05 | 165 | 25 | 15.15% | PKR 578,410 | PKR 62,581 | PKR 68,531 | -5.25 |
| 2026-06 | 228 | 35 | 15.35% | PKR 826,331 | PKR 81,875 | PKR 89,975 | +0.20 |

March is the worst month by rate (23.73%, +9.27 points on
February) — the sale spike. But the most money was lost in April
(PKR 156,761), when pricier orders came back. Best month:
February (14.46%).

## The four recommendations (from the data, not from a template)

1. **Fix sizing first — it is the biggest leak.** Wrong size is
   the top return reason: 47 of 217 returns (21.66%), costing
   PKR 157,156 (25.51% of all money lost). Fashion is the worst
   category (30.17% return rate, PKR 317,659 lost) and Women's
   Lawn 3pc Suit is the worst product by rate (33.33%, 32 of 96
   orders). Publish real measurements and a size chart on the
   Fashion listings before spending on anything else.
2. **Tighten the Daraz channel.** Daraz has the highest return
   rate at 21.30% (82 of 385 orders), against Facebook at 14.67%,
   and lost PKR 212,039. Check that its listing photos, sizes and
   descriptions match what is actually shipped — the gap against
   the Website channel is a listing problem, not a customer
   problem.
3. **Quality-check before dispatch, starting with Electronics.**
   Defective/damaged (32 returns, PKR 50,337 lost) plus wrong
   item sent (25 returns, PKR 38,590 lost) cost PKR 88,927
   together — every one of those is a return the store caused.
   Running Sneakers alone lost PKR 105,378 (27 returns). A
   one-minute check before sealing the box removes most of it.
4. **Plan for the sale spike before the next one.** 2026-03 had
   the highest return rate of the half year (23.73%, 56 returns)
   and 2026-04 lost the most money (PKR 156,761). Sale months
   bring the volume — and the returns. Staff the returns desk and
   tighten size and quality checks before the next sale, not after
   it.

## Same logic, two ways — and they must agree

The headline KPIs, the category, channel, reason, product and
monthly totals exist in two forms: the Python in `src/analyze.py`
and this workbook's formulas. `src/checks.py` runs both against
the raw file and compares them — a mistake in either one shows up
as a failure. The workbook's formulas were also recalculated in
LibreOffice (including the `_xlfn.MAXIFS` / `_xlfn.MINIFS` worst /
best category return-rate cells): every total matches the
recompute exactly, with zero error values.

## Reproduce it (2 minutes)

```bash
python3 src/generate_data.py   # rebuild the SAMPLE file (seed 20261117)
python3 src/analyze.py         # summary.json + returns_report.md
python3 src/build_workbook.py  # outputs/returns_report_SAMPLE.xlsx
python3 src/build_report.py    # outputs/returns_report.pdf (2 pages)
python3 src/checks.py          # 326 checks — expect all passing
```

On a real client file, replace `data/raw/orders_SAMPLE.csv`
with the orders export (same columns) and skip
`generate_data.py` — everything downstream runs unchanged.
