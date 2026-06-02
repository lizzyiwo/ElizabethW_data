# Zumi Holdings — Financial Analysis Portfolio

**Prepared by:** Elizabeth Williams (Iwo) | Data Analyst & Business Analyst | Lagos, Nigeria

[![Portfolio](https://img.shields.io/badge/Portfolio-Visit%20Site-b8924a?style=flat-square)](https://lizzyiwo.github.io/ElizabethW_data/)
[![LinkedIn](https://img.shields.io/badge/LinkedIn-Connect-0077B5?style=flat-square&logo=linkedin)](https://www.linkedin.com/in/elizabeth-williams50)
[![Email](https://img.shields.io/badge/Email-Get%20In%20Touch-D14836?style=flat-square&logo=gmail)](mailto:lizzyiwo@gmail.com)

---

This repository contains three end-to-end financial analysis projects built for **Zumi Holdings**, a fictional Nigerian multi-brand technology and commerce holding company with three subsidiaries: **Zumi Eats** (food delivery), **Zumi Mart** (e-commerce retail), and **Zumi Logistics** (last-mile delivery). Every model was built from raw data in Microsoft Excel without templates. Each project follows the same structure: a clear problem statement, documented methodology, real findings with exact figures, and actionable recommendations grounded in the data.

---

## Table of Contents

- [Project 1: 24-Month Cashflow and Runway Model](#project-1)
- [Project 2: Unit Economics and Break-Even Model](#project-2)
- [Project 3: Q3 2025 Group Ledger Reconciliation and P&L](#project-3)
- [Repository Structure](#repository-structure)
- [About](#about-elizabeth-williams)

---

<a name="project-1"></a>
## Project 1: Zumi Holdings — 24-Month Cashflow and Runway Model

**File:** [Zumi-cashflow_Analysis.xlsx](https://github.com/lizzyiwo/financial-analysis-portfolio/raw/main/Zumi-cashflow_Analysis.xlsx)

### Dashboard

![Zumi Holdings Cashflow and Runway — Board View](https://raw.githubusercontent.com/lizzyiwo/ElizabethW_data/main/screenshots/cashflow_dashboard.png.png)

> **Board View:** 12-Month Cash Runway by Scenario. Base Case (blue) stays positive 12+ months and closes at N131.9M by September 2026. Worst Case (red) turns cash negative in December 2025. Best Case (green) reaches N584M after the N400M Series A raise in January 2026.

---

### Problem Statement

The Zumi Holdings board meets in nine days. The company holds N142M in cash across a GTBank NGN operating account and a USD domiciliary account, but the CFO cannot answer the most important question with confidence: how long will that money last? The business has three active debt facilities, a CapEx pipeline worth over N80M, an upcoming N18.4M corporate income tax instalment due 31 December 2025, and a receivables book where three customers — Dangote Group, MTN Nigeria, and Flutterwave — represent 62% of all outstanding invoices. The board needs one document that tells them where they stand, what could go wrong, and what decisions must be approved before December.

### Methodology

The model is structured across twelve interconnected sheets.

**Historical Monthly** contains 24 months of revenue, COGS, gross profit, OpEx, EBITDA, and cash collected across all three brands from October 2023 to September 2025. These figures anchor the growth rate and margin assumptions used in the forecast.

**AR Ageing** maps every outstanding receivable across four ageing buckets: Current, 1-30 days, 31-60 days, and 61-90 days. Total outstanding receivables: N29.7M. Dangote Group (N8.4M), MTN Nigeria (N6.2M), and Flutterwave (N3.8M) hold 62% of this balance. Both MTN Nigeria invoices sit in the 61-90 day bucket, the threshold at which collection risk escalates sharply.

**AP Schedule** tracks all outstanding payables across ten vendors. The Bento Africa payroll at N12.4M per month is classified as recurring and cannot slip. Combined AP due in October 2025 alone totals N36.58M.

**Loans** documents three active debt facilities: a GTBank revolving overdraft at 26.5% (N2.35M monthly, 8 months remaining); a Bank of Industry term loan at 18% (N2.01M monthly, 41 months remaining); and a Renaissance Capital equipment lease at 22% (N1.22M monthly, 19 months remaining). Total monthly debt service: N5.58M.

**CapEx Plan** tracks N80.2M in capital commitments across October 2025 to June 2026. The two committed items — a cold storage truck (N14M, October) and a POS upgrade (N6.8M, November) — cannot be deferred. Three planned items total N59.4M and represent optionality the board can exercise or defer depending on the scenario.

**Assumptions** drives the entire forecast from a single source of truth. Base case: 2.5% monthly revenue growth, 66% COGS, 1.2% monthly OpEx growth, 38-day AR collection, 32-day AP payment. Worst case: 0.5% revenue growth, 70% COGS, 55-day AR collection, 15% FX devaluation, 60-day simultaneous delay from all three top customers.

**Stress Test** isolates the single worst month and the most dangerous combination of events that could occur simultaneously.

**Board View** distils all twelve sheets into one page: the scenario chart, key metrics, and a plain-language verdict written for a non-financial reader.

### Key Findings

**Base case:** Closing cash grows from N142M in October 2025 to N131.9M by September 2026. Runway exceeds 12 months throughout. The business is stable under normal operating conditions.

**Worst case:** Cash turns negative in December 2025 — just three months of safe runway. The N18.4M CIT instalment due 31 December 2025 is the instrument that finishes it. If Dangote, MTN Nigeria, and Flutterwave delay payment by two months simultaneously during a 15% FX shock and a 20% revenue decline, the company cannot cover its December obligations.

**Best case:** With a N400M Series A closing in January 2026, closing cash reaches N584M by September 2026. The equity raise transforms December 2025 from the critical month into a footnote.

**The concentration risk is the real story.** Three customers holding 62% of the AR book is not a cash problem yet — it is a structural vulnerability that the base case conceals and the worst case exposes.

### Recommendations

**Reduce top-3 AR concentration below 45% immediately.** Accelerate collections from Dangote Group, MTN Nigeria, and Flutterwave, and diversify Zumi Logistics' contract client base so that no three customers again hold the majority of receivables.

**Secure a committed credit facility before October CapEx hits.** The N20.8M in committed CapEx in October and November must have a funding source other than operating cash if the worst case is even possible. A committed overdraft extension or bridging facility buys the board options.

**Ring-fence the Bento Africa payroll at N12.4M per month.** This is the one AP obligation that cannot slip. Every other vendor can be negotiated with. Payroll cannot. It must be reserved before any other disbursement decision is made.

### Tools and Techniques

Microsoft Excel, Power Query, dynamic named ranges, OFFSET and INDEX-MATCH for scenario switching, conditional formatting for stress test flagging, data validation for assumption inputs.

---

<a name="project-2"></a>
## Project 2: Zumi Eats — Unit Economics and Break-Even Model

**File:** [Zumi-eats-cost-model_Building__2_.xlsx](https://github.com/lizzyiwo/financial-analysis-portfolio/raw/main/Zumi-eats-cost-model_Building__2_.xlsx)

### Dashboard

![Zumi Eats Unit Economics and Break-Even Verdict](https://raw.githubusercontent.com/lizzyiwo/ElizabethW_data/main/screenshots/unit_economics.png.png)

> **Verdict Sheet:** All three cities are below break-even. The network needs 115 orders per day but receives only 63. Top 3 cost levers per order: Food Cost N2,484 | Rider Payout N713 | Aggregator Commission N635.

---

### Problem Statement

Zumi Eats is losing money in every city it operates in. Management knows this in a general sense but does not know exactly why, which cost is the biggest problem, or how many orders per day it would take to stop losing money. The CFO needs a unit economics model that answers five specific questions: what is the average contribution margin per order by city; how many orders per day does each city need to break even; is any city currently above that threshold; what happens to the break-even if aggregator commissions rise and food costs increase simultaneously; and which three cost items are stealing the most money per order.

### Methodology

The analysis builds from the menu level upward to the network level.

**Menu analysis** covers 24 SKUs across four categories — Mains, Bowls, Sides, and Drinks. Food cost as a percentage of menu price ranges from 30% for Bottle Water to 43% for the Suya Platter, which is already flagged internally as a high-waste item.

**Order log analysis** uses 5,663 actual order records from 1 July to 30 September 2025 across Lagos, Abuja, and Port Harcourt. Each record includes city, channel (Direct Mobile App, Direct Website, Glovo, Bolt Food, Jumia Food), average order value, promo code, and discount amount. This dataset drives all city-level calculations rather than assumptions.

**Aggregator terms** establish the commission structure that shapes channel profitability: Jumia Food 25% commission with 21-day payout and mandatory 10% co-marketing spend; Glovo 22% with 14-day payout; Bolt Food 18% with 10-day payout; Direct channels pay zero commission. Every aggregator order surrenders between 18% and 25% of revenue to a third party before any ingredient is costed.

**City costs** document fixed monthly overhead by market. Lagos: N2.84M, 38 riders, 39 average daily orders achieved. Abuja: N1.64M, 22 riders, 15 average daily orders. Port Harcourt: N980K, 14 riders, 8 average daily orders.

**Per-order margin** deducts food cost, packaging, rider payout, and aggregator commission from net AOV to produce the contribution margin per order.

**Break-even analysis** divides each city's monthly fixed cost by its contribution margin to produce a daily orders threshold, then compares that against the actual daily average from the 90-day order log.

**Sensitivity analysis** runs a two-way table across ranges of aggregator commission rates and food cost changes to show how the network break-even threshold shifts.

### Key Findings

| Metric | Lagos | Abuja | Port Harcourt | Network Avg |
|--------|-------|-------|---------------|-------------|
| Avg Net Revenue/Order | N5,656 | N5,638 | N5,765 | N5,686 |
| Food Cost | N2,470 | N2,454 | N2,528 | N2,484 |
| Packaging | N260 | N261 | N258 | N260 |
| Rider Payout | N720 | N680 | N740 | N713 |
| Aggregator Commission | N631 | N653 | N621 | N635 |
| **Contribution Margin** | **N1,574** | **N1,589** | **N1,618** | **N1,594** |

| City | Break-Even Orders/Day | Actual Orders/Day | Gap |
|------|-----------------------|-------------------|-----|
| Lagos | 60 | 39 | 21 below |
| Abuja | 34 | 15 | 19 below |
| Port Harcourt | 20 | 8 | 12 below |
| **Network Total** | **115** | **63** | **52 below** |

Every city is below break-even. Every day the business opens, it loses money before a single naira of profit is possible.

If aggregator commission rates rise by 3 percentage points simultaneously with a 10% food cost increase, the daily network break-even climbs from 115 to 155 orders. The business would need 35% volume growth just to stay in the same loss-making position it occupies today.

### Recommendations

**Renegotiate Glovo and Jumia Food contracts immediately.** Every 1% reduction in Glovo's 22% commission returns approximately N57 per order. At 5,663 orders over 90 days, that is N3.2M recovered per quarter from Glovo alone. Jumia Food at 25% commission with mandatory co-marketing spend is the most expensive channel in the network.

**Review food cost on high-waste items.** The Suya Platter at N2,940 food cost is already flagged as high waste. Repricing from N6,800 to N7,800 recovers N1,000 per order across approximately 1,620 orders per 90 days — N1.62M per cycle.

**Grow the Direct App channel aggressively.** Every order shifted from an aggregator to the Direct Mobile App pays zero commission. The N635 average commission saving per order, applied to the 63 daily orders currently flowing through aggregators, represents approximately N730K per month recovered without adjusting a single cost line.

### Tools and Techniques

Microsoft Excel, 5,663-row order log analysis, AVERAGEIFS and SUMIFS for city and channel segmentation, contribution margin modelling, two-way sensitivity table, break-even calculation framework.

---

<a name="project-3"></a>
## Project 3: Zumi Holdings — Q3 2025 Group Ledger Reconciliation and P&L

**File:** [Zumi-q3-ledger_completed_Analysis.xlsx](https://github.com/lizzyiwo/financial-analysis-portfolio/raw/main/Zumi-q3-ledger_completed_Analysis.xlsx)

### Dashboard

![Zumi Holdings Q3 2025 Financial Summary — CEO View](https://raw.githubusercontent.com/lizzyiwo/ElizabethW_data/main/screenshots/q3_ledger_dashboard.png.png)

> **CEO Summary Sheet:** Group net loss of N98.8M in Q3 2025. All three brands operated at a loss. Six data quality issues identified and corrected totalling N6,442,453.

---

### Problem Statement

1,436 transaction records across three source systems — POS-Square, NetSuite, and QuickBooks — need to be cleaned and reconciled into an auditable Q3 2025 Profit and Loss statement by brand and by month. The raw data has not been reviewed. Known issues include at least one duplicate vendor payment, refunds that may have been booked incorrectly, and transactions in USD and GHS that need converting to NGN before any P&L figure can be trusted.

### Methodology

**Chart of accounts mapping** matched every transaction code (4000 to 6080) to the correct P&L bucket: Revenue, COGS, or Operating Expenses. Twenty codes across six categories were mapped and documented.

**FX conversion** used documented month-end rates: July (USD/NGN 1,530; GHS/NGN 102), August (USD/NGN 1,572; GHS/NGN 99), September (USD/NGN 1,611; GHS/NGN 96). All non-NGN transactions were converted at the rate of the transaction month with the source rate documented.

**Issues identification** scanned 1,436 rows for data quality problems. Six were found totalling N6,442,453 in financial impact:

| Issue | Brand | Amount (NGN) | Action Taken |
|-------|-------|-------------|--------------|
| Duplicate Distell Wholesale invoice | Zumi Mart | N1,840,000 | Excluded from P&L |
| Duplicate payroll entry — Bento Africa | Zumi Mart | N680,000 | Excluded from P&L |
| Salary misclassified — Rider supervisor | Zumi Eats | N540,000 | Reclassified to Zumi Logistics |
| Uncategorised cash withdrawal | Zumi Eats | N180,000 | Parked — CEO must explain |
| Missing transaction date — Fuel invoice | Zumi Logistics | N240,000 | Excluded — cannot assign to month |
| Refunds booked as positive revenue | All brands | N2,962,453 | Fixed in Clean_Amount formula |
| **Total financial impact** | | **N6,442,453** | |

**Brand-level P&L construction** used the cleaned transaction table filtered by brand and month to produce separate statements for Zumi Eats, Zumi Mart, and Zumi Logistics plus a consolidated group statement.

**Bank reconciliation** matched the P&L-derived net cash movement against GTBank NGN operating and USD domiciliary account statements across all three months.

### Key Findings

| | Zumi Eats | Zumi Mart | Zumi Logistics | GROUP TOTAL |
|-|-----------|-----------|----------------|-------------|
| Net Revenue | N13,080,392 | N20,263,056 | N16,728,931 | N50,072,379 |
| Total COGS | (N25,444,000) | (N47,763,000) | (N23,219,000) | (N96,426,000) |
| Gross Profit | (N12,363,608) | (N27,499,944) | (N6,490,069) | (N46,353,621) |
| Total OpEx | (N20,515,449) | (N21,226,494) | (N10,687,419) | (N52,429,362) |
| **Net Profit** | **(N32,879,057)** | **(N48,726,438)** | **(N17,177,488)** | **(N98,782,983)** |

All three brands operated at a loss. Zumi Holdings lost N98.8M in Q3 2025.

Zumi Mart is the largest loss-maker. Its inventory purchases of N46,579,000 exceed total net revenue of N20,263,056 by 2.3 times. The refund misclassification of N2,962,453 would have overstated net revenue by 5.9% if left uncorrected.

### Recommendations

**Zumi Mart requires immediate structural intervention.** Buying N46.6M of inventory to generate N20.3M of revenue is not a slow business — it is a failing business model. Stock mix, supplier pricing, and pricing-to-market strategy need a complete review before Q4 purchasing decisions are made.

**Zumi Eats OpEx exceeds its gross revenue.** At N20.5M OpEx against N13.1M gross revenue, the business spends more on running itself than it generates from selling food. Salary costs, rent, and marketing need an immediate line-by-line review.

**Zumi Logistics has the clearest path to profitability.** Contract-backed revenue with Dangote, MTN Nigeria, and Flutterwave is the most predictable of the three brands. Reducing fuel costs and growing the contract base are the right levers.

**Implement duplicate payment controls before Q4 closes.** The Distell invoice duplicate (N1.84M) and the payroll duplicate (N680K) may have already left the bank account. If so, they are recoverable but require immediate action.

### Tools and Techniques

Microsoft Excel, Power Query, 1,436-row ledger cleaning, VLOOKUP and INDEX-MATCH for chart of accounts mapping, documented Issues Log with full audit trail, FX conversion with month-specific rates, brand-level P&L construction, bank reconciliation.

---

<a name="repository-structure"></a>
## Repository Structure

```
financial-analysis-portfolio/
├── Zumi-cashflow_Analysis.xlsx
├── Zumi-eats-cost-model_Building__2_.xlsx
├── Zumi-q3-ledger_completed_Analysis.xlsx
└── README.md
```

---

<a name="about-elizabeth-williams"></a>
## About Elizabeth Williams

I am a Data Analyst and Business Analyst based in Lagos, Nigeria, with over five years of professional experience in financial management and accounting. I completed certifications in Data Analysis (ALX Africa), Business Analysis (BlackForce), and Python Programming in 2025, and built these models to demonstrate financial analysis grounded in real accounting practice.

**Portfolio:** [lizzyiwo.github.io/ElizabethW_data](https://lizzyiwo.github.io/ElizabethW_data/)
**LinkedIn:** [linkedin.com/in/elizabeth-williams50](https://www.linkedin.com/in/elizabeth-williams50)
**Email:** [lizzyiwo@gmail.com](mailto:lizzyiwo@gmail.com)
**GitHub:** [github.com/lizzyiwo](https://github.com/lizzyiwo)
