# Zumi Holdings — Financial Analysis Portfolio

**Prepared by:** Elizabeth Williams (Iwo)
**Role:** Data Analyst | Business Analyst
**Location:** Lagos, Nigeria
**Portfolio:** [lizzyiwo.github.io/ElizabethW_data](https://lizzyiwo.github.io/ElizabethW_data/)
**LinkedIn:** [linkedin.com/in/elizabeth-williams50](https://www.linkedin.com/in/elizabeth-williams50)

---

This repository contains three end-to-end financial analysis projects built for Zumi Holdings, a fictional Nigerian multi-brand technology and commerce holding company operating three subsidiaries: **Zumi Eats** (food delivery), **Zumi Mart** (e-commerce retail), and **Zumi Logistics** (last-mile delivery). Every model was built from raw data using Microsoft Excel, without templates, covering real-world financial scenarios drawn from the Nigerian business environment.

The work here reflects a deliberate combination of accounting discipline and analytical rigour. It is not decorative. Every number has a source, every recommendation connects to a specific finding, and every model is built so that anyone with Excel can open it, trace the logic, and update the assumptions.

---

## Project 1: Zumi Holdings — 24-Month Cashflow and Runway Model

**File:** `Zumi-cashflow_Analysis.xlsx`

### Problem Statement

The Zumi Holdings board is meeting in nine days. The company holds ₦142M in cash across a GTBank NGN operating account and a USD domiciliary account, but the CFO cannot answer a simple question with confidence: how long will that money last? The business has three active debt facilities, a CapEx pipeline worth over ₦80M, an upcoming ₦18.4M corporate income tax instalment, and a receivables book where three customers — Dangote Group, MTN Nigeria, and Flutterwave — represent 62% of all outstanding invoices. The board needs a single document that tells them where they stand, what could go wrong, and what they need to decide before December.

### Methodology

The model is structured across twelve sheets, each handling one layer of the financial picture before feeding into the overall cashflow forecast.

**Historical Monthly** contains 24 months of revenue, COGS, gross profit, OpEx, EBITDA, and cash collected data across all three brands from October 2023 to September 2025. This was used to derive blended growth rates and gross margin trends that anchor the forecast assumptions.

**AR Aging** maps every outstanding receivable by customer, invoice date, due date, and days past due across four ageing buckets: Current, 1-30 days, 31-60 days, and 61-90 days. The total outstanding receivables book stands at approximately ₦29.7M. Dangote Group (₦8.4M), MTN Nigeria (₦6.2M), and Flutterwave (₦3.8M) collectively represent 62% of this balance. Both MTN Nigeria invoices and two Bolt Food aggregator settlements are sitting in the 61-90 day bucket, which is the threshold after which collection risk escalates sharply.

**AP Schedule** captures all outstanding payables by vendor, bill date, due date, and status. The most significant obligation is the Bento Africa payroll at ₦12.4M per month, which is classified as recurring and cannot slip. The combined AP balance due in October 2025 alone is approximately ₦36.58M across ten vendors.

**Loans** documents three active debt facilities: a GTBank revolving overdraft at 26.5% interest with ₦2.35M monthly payment and eight months remaining; a Bank of Industry term loan at 18% for equipment financing with ₦2.01M monthly payment and 41 months remaining; and a Renaissance Capital equipment lease at 22% covering refrigeration units with ₦1.22M monthly and 19 months remaining. Total debt service runs approximately ₦5.58M per month.

**CapEx Plan** tracks five capital commitments totalling ₦80.2M between October 2025 and June 2026. The two committed items — a cold storage truck for Zumi Logistics at ₦14M in October and a POS upgrade for Zumi Mart at ₦6.8M in November — cannot be deferred. The remaining three items (Abuja cloud kitchen at ₦22M, warehouse racking at ₦9.4M, and fleet expansion at ₦28M) are classified as planned and represent optionality the board can use if the worst-case scenario materialises.

**Assumptions** drives the entire forecast from a single source of truth. Base case revenue growth is 2.5% month-on-month, COGS at 66% of revenue, OpEx growing at 1.2% monthly, AR collection at 38 days, and AP payment at 32 days. Upside assumptions include 4.5% revenue growth and a ₦120M equity raise in Q1 2026. Downside assumptions include 0.5% revenue growth, 70% COGS, 55-day AR collection, a 15% FX devaluation shock, and 60-day simultaneous delay from the top-three customers.

**Cash Forecast (12 months)** projects opening cash, cash inflows from AR collections, cash outflows for AP payments, debt service, CapEx, and tax, producing a monthly closing cash balance for each scenario.

**Runway** converts the closing cash balance into a month-count of runway remaining under each scenario.

**Stress Test** isolates the single worst month and the single most dangerous combination of events.

**Board View** distils all of this into one page: the runway chart, the key metrics, and a plain-language verdict written for a reader who does not open spreadsheets daily.

### Key Findings

**Base case (revenue growing at 2.5% per month, normal collections):** Closing cash grows from ₦142M in October 2025 to ₦131.9M by September 2026. Runway exceeds 12 months and cash stays positive throughout. The business is stable under normal operating conditions.

**Worst case (0.5% revenue growth, 15% FX devaluation, top-3 AR customers delay 60 days simultaneously):** Cash turns negative in December 2025, just three months of safe runway from today. The ₦18.4M CIT instalment due on 31 December 2025 is the instrument that finishes it. If receivables from Dangote, MTN, and Flutterwave are delayed by just two months at the same time as the FX shock and reduced revenue, the company cannot cover its December obligations.

**Best case (4.5% revenue growth, ₦400M Series A in January 2026):** Closing cash reaches ₦584M by September 2026. The equity raise fundamentally changes the risk profile and transforms December 2025 from the critical month into a footnote.

**The concentration risk is the real story.** Three customers holding 62% of the AR book is not a cash problem yet, it is a structural vulnerability that the base case conceals. The model makes this visible in a way that a simple bank balance cannot.

### Recommendations

Three actions require board approval before the December board meeting.

First, reduce top-3 AR concentration below 45% immediately. This means accelerating collections from Dangote Group, MTN Nigeria, and Flutterwave, and diversifying the Zumi Logistics contract client base so that no three customers ever again represent the majority of the receivables book.

Second, secure a committed credit facility before October CapEx hits. The ₦20.8M in committed CapEx (cold storage truck plus POS upgrade) in October and November must be funded from somewhere other than operating cash if the worst-case scenario is even possible. A committed overdraft extension or bridging facility buys the board options.

Third, ring-fence the Bento Africa payroll at ₦12.4M per month. This is the one AP obligation in the schedule that cannot slip. Every other vendor can be negotiated with. Payroll cannot. It must be treated as cash-equivalent and reserved before any other disbursement decision is made.

### Tools and Techniques

Microsoft Excel, Power Query, dynamic named ranges, OFFSET and INDEX-MATCH for scenario switching, conditional formatting for stress test flagging, data validation for assumption inputs, chart design for board presentation.

---

## Project 2: Zumi Eats — Unit Economics and Break-Even Model

**File:** `Zumi-eats-cost-model_Building__2_.xlsx`

### Problem Statement

Zumi Eats is losing money on every city it operates in. The management team knows this in a general sense but does not know exactly why, which cost is the biggest problem, or how many orders per day it would take to stop losing money. The CFO has asked for a unit economics model that answers five questions: what is the average contribution margin per order by city; how many orders per day does each city need to break even; is any city currently above that threshold; what happens to the break-even if aggregator commissions rise and food costs increase at the same time; and which three cost items are stealing the most money per order.

### Methodology

The analysis begins at the menu level and builds upward to the network level.

**Menu analysis** covers 24 SKUs across four categories: Mains, Bowls, Sides, and Drinks. Each item has a Menu Price, Food Cost, Packaging Cost, and 90-day order count from the order log. Food cost as a percentage of menu price ranges from 30% for Bottle Water (₦120 cost against ₦400 price) to 43% for the Suya Platter (₦2,940 cost against ₦6,800 price, flagged in the raw data as "High waste"). The best-margin items by category are Chapman in Drinks, the Power Bowl in Bowls, Moi Moi in Sides, and Amala and Ewedu in Mains. The worst-margin items are the Suya Platter and Pepper Soup Goat, both of which carry high input costs relative to their price points.

**Order log analysis** uses 5,663 actual order records from July 1 to September 30, 2025 across Lagos, Abuja, and Port Harcourt. Each record includes order ID, date, city, channel (Direct Mobile App, Direct Website, Glovo, Bolt Food, Jumia Food), average order value, promo code applied, and discount amount. This dataset drives all city-level calculations rather than assumptions.

**Aggregator terms** establish the commission structure that shapes channel profitability. Jumia Food charges 25% commission with a 21-day payout cycle and requires 10% co-marketing spend. Glovo charges 22% with 14-day payout. Bolt Food charges 18% with 10-day payout. The Direct channels (Mobile App and Website) charge zero commission, have immediate settlement, and are funded entirely by Zumi's own marketing spend. Every order routed through an aggregator surrenders between 18% and 25% of revenue to a third party before a single ingredient has been costed.

**City costs** document the fixed monthly overhead for each market. Lagos carries ₦2.84M in monthly fixed costs with 38 active riders and an average daily order volume of 320 (achieved average during the 90-day period). Abuja carries ₦1.64M in fixed costs with 22 riders and 180 daily orders (achieved average). Port Harcourt carries ₦980K in fixed costs with 14 riders and 95 daily orders (achieved average).

**Per-order margin** is calculated for each city by taking the average net revenue per order (gross AOV minus average discount) and deducting average food cost, packaging, rider payout, and aggregator commission. The result is the contribution margin: the amount each order contributes toward covering fixed costs after all variable costs are paid.

**Break-even analysis** divides each city's monthly fixed cost by its contribution margin per order to determine how many orders are needed to cover fixed costs. This is then divided by the days in a month to produce a daily break-even threshold, which is compared against the actual daily order average from the 90-day log.

**Sensitivity analysis** runs a two-way table showing how the network-level daily break-even changes when aggregator commission rates and food costs move simultaneously, across a range of plausible scenarios.

### Key Findings

**Per-order profitability by city:**

Lagos: Average gross AOV of ₦6,022, average discount of ₦366, net revenue of ₦5,656. Variable costs total ₦4,081 (food cost ₦2,470, packaging ₦260, rider payout ₦720, aggregator commission ₦631). Contribution margin per order: ₦1,574.

Abuja: Average gross AOV of ₦6,010, average discount of ₦372, net revenue of ₦5,638. Variable costs total ₦4,049 (food cost ₦2,454, packaging ₦261, rider payout ₦680, aggregator commission ₦653). Contribution margin per order: ₦1,589.

Port Harcourt: Average gross AOV of ₦6,137, average discount of ₦372, net revenue of ₦5,765. Variable costs total ₦4,147 (food cost ₦2,528, packaging ₦258, rider payout ₦740, aggregator commission ₦621). Contribution margin per order: ₦1,618.

Network average contribution margin: ₦1,594 per order.

**Break-even thresholds:**

Lagos needs 60 orders per day to cover its ₦2.84M monthly fixed cost. It is currently achieving 39 orders per day. It is 21 orders below break-even.

Abuja needs 34 orders per day to cover its ₦1.64M monthly fixed cost. It is currently achieving 15 orders per day. It is 19 orders below break-even.

Port Harcourt needs 20 orders per day to cover its ₦980K monthly fixed cost. It is currently achieving 8 orders per day. It is 12 orders below break-even.

The network needs 115 orders per day total. It is receiving 63. Every city is below break-even. Every city loses money before a single naira of profit is possible.

**The three biggest cost levers per order:**
1. Food cost: ₦2,484 average across the network. This is the single largest variable cost, consuming 44% of average net revenue before any other deduction.
2. Rider payout: ₦713 average. This is relatively fixed per city and not easily compressed without affecting service speed or rider retention.
3. Aggregator commission: ₦635 average. Unlike food cost and rider payout, this is not a cost of production. It is a channel tax that only applies to orders that do not come through the Direct app or website.

**Sensitivity:** If aggregator commission rates rise by 3 percentage points simultaneously with a 10% increase in food costs, the network daily break-even rises from 115 to 155 orders per day. The business would need to grow order volume by 35% just to stay in the same loss-making position it occupies today.

### Recommendations

Three things need to happen in a specific order.

First, renegotiate the Glovo and Jumia Food contracts immediately. Every percentage point reduction in Glovo's 22% commission rate returns roughly ₦57 per order directly to the contribution margin. At 5,663 orders over 90 days, that is approximately ₦3.2M recovered per quarter from Glovo alone. Jumia Food at 25% commission with a 21-day payout cycle and mandatory co-marketing spend is the most expensive channel in the mix. Unless Jumia is delivering a meaningfully higher AOV, the relationship needs to be renegotiated or terminated.

Second, review the food cost on high-waste items. The Suya Platter at ₦2,940 food cost against a ₦6,800 menu price has the worst margin in the mains category and is already flagged internally as a high-waste item. Repricing it to ₦7,800 or reformulating the portion size would recover ₦1,000 per order on a SKU that still accounts for 1,620 orders in 90 days.

Third, invest in growing the Direct App channel aggressively. Every order that moves from an aggregator to the Direct Mobile App or Direct Website pays zero commission. That ₦635 average saving per order, applied to the 63 daily orders currently flowing through aggregators, represents ₦40 per order or roughly ₦730K per month recovered without touching a single cost line. At scale, this is not a marginal improvement. It is a structural shift in the business model.

Without these changes, all three cities will remain cash-negative through Q4 and beyond.

### Tools and Techniques

Microsoft Excel, 5,663-row order log analysis, AVERAGEIFS and SUMIFS for city and channel segmentation, contribution margin modelling, two-way sensitivity table, break-even calculation framework, structured verdict narrative.

---

## Project 3: Zumi Holdings — Q3 2025 Group Ledger Reconciliation and P&L

**File:** `Zumi-q3-ledger_completed_Analysis.xlsx`

### Problem Statement

The new analyst has inherited 1,436 transaction records across three source systems: POS-Square (point-of-sale for Zumi Eats and Zumi Mart), NetSuite (ERP for Zumi Logistics contracts), and QuickBooks (expense management for all brands). The CFO needs a clean, auditable Q3 2025 Profit and Loss statement broken down by brand and by month before the end of the week. The raw data has not been reviewed. There are known issues: at least one duplicate vendor payment, refunds that may have been booked incorrectly, and transactions in USD and GHS that need converting to NGN before any P&L figures can be trusted.

### Methodology

The cleaning process followed a documented, sequential approach so that every change is traceable and every number in the final P&L has an unbroken chain back to the raw ledger.

**Chart of accounts mapping** was the first step. The raw transactions carry account codes (4000 to 6080) that needed to be matched to the correct P&L bucket: Revenue, COGS, or Operating Expenses. Twenty account codes across six categories were mapped: Order Revenue, Subscription Revenue, Logistics Contract Revenue, Refunds and Chargebacks (Revenue), Food Cost, Inventory Purchases, Packaging, Fuel and Vehicle, Rider Payouts, Aggregator Commission (COGS), and Salaries and Wages, Rent, Utilities, Marketing, Software and Tools, Professional Services, Bank Fees, Travel, and Office Supplies (OpEx).

**FX conversion** was applied using documented month-end exchange rates. July 2025: USD/NGN 1,530, GHS/NGN 102. August 2025: USD/NGN 1,572, GHS/NGN 99. September 2025: USD/NGN 1,611, GHS/NGN 96. All non-NGN transactions were converted using the rate for the month of the transaction date, with the rate source documented in the FX table.

**Issues identification** required scanning 1,436 rows for data quality problems. Six issues were found totalling ₦6,442,453 in financial impact.

Issue 1: Transaction T101433, dated 12 August 2025, Zumi Mart. A Distell Wholesale vendor payment of ₦1,840,000 appearing twice in the ledger. One entry excluded from the P&L. Action: Excluded.

Issue 2: Transaction T101402, dated 25 August 2025, Zumi Mart. A duplicate payroll entry of ₦680,000 for the same employee in the same period processed through both QuickBooks and a manual journal. Action: Excluded.

Issue 3: Transaction T101419, dated 25 September 2025, originally coded to Zumi Eats. A rider supervisor salary of ₦540,000 misclassified to Zumi Eats rather than Zumi Logistics based on the description and cost centre. Action: Reclassified to Zumi Logistics.

Issue 4: Transaction T101435, dated 22 September 2025. An uncategorised cash withdrawal of ₦180,000 with no GL code and no description beyond "cash out." Action: Parked in a separate holding line with a note for the CFO to investigate and explain before Q4 closes.

Issue 5: Transaction T101434, Zumi Logistics. A fuel invoice of ₦240,000 with a missing transaction date that cannot be assigned to any specific month in Q3. Action: Excluded from the P&L with documentation.

Issue 6: Multiple account 4900 entries across all brands. A pattern of refunds and chargebacks booked as positive revenue figures rather than negative. Total financial impact ₦2,962,453 across the quarter. Action: Fixed in the Clean Amount formula so that all 4900-coded transactions are treated as revenue reductions.

**Brand-level P&L construction** used the cleaned transaction table filtered by brand and month to build three separate P&L statements (Zumi Eats, Zumi Mart, Zumi Logistics) and one consolidated group statement, all following the same structure: Net Revenue, COGS by line, Gross Profit, Operating Expenses by line, Net Profit.

**Bank reconciliation** matched the P&L-derived net cash movement for each brand against the bank statement totals for the GTBank NGN operating account and the USD domiciliary account across all three months.

### Key Findings

**Did Zumi Holdings make money in Q3 2025? No. All three brands operated at a loss.**

**Group consolidated result:**

Net Revenue: ₦50,072,379
Total COGS: ₦96,426,000
Gross Profit: (₦46,353,621)
Total Operating Expenses: ₦52,429,362
Net Profit: (₦98,782,983)

The group lost ₦98.8M in Q3 2025.

**By brand:**

Zumi Eats: Net Revenue ₦13,080,392. Total COGS ₦25,444,000. Gross Profit (₦12,363,608). Total OpEx ₦20,515,449. Net Profit (₦32,879,057).

Zumi Mart: Net Revenue ₦20,263,056. Total COGS ₦47,763,000. Gross Profit (₦27,499,944). Total OpEx ₦21,226,494. Net Profit (₦48,726,438). Inventory purchases alone at ₦46,579,000 exceed total net revenue by 2.3 times.

Zumi Logistics: Net Revenue ₦16,728,931. Total COGS ₦23,219,000. Gross Profit (₦6,490,069). Total OpEx ₦10,687,419. Net Profit (₦17,177,488). Closest to breakeven of the three brands and carries the strongest revenue base relative to fixed costs.

**Data quality issues summary:** Six issues found with a combined financial impact of ₦6,442,453. The largest single issue was the pattern of refunds booked as positive revenue at ₦2,962,453, which would have overstated net revenue by 5.9% if left uncorrected. The duplicate Distell Wholesale invoice at ₦1,840,000 would have overstated Zumi Mart's COGS and understated its already-negative gross profit further.

### Recommendations

Zumi Holdings lost ₦98.8M in Q3 2025. The finding is stark and the causes are traceable.

Zumi Mart is the most urgent problem. Its inventory costs exceed its revenue by 2.3 times. This is not an efficiency problem, it is a structural mismatch between what the business is buying and what it is selling. The stock mix, supplier pricing, and pricing-to-market strategy need a complete review before Q4 purchasing decisions are made. Buying ₦46.6M of inventory to generate ₦20.3M of revenue is not a slow business, it is a failing business model.

Zumi Eats carries the highest OpEx burden at ₦20.5M against gross revenue of ₦13.1M. OpEx is exceeding gross revenue, which means the business is spending more on running itself than it is generating from selling food. Salary costs, rent, and marketing need an immediate line-by-line review.

Zumi Logistics is the only brand that demonstrates a path to profitability. Its net revenue of ₦16.7M against COGS of ₦23.2M gives a negative gross profit, but the gap is the smallest of the three brands and the revenue base is the most consistent (contract-backed with Dangote, MTN, and Flutterwave rather than consumer-demand-driven). Reducing fuel and vehicle costs and growing the contract client base are the levers.

The finance team must implement duplicate payment controls before Q4 closes. Two of the six data quality issues (the Distell invoice duplicate and the payroll duplicate) represent genuine double payments that may already have left the company's bank account. If they have, the ₦2.52M is recoverable but requires immediate action. If they have not, the controls need to be in place to stop them processing.

### Tools and Techniques

Microsoft Excel, Power Query, 1,436-row ledger cleaning, VLOOKUP and INDEX-MATCH for chart of accounts mapping, documented Issues Log with action trail, FX conversion table with month-specific rates, brand-level P&L construction, bank reconciliation, CEO dashboard design.

---

## Repository Structure

```
financial-analysis-portfolio/
├── Zumi-cashflow_Analysis.xlsx          12-sheet cashflow and runway model
├── Zumi-eats-cost-model_Building__2_.xlsx   Unit economics and break-even model
├── Zumi-q3-ledger_completed_Analysis.xlsx   Ledger reconciliation and P&L
├── screenshots/
│   ├── cashflow_dashboard_png.png
│   ├── unit_economics_png.png
│   └── q3_ledger_dashboard_png.png
└── README.md
```

---

## About Elizabeth Williams

I am a Data Analyst and Business Analyst based in Lagos, Nigeria, with over five years of professional experience in financial management and accounting. I completed certifications in Data Analysis and Business Analysis in 2025 and built these models to demonstrate the kind of financial analysis work I do: grounded in real accounting practice, built on verifiable data, and written so that the findings are usable rather than impressive.

My financial modelling work is not theoretical. I understand these numbers because I have spent five years working with numbers like them in real organisations, under real pressure, with real consequences for getting them wrong.

**Contact:** lizzyiwo@gmail.com
**Portfolio:** [lizzyiwo.github.io/ElizabethW_data](https://lizzyiwo.github.io/ElizabethW_data/)
**LinkedIn:** [linkedin.com/in/elizabeth-williams50](https://www.linkedin.com/in/elizabeth-williams50)

