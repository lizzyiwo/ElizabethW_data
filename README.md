# Financial Analysis Portfolio — Elizabeth Williams

**Data Analyst | Business Analyst | Lagos, Nigeria**
[Portfolio Website](https://lizzyiwo.github.io/ElizabethW_data/) | [LinkedIn](https://www.linkedin.com/in/elizabeth-williams50) | lizzyiwo@gmail.com

---

## About This Repository

This repository contains three financial analysis projects completed as part of my professional development in data analytics and business intelligence. Each project was built using real-world financial modelling frameworks applied to a fictional Nigerian holding company called Zumi Holdings and its subsidiaries.

The work here reflects my background as an accountant of five years combined with recent training in data analysis, Python, and business analysis. Every model was built from scratch in Microsoft Excel, with no templates used.

---

## Project 1: Zumi Holdings — 24-Month Cashflow and Runway Analysis

**File:** `ElizabethWilliams_zumi-cashflow_Analysis.xlsx`

### What This Project Is About

Zumi Holdings is a fictional Nigerian tech holding company with three business lines: Zumi Eats (food delivery), Zumi Mart (e-commerce), and Zumi Logistics (last-mile delivery). This project models the company's full cashflow position across a 24-month period and answers the most important question any business needs to answer: how long can we survive, and under what conditions?

### What the Workbook Contains

The workbook is built across 12 sheets covering every stage of the cashflow analysis:

- **Revenue Model** — Monthly revenue breakdown by business line, based on assumed growth rates and seasonal adjustments
- **Cost Structure** — Fixed and variable cost modelling for each subsidiary, including headcount, operations, and marketing spend
- **EBITDA Summary** — Earnings before interest, tax, depreciation, and amortisation calculated monthly across all three lines
- **Accounts Receivable Ageing** — Outstanding invoice tracking with 30, 60, and 90-day ageing buckets
- **Accounts Payable Schedule** — Payment obligations mapped by due date to protect liquidity
- **12-Month Cash Forecast** — Month-by-month projected cash inflows and outflows with opening and closing balances
- **Runway Analysis** — Calculation of how many months of runway the business has under current burn rates
- **Assumptions Sheet** — All key inputs and drivers in one place so the model can be updated easily
- **Stress Testing** — Three full scenarios modelled: best case (20% revenue uplift), base case (current trajectory), and worst case (30% revenue decline with cost pressures)
- **Sensitivity Analysis** — How changes in individual variables (pricing, churn, cost of goods) affect the runway figure
- **Variance Tracker** — Comparison of actuals against forecast for accountability
- **Board Dashboard** — A single-page executive summary with key metrics, charts, and a plain-language narrative for non-financial stakeholders

### Key Skills Demonstrated

- Multi-entity financial modelling across three business lines
- Cash forecasting with dynamic assumptions
- Accounts receivable and payable scheduling
- Scenario and sensitivity analysis
- Executive dashboard design
- Advanced Excel functions including dynamic named ranges, OFFSET, INDEX MATCH, and conditional formatting logic

### Tools Used

Microsoft Excel, Power Query, PivotTables

### Screenshot

![Cashflow Dashboard](screenshots/cashflow_dashboard.png)

---

## Project 2: Zumi Eats — Unit Economics and Break-Even Model

**File:** `ElizabthWilliams_zumi-eats-cost-model_Building.xlsx`

### What This Project Is About

Unit economics is the discipline of understanding whether a single transaction makes money before you worry about whether the whole business does. For a food delivery service like Zumi Eats, this means knowing the exact profitability of every order placed, in every city, for every item on the menu. This model answers that question in detail for operations across Lagos, Abuja, and Port Harcourt.

### What the Workbook Contains

- **Menu Analysis** — Food cost and gross margin calculated for all 25 items on the menu, from ingredient cost through to customer-facing price
- **Contribution Margin by SKU** — Which items are pulling their weight and which are not
- **City Cost Model** — Rider payout rates, fuel costs, platform fees, and delivery logistics costs broken out separately for Lagos, Abuja, and Port Harcourt, because the numbers are meaningfully different in each city
- **Per-Order Profitability** — Revenue minus food cost minus delivery cost minus platform fee for every combination of city and menu item
- **Break-Even Analysis** — How many orders per day Zumi Eats needs to cover its fixed costs at city level and at group level
- **Sensitivity Analysis** — What happens to profitability if food costs rise by 10%, if rider payouts increase, or if the delivery fee to customers changes
- **Recommendation Summary** — A plain-language verdict on which cities are profitable, which menu items to push or retire, and what levers the business should pull first

### Key Skills Demonstrated

- Unit economics modelling from first principles
- Multi-location cost comparison
- Break-even and contribution margin analysis
- Sensitivity modelling on operational cost drivers
- Translating financial analysis into business recommendations

### Tools Used

Microsoft Excel

### Screenshot

![Unit Economics Model](screenshots/unit_economics.png)

---

## Project 3: Zumi Holdings — Q3 2025 Group Ledger Reconciliation and P&L

**File:** `ElizabethWilliams_zumi-q3-ledger_completed_Analysis.xlsx`

### What This Project Is About

Raw financial data almost never arrives clean. It comes with duplicates, inconsistent category labels, missing values, transactions recorded in different currencies, and entries that have been posted to the wrong account. This project starts with exactly that kind of messy raw data and ends with a clean, auditable Profit and Loss report that a finance director or CEO could present to a board.

This is the kind of work that sits at the core of financial accounting and management reporting, and it is where a strong grasp of both accounting principles and data cleaning techniques genuinely matters.

### What the Workbook Contains

- **Raw Data Sheet** — The original transaction-level data as received, including errors, duplicates, FX entries in USD and GBP alongside NGN, and miscategorised expenses
- **Data Cleaning Log** — A documented record of every change made to the raw data, including what was changed, why, and what the corrected value is. This is standard practice in audit-ready financial reporting
- **Duplicate Detection** — Flagged and resolved duplicate transaction entries using COUNTIFS logic
- **FX Conversion** — All non-NGN transactions converted to Nigerian Naira at the relevant Q3 2025 exchange rates, with source rates documented
- **Chart of Accounts Mapping** — Every transaction category mapped to the correct account code across assets, liabilities, revenue, cost of goods sold, and operating expenses
- **Brand-Level P&L** — Separate Profit and Loss statements for Zumi Eats, Zumi Mart, and Zumi Logistics, each showing revenue, gross profit, operating expenses, and net profit by month across July, August, and September 2025
- **Group Consolidated P&L** — All three brands consolidated into a single group-level statement
- **Bank Reconciliation** — Closing cash balances reconciled against the bank statement figures for each entity
- **CEO Dashboard** — A single-page visual summary showing group revenue, gross margin, operating expenses, and net profit with month-on-month trend lines, designed for a non-technical reader

### Key Skills Demonstrated

- Raw financial data cleaning and error detection
- FX conversion and multi-currency reconciliation
- Chart of accounts mapping
- Month-by-month P&L construction by entity
- Bank reconciliation
- Consolidated group financial reporting
- Executive dashboard design for non-financial audiences

### Tools Used

Microsoft Excel, Power Query

### Screenshot

![Q3 Ledger Dashboard](screenshots/q3_ledger_dashboard.png)

---

## About Me

I am Elizabeth Williams (Iwo), a Data Analyst and Business Analyst based in Lagos, Nigeria. I have over five years of professional experience in financial management and accounting, and in 2025 I completed certifications in Data Analysis (ALX Africa), Business Analysis (BlackForce), Python Programming (ALX Africa), and Data Analysis Bootcamp (Women Techsters / Tech4Dev).

My financial modelling work is grounded in real accounting practice. I understand these numbers not just as data, but as business decisions waiting to be made.

**Links**
- Portfolio: https://lizzyiwo.github.io/ElizabethW_data/
- LinkedIn: https://www.linkedin.com/in/elizabeth-williams50
- Email: lizzyiwo@gmail.com
