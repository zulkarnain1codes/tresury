# Treasury Cash Dashboard

Website dashboard fed directly by your Excel workbooks.

## Run
    npm install
    npm start          # then open http://localhost:3000

## Where it reads the data
Every `.xlsx` in the project home folder (change `dataDir` in `config.json`, or set `DATA_DIR`).
One workbook per customer, same structure as the Pando file, named:

    Treasury_Cash_Dashboard_<Customer>.xlsx      ->  customer name = text after the last underscore group

To add a customer: copy their workbook into the folder and press **Refresh from files**.

## Weekly / monthly routine
1. In the workbook, enter Received Amount / Received Date (Fact_Receipts) and Paid (USD) / Paid Date (Fact_Payables); add new orders and POs as new rows.
2. **Save in Excel** (the dashboard reads the values Excel stores).
3. Press Refresh (it also auto-refreshes within a minute of the file changing).

## What it reads
- Fact_Receipts / Fact_Payables: orders, tranches, supplier POs, statuses.
- Dashboard sheet: the cash waterfall (receipts, payments, closing cash, deferred payables, net position) by month.
- Assumptions!C4 (report date).

## Tabs
Overview, Profitability, Working capital, Orders & payables. Filters: customer chips (multi-select, "All customers" = compiled) and period (full timeline, FY Apr-Mar, calendar year, actuals, forecast).
