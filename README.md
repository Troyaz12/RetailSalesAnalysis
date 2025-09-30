# RetailSalesAnalysis

## Data Cleaning
- Filled all numeric columns that were NA with the median
- Fill missing Customer IDs with 'Unknown'
- Fill missing descriptions with 'No Description Available'
- Drop the Invoice column
- Break InvoiceDate out by year, month, weekday and hour
- export cleaned dataset to .csv for analysis

## This project analyzes retail data. The data contains information such as Quantity sold, invoice dates, pricing, country and customer transactions.

## Dataset Description
columns      Description
Invoice      Primary Identifier of an order
Stock Code   Identifier for an item 
Description  Description of item
Quantity     Quantity of item sold per order
Invoice Date Date of the customers order
Price        Price of the Item
CustomerID   Unique Identifier of the customer
Country      Country the customer resides in

## Key Insights

## 1. Top Customers Gross Sales by Quantity
- Shows top ten customers total purchased items, ignoring cancelations and returns.
- Helps to identify demand consentration across the customer base.
- The top ten cusomers accounted for 17% of the gross sales for the time period.
- The single largest customer contributed 3% to gross sales alone.

## 2. Top Products Gross Sales by Quantity
- Shows top ten products sold, ignoring cancelations and returns.
- Useful for understanding product popularity.
- The top ten products accounted for 6% of sales by product.
- The largest product by gross sales accounted for 1% of sales alone.

## 3. Gross Sales Trends Over Time by Quantity
- Monthly sales trends, ignoring cancelations and returns.
- Captures seasonality and demand spikes.
- Sales spiked in September through October for both 2010 and 2011.
- 2010 showed a steady performance with additional spikes in March and August.
- 2011 saw weak sales in February, April, and December compared to the prior year.

## Methodology