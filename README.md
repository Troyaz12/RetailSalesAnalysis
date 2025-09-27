# RetailSalesAnalysis

## Data Cleaning
- Filled all numeric columns that were NA with the median
- Fill missing Customer IDs with 'Unknown'
- Fill missing descriptions with 'No Description Available'
- Drop the Invoice column
- Break InvoiceDate out by year, month, weekday and hour
- export cleaned dataset to .csv for analysis