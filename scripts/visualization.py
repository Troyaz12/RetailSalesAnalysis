import pandas as pd
import matplotlib.pyplot as plt

path = r"C:\Users\Troy\OneDrive\Desktop\python\datasets\Online Retail\data\out.csv"
retailData = pd.read_csv(path)
retailData['InvoiceDate'] = pd.to_datetime(retailData['InvoiceDate'])

sales = retailData[retailData['Quantity']>0]

monthly_sales = sales.resample('ME', on='InvoiceDate')['Quantity'].sum()

ax = monthly_sales.plot(kind='bar', figsize=(12,6))
fig = ax.get_figure()
ax.set_xticklabels([d.strftime('%b %Y') for d in monthly_sales.index], rotation=45, ha='right')

plt.ylabel('Quantity Sold')
plt.title('Monthly Sales')
plt.tight_layout()

chartPathMonthlySales = 'data\MonthlySales_chart.png'
plt.savefig(chartPathMonthlySales)
plt.close(fig)
#plt.show()

monthly_sales_df = monthly_sales.reset_index()
monthly_sales_df['InvoiceDate'] = monthly_sales_df['InvoiceDate'].dt.strftime('%Y-%m-%d')

monthly_sales_df.columns = ['Invoice Date', 'Quantity']
# Add Ranking column starting at 1
monthly_sales_df.insert(0, 'Ranking', range(1, len(monthly_sales_df)+1))

#formatted columns
with pd.ExcelWriter('data/MonthlySales.xlsx', engine='xlsxwriter') as writer:
    monthly_sales_df.to_excel(writer, index=False, sheet_name='Sales')    
    workbook  = writer.book
    worksheet = writer.sheets['Sales']
    
    number_format = workbook.add_format({'num_format': '#,##0'})
    worksheet.set_column('C:C', None, number_format)
    worksheet.autofit()    
   
    # Insert chart in worksheet
    worksheet.insert_image('E2', chartPathMonthlySales, {'x_scale': 1.00, 'y_scale': 1.00})





# Calculate total quantity sold per product
top_products = retailData.groupby('Description')['Quantity'].sum().sort_values(ascending=False).head(10)

# Plot top products
fig, ax = plt.subplots(figsize=(12,6))
top_products.plot(kind='bar', ax=ax)

# Rotate labels and set title
ax.set_xticklabels(ax.get_xticklabels(), rotation=45, ha='right')
ax.set_ylabel('Total Quantity Sold')
ax.set_title('Top 10 Products by Quantity Sold')

# Add horizontal grid lines aligned with y-ticks
ax.yaxis.grid(True, linestyle='--', linewidth=1, color='gray')

plt.tight_layout()

chartPathTotalQuantity = 'data\TopTenProductsSold_chart.png'
plt.savefig(chartPathTotalQuantity)
plt.close(fig)

#plt.show()
top_products_df = top_products.reset_index()
top_products_df.columns = ['Description', 'Quantity']
# Add Ranking column starting at 1
top_products_df.insert(0, 'Ranking', range(1, len(top_products_df)+1))

#formatted columns
with pd.ExcelWriter('data/TopTenProductsSold.xlsx', engine='xlsxwriter') as writer:
    top_products_df.to_excel(writer, index=False, sheet_name='Top Ten')    
    workbook  = writer.book
    worksheet = writer.sheets['Top Ten']
    
    number_format = workbook.add_format({'num_format': '#,##0'})
    worksheet.set_column('C:C', None, number_format)
    worksheet.autofit()    
   
    # Insert chart in worksheet
    worksheet.insert_image('E2', chartPathTotalQuantity, {'x_scale': 1.00, 'y_scale': 1.00})



# Calculate total quantity sold per customers
top_customers = retailData.groupby('Customer ID')['Quantity'].sum().nlargest(10)
top_customers.index = top_customers.index.astype(int).astype(str)
# Plot top customers
fig, ax = plt.subplots(figsize=(12,6))
top_customers.plot(kind='bar', ax=ax)

# Rotate labels and set title
ax.set_xticklabels(ax.get_xticklabels(), rotation=45, ha='right')
ax.set_ylabel('Total Quantity Sold')
ax.set_title('Top 10 Customers by Quantity Sold')

# Add horizontal grid lines aligned with y-ticks
ax.yaxis.grid(True, linestyle='--', linewidth=1, color='gray')

plt.tight_layout()
#plt.show()
chartPathTopTen = 'data\TopTenCustomers_chart.png'
plt.savefig(chartPathTopTen)
plt.close(fig)

top_customers_df = top_customers.reset_index()
top_customers_df.columns = ['Customer ID', 'Total Quantity Sold']
# Add Ranking column starting at 1
top_customers_df.insert(0, 'Ranking', range(1, len(top_customers_df)+1))

#Not formatted
#top_customers_df.to_excel('data/TopTenCustomers.xlsx', index=False)

#formatted columns
with pd.ExcelWriter('data/TopTenCustomers.xlsx', engine='xlsxwriter') as writer:
    top_customers_df.to_excel(writer, index=False, sheet_name='Top Customers')    
    workbook  = writer.book
    worksheet = writer.sheets['Top Customers']
    
    number_format = workbook.add_format({'num_format': '#,##0'})
    worksheet.set_column('C:C', None, number_format)
    worksheet.autofit()    
   
    # Insert chart in worksheet
    worksheet.insert_image('E2', chartPathTopTen, {'x_scale': 0.75, 'y_scale': 0.75})