import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter

path = r"C:\Users\Troy\OneDrive\Desktop\python\datasets\Online Retail\data\out.csv"
retailData = pd.read_csv(path)
retailData['InvoiceDate'] = pd.to_datetime(retailData['InvoiceDate'])

sales = retailData[retailData['Quantity']>0]

monthly_sales = sales.resample('ME', on='InvoiceDate')['Quantity'].sum()

ax = monthly_sales.plot(kind='bar', figsize=(12,6))
fig = ax.get_figure()
ax.set_xticklabels([d.strftime('%b %Y') for d in monthly_sales.index], rotation=45, ha='right')

plt.ylabel('Quantity Sold')
plt.title('Monthly Sales (Quantity)')
plt.tight_layout()

chartPathMonthlySales = 'data\Charts\MonthlySales_chart.png'
plt.savefig(chartPathMonthlySales)
plt.close(fig)
#plt.show()

monthly_sales_df = monthly_sales.reset_index()
monthly_sales_df['InvoiceDate'] = monthly_sales_df['InvoiceDate'].dt.strftime('%Y-%m-%d')

monthly_sales_df.columns = ['Invoice Date', 'Quantity']

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

chartPathTotalQuantity = 'data\Charts\TopTenProductsSold_chart.png'
plt.savefig(chartPathTotalQuantity)
plt.close(fig)

#plt.show()
top_products_df = top_products.reset_index()
top_products_df.columns = ['Description', 'Quantity']
# Add Ranking column starting at 1
top_products_df.insert(0, 'Ranking', range(1, len(top_products_df)+1))

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
chartPathTopTen = 'data\Charts\TopTenCustomers_chart.png'
plt.savefig(chartPathTopTen)
plt.close(fig)

top_customers_df = top_customers.reset_index()
top_customers_df.columns = ['Customer ID', 'Total Quantity Sold']
# Add Ranking column starting at 1
top_customers_df.insert(0, 'Ranking', range(1, len(top_customers_df)+1))

#### Monthly Revenue
monthly_revenue = retailData.resample('ME', on='InvoiceDate')['Revenue'].sum()

ax = monthly_revenue.plot(kind='bar', figsize=(12,6))
fig = ax.get_figure()
ax.set_xticklabels([d.strftime('%b %Y') for d in monthly_revenue.index], rotation=45, ha='right')

ax.yaxis.set_major_formatter(FuncFormatter(lambda x, _: f'${x:,.0f}'))

plt.ylabel('Revenue ($)')
plt.title('Monthly Sales (Revenue)')
plt.tight_layout()

chartPathMonthlyRevenue = 'data\Charts\MonthlySalesRevenue_chart.png'
plt.savefig(chartPathMonthlyRevenue)
plt.close(fig)
#plt.show()

monthly_Revenue_df = monthly_revenue.reset_index()
monthly_Revenue_df['InvoiceDate'] = monthly_Revenue_df['InvoiceDate'].dt.strftime('%Y-%m-%d')

monthly_Revenue_df.columns = ['Invoice Date', 'Revenue']


# Calculate total revenue per product
top_products_Revenue = retailData.groupby('Description')['Revenue'].sum().sort_values(ascending=False).head(10)

# Plot top products
fig, ax = plt.subplots(figsize=(12,6))
top_products_Revenue.plot(kind='bar', ax=ax)

# Rotate labels and set title
ax.set_xticklabels(ax.get_xticklabels(), rotation=45, ha='right')
ax.set_ylabel('Total Revenue')
ax.set_title('Top 10 Products by Revenue ($)')

# Add horizontal grid lines aligned with y-ticks
ax.yaxis.grid(True, linestyle='--', linewidth=1, color='gray')

plt.tight_layout()

chartPathTotalRevenue = 'data\Charts\TopTenProductsByRevenue_chart.png'
plt.savefig(chartPathTotalRevenue)
plt.close(fig)

#plt.show()
top_products_Revenue_df = top_products_Revenue.reset_index()
top_products_Revenue_df.columns = ['Description', 'Revenue']
# Add Ranking column starting at 1
top_products_Revenue_df.insert(0, 'Ranking', range(1, len(top_products_Revenue_df)+1))

# Calculate top ten Revenue per customers
top_customers_ByRevenue = retailData.groupby('Customer ID')['Revenue'].sum().nlargest(10)
top_customers_ByRevenue.index = top_customers_ByRevenue.index.astype(int).astype(str)
# Plot top customers
fig, ax = plt.subplots(figsize=(12,6))
top_customers_ByRevenue.plot(kind='bar', ax=ax)

# Rotate labels and set title
ax.set_xticklabels(ax.get_xticklabels(), rotation=45, ha='right')
ax.set_ylabel('Total Revenue')
ax.set_title('Top 10 Customers by Revenue ($)')

# Add horizontal grid lines aligned with y-ticks
ax.yaxis.grid(True, linestyle='--', linewidth=1, color='gray')

plt.tight_layout()
#plt.show()
chartPathTopCustomerRevenue = 'data\Charts\TopTenCustomersByRevenue_chart.png'
plt.savefig(chartPathTopCustomerRevenue)
plt.close(fig)

top_customers_ByRevenue_df = top_customers_ByRevenue.reset_index()
top_customers_ByRevenue_df.columns = ['Customer ID', 'Total Revenue']
# Add Ranking column starting at 1
top_customers_ByRevenue_df.insert(0, 'Ranking', range(1, len(top_customers_ByRevenue_df)+1))

#formatted columns
with pd.ExcelWriter('data/RetailAnalysis.xlsx', engine='xlsxwriter') as writer:
    top_customers_df.to_excel(writer, index=False, sheet_name='Top Customers By Quantity')    
    workbook  = writer.book
    worksheet = writer.sheets['Top Customers By Quantity']
    
    number_format = workbook.add_format({'num_format': '#,##0'})
    worksheet.set_column('C:C', None, number_format)
    worksheet.autofit()    
   
    # Insert chart in worksheet
    worksheet.insert_image('E2', chartPathTopTen, {'x_scale': 0.75, 'y_scale': 0.75})

    monthly_sales_df.to_excel(writer, index=False, sheet_name='Sales Trends By Quantity')
    worksheet = writer.sheets['Sales Trends By Quantity']
    
    number_format = workbook.add_format({'num_format': '#,##0'})
    worksheet.set_column('B:B', None, number_format)
    worksheet.autofit()    
   
    # Insert chart in worksheet
    worksheet.insert_image('E2', chartPathMonthlySales, {'x_scale': 1.00, 'y_scale': 1.00})

    top_products_df.to_excel(writer, index=False, sheet_name='Top Ten Products By Quantity')
    worksheet = writer.sheets['Top Ten Products By Quantity']
        
    number_format = workbook.add_format({'num_format': '#,##0'})
    worksheet.set_column('C:C', None, number_format)
    worksheet.autofit()    
    
    # Insert chart in worksheet
    worksheet.insert_image('E2', chartPathTotalQuantity, {'x_scale': 1.00, 'y_scale': 1.00})

    #Monthly Revenue
    monthly_Revenue_df.to_excel(writer, index=False, sheet_name='Sales Revenue Trend')
    worksheet = writer.sheets['Sales Revenue Trend']
    
    number_format = workbook.add_format({'num_format': '$#,##0'})
    worksheet.set_column('B:B', None, number_format)
    worksheet.autofit()    
   
    # Insert chart in worksheet
    worksheet.insert_image('E2', chartPathMonthlyRevenue, {'x_scale': 1.00, 'y_scale': 1.00})

    #Top ten products by Revenue
    top_products_Revenue_df.to_excel(writer, index=False, sheet_name='Top Ten Products By Revenue')
    worksheet = writer.sheets['Top Ten Products By Revenue']
        
    number_format = workbook.add_format({'num_format': '$#,##0'})
    worksheet.set_column('C:C', None, number_format)
    worksheet.autofit()    
    
    # Insert chart in worksheet
    worksheet.insert_image('E2', chartPathTotalRevenue, {'x_scale': 1.00, 'y_scale': 1.00})

    #Top ten Customers by Revenue
    top_customers_ByRevenue_df.to_excel(writer, index=False, sheet_name='Top Ten Customers By Revenue')
    worksheet = writer.sheets['Top Ten Customers By Revenue']
        
    number_format = workbook.add_format({'num_format': '$#,##0'})
    worksheet.set_column('C:C', None, number_format)
    worksheet.autofit()    
    
    # Insert chart in worksheet
    worksheet.insert_image('E2', chartPathTopCustomerRevenue, {'x_scale': 1.00, 'y_scale': 1.00})
