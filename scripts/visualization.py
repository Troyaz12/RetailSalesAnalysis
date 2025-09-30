import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter

#---- Helper Methods ----
def create_bar_chart(series, title, ylabel, chart_path, currency=False, date=False):
    fig, ax = plt.subplots(figsize=(12,6))
    series.plot(kind='bar', ax=ax)

    if date:
        ax.set_xticklabels([d.strftime('%b %Y') for d in series.index], rotation=45, ha='right')
    else:
        ax.set_xticklabels(ax.get_xticklabels(), rotation=45, ha='right')

    ax.set_ylabel(ylabel)
    ax.set_title(title)
    ax.yaxis.grid(True, linestyle='--', linewidth=1, color='gray')

    if currency:
        ax.yaxis.set_major_formatter(FuncFormatter(lambda x, _: f'${x:,.0f}'))       

    plt.tight_layout()
    plt.savefig(chart_path)
    plt.close(fig)

def export_to_excel(writer, df, sheet_name, col_format, chart_path=None, chart_scale=(1.0,1.0)):
    df.to_excel(writer, index=False, sheet_name=sheet_name)
    worksheet = writer.sheets[sheet_name]
    worksheet.set_column(1, len(df.columns), col_format)
    worksheet.autofit()
    if chart_path:
        worksheet.insert_image('E2', chart_path, {'x_scale': chart_scale[0], 'y_scale': chart_scale[1]})

#---- Main Script ----
path = r"C:\Users\Troy\OneDrive\Desktop\python\datasets\Online Retail\data\out.csv"
retailData = pd.read_csv(path)
retailData['InvoiceDate'] = pd.to_datetime(retailData['InvoiceDate'])

sales = retailData[retailData['Quantity']>0]

monthly_sales = sales.resample('ME', on='InvoiceDate')['Quantity'].sum()
chartPathMonthlySales = 'data\Charts\MonthlySales_chart.png'

create_bar_chart(monthly_sales, "Monthly Sales (Quantity Sold)", "Quantity Sold", chartPathMonthlySales, currency = False, date = True)

monthly_sales_df = monthly_sales.reset_index()
monthly_sales_df['InvoiceDate'] = monthly_sales_df['InvoiceDate'].dt.strftime('%Y-%m-%d')

monthly_sales_df.columns = ['Invoice Date', 'Quantity']

#---- Calculate total quantity sold per product ----
top_products = retailData.groupby('Description')['Quantity'].sum().sort_values(ascending=False).head(10)
chartPathTotalQuantity = 'data\Charts\TopTenProductsSold_chart.png'

create_bar_chart(top_products, "Top 10 Products by Quantity Sold", "Total Quantity Sold", chartPathTotalQuantity, currency = False, date = False)

top_products_df = top_products.reset_index()
top_products_df.columns = ['Description', 'Quantity']
# Add Ranking column starting at 1
top_products_df.insert(0, 'Ranking', range(1, len(top_products_df)+1))

#---- Calculate total quantity sold per customers ----
top_customers = retailData.groupby('Customer ID')['Quantity'].sum().nlargest(10)
top_customers.index = top_customers.index.astype(int).astype(str)
chartPathTopTenCustomers = 'data\Charts\TopTenCustomers_chart.png'

create_bar_chart(top_customers, "Top 10 Customers by Quantity Sold", "Total Quantity Sold", chartPathTopTenCustomers, currency = False, date = False)

top_customers_df = top_customers.reset_index()
top_customers_df.columns = ['Customer ID', 'Total Quantity Sold']
# Add Ranking column starting at 1
top_customers_df.insert(0, 'Ranking', range(1, len(top_customers_df)+1))

#---- Monthly Revenue ----
monthly_revenue = retailData.resample('ME', on='InvoiceDate')['Revenue'].sum()
chartPathMonthlyRevenue = 'data\Charts\MonthlySalesRevenue_chart.png'

create_bar_chart(monthly_revenue, "Monthly Sales (Revenue)", "Revenue ($)", chartPathMonthlyRevenue, currency = True, date = True)

monthly_Revenue_df = monthly_revenue.reset_index()
monthly_Revenue_df['InvoiceDate'] = monthly_Revenue_df['InvoiceDate'].dt.strftime('%Y-%m-%d')

monthly_Revenue_df.columns = ['Invoice Date', 'Revenue']


#---- Calculate total revenue per product ----
top_products_Revenue = retailData.groupby('Description')['Revenue'].sum().sort_values(ascending=False).head(10)
chartPathTotalRevenue = 'data\Charts\TopTenProductsByRevenue_chart.png'

create_bar_chart(top_products_Revenue, "Top 10 Products by Revenue ($)", "Total Revenue", chartPathTotalRevenue, currency = True, date = False)

top_products_Revenue_df = top_products_Revenue.reset_index()
top_products_Revenue_df.columns = ['Description', 'Revenue']
# Add Ranking column starting at 1
top_products_Revenue_df.insert(0, 'Ranking', range(1, len(top_products_Revenue_df)+1))

#---- Calculate top ten Revenue per customers ----
top_customers_ByRevenue = retailData.groupby('Customer ID')['Revenue'].sum().nlargest(10)
top_customers_ByRevenue.index = top_customers_ByRevenue.index.astype(int).astype(str)
chartPathTopCustomerRevenue = 'data\Charts\TopTenCustomersByRevenue_chart.png'

create_bar_chart(top_customers_ByRevenue, "Top 10 Customers by Revenue ($)", "Total Revenue", chartPathTopCustomerRevenue, currency = True, date = False)

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
    worksheet.insert_image('E2', chartPathTopTenCustomers, {'x_scale': 0.75, 'y_scale': 0.75})

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



# quantity_format = workbook.add_format({'num_format': '#,##0'})
# currency_format = workbook.add_format({'num_format': '$#,##0'})

# export_to_excel(writer, top_customers_df, "Top Customers", quantity_format, chartPathTopTen, (0.75,0.75))
# export_to_excel(writer, monthly_sales_df, "Sales Trends", quantity_format, chartPathMonthlySales)
# export_to_excel(writer, top_products_df, "Top Ten Products", quantity_format, chartPathTotalQuantity)
# export_to_excel(writer, monthly_Revenue_df, "Sales Revenue Trend", currency_format, chartPathMonthlyRevenue)
# export_to_excel(writer, top_products_Revenue_df, "Top Ten Products By Revenue", currency_format, chartPathTotalRevenue)
# export_to_excel(writer, top_customers_ByRevenue_df, "Top Ten Customers By Revenue", currency_format, chartPathTopCustomerRevenue)
