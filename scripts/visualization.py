import pandas as pd
import matplotlib.pyplot as plt  # add this

path = r"C:\Users\Troy\OneDrive\Desktop\python\datasets\Online Retail\data\out.csv"
retailData = pd.read_csv(path)
retailData['InvoiceDate'] = pd.to_datetime(retailData['InvoiceDate'])

sales = retailData[retailData['Quantity']>0]

monthly_sales = sales.resample('ME', on='InvoiceDate')['Quantity'].sum()

ax = monthly_sales.plot(kind='bar', figsize=(12,6))

ax.set_xticklabels([d.strftime('%b %Y') for d in monthly_sales.index], rotation=45, ha='right')

plt.ylabel('Quantity Sold')
plt.title('Monthly Sales')
plt.tight_layout()







# # Plot cumulative monthly sales
# ax = monthly_sales.plot(figsize=(12,6))
# ax.set_title('Cumulative Monthly Sales')
# ax.set_ylabel('Quantity Sold')
# ax.set_xlabel('Month')
# ax.grid(True)
# ax.ticklabel_format(style='plain', axis='y')  # disables scientific notation

# # ts = retailData.set_index('InvoiceDate')['Quantity']
# # ts_daily = ts.resample('ME').sum()
# # ts_daily.cumsum().plot()

# # print(retailData[retailData['Quantity']<0].head(1000))
# # print(retailData['Quantity'].min())
# print(monthly_sales.max())

# # print(retailData.info())
plt.show()