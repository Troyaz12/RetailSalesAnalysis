"""
Dataset Column Descriptions:

StockCode : Item identifier
Quantity : Number of items purchased
InvoiceDate : Date the item was purchased
Price : Price per item
Customer ID : Unique customer identifier
Country : Country of the customer

"""
import pandas as pd

path = r"C:\Users\Troy\OneDrive\Desktop\python\datasets\Online Retail\data\online_retail_II.xlsx"
retailData = pd.read_excel(path, sheet_name=None)

cleaned_dfs = []

for sheet_name, df in retailData.items():
 if 'Cusomer ID' in df.columns:
  df['Customer ID'] = df['Customer ID'].fillna('Unknown')
 if 'Description' in df.columns:
  df['Description'] = df['Description'].fillna('No Description Available')

 df.drop(columns=[col for col in ['Invoice'] if col in df.columns], inplace=True)

 df.drop_duplicates(inplace=True)
 cleaned_dfs.append(df)

 combined_df = pd.concat(cleaned_dfs, ignore_index=True)

 # Drop rows with missing Quantity or Price after combining
 combined_df.dropna(subset=['Quantity', 'Price'], inplace=True)

 combined_df["InvoiceYear"] = combined_df["InvoiceDate"].dt.year
 combined_df["InvoiceMonth"] = combined_df["InvoiceDate"].dt.month
 combined_df["InvoiceWeekday"] = combined_df["InvoiceDate"].dt.day_name()
 combined_df["InvoiceHour"] = combined_df["InvoiceDate"].dt.hour

 combined_df['Quantity'] = pd.to_numeric(combined_df['Quantity'], errors='coerce')
 combined_df['Price'] = pd.to_numeric(combined_df['Price'], errors='coerce')
 combined_df.dropna(subset=['Quantity','Price'], inplace=True)

 combined_df['Revenue'] = combined_df['Quantity'] * combined_df['Price']
 


 combined_df.to_csv('data/out.csv', index=False)

print(combined_df.head())
print(combined_df.columns)
print(combined_df.shape)
print(combined_df.info())
print(combined_df.describe())