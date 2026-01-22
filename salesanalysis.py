import pandas as pd
import numpy as np


data = {
    'Order ID': range(1, 21),
    'Product Name': [
        'Laptop', 'Phone', 'Tablet', 'Headphones', 'Charger',
        'Mouse', 'Keyboard', 'Monitor', 'Camera', 'Smartwatch',
        'Laptop', 'Phone', 'Tablet', 'Headphones', 'Charger',
        'Mouse', 'Keyboard', 'Monitor', 'Camera', 'Smartwatch'
    ],
    'Region': np.random.choice(['North', 'South', 'East', 'West'], 20),
    'Sales': np.random.randint(5000, 25000, 20),
    'Profit': np.random.randint(500, 8000, 20),
    'Order Date': pd.date_range(start='2024-01-01', periods=20, freq='15D')
}

df = pd.DataFrame(data)


print("----- First 5 Rows -----")
print(df.head(), "\n")

print("----- Dataset Info -----")
print(df.info(), "\n")

total_sales = df['Sales'].sum()
total_profit = df['Profit'].sum()
total_orders = len(df)

print("----- Overall Summary -----")
print(f"Total Sales: ₹{total_sales:,.2f}")
print(f"Total Profit: ₹{total_profit:,.2f}")
print(f"Total Orders: {total_orders}\n")


top_products = df.groupby('Product Name')['Sales'].sum().sort_values(ascending=False).head(3)
print("----- Top 3 Products by Sales -----")
print(top_products, "\n")


sales_by_region = df.groupby('Region')['Sales'].sum().sort_values(ascending=False)
print("----- Sales by Region -----")
print(sales_by_region, "\n")


df['Month'] = df['Order Date'].dt.to_period('M')
monthly_sales = df.groupby('Month')['Sales'].sum()
print("----- Monthly Sales Trend -----")
print(monthly_sales, "\n")


growth_rate = monthly_sales.pct_change().fillna(0) * 100
print("----- Monthly Growth Rate (%) -----")
print(growth_rate, "\n")


summary = pd.DataFrame({
    'Total Sales': [total_sales],
    'Total Profit': [total_profit],
    'Total Orders': [total_orders]
})

with pd.ExcelWriter("Sales_Analysis_Report.xlsx") as writer:
    df.to_excel(writer, sheet_name='Raw_Data', index=False)
    summary.to_excel(writer, sheet_name='Summary', index=False)
    top_products.to_excel(writer, sheet_name='Top_Products')
    sales_by_region.to_excel(writer, sheet_name='Sales_by_Region')
    monthly_sales.to_excel(writer, sheet_name='Monthly_Sales')
    growth_rate.to_excel(writer, sheet_name='Growth_Rate')


