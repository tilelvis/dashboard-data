import pandas as pd
from faker import Faker
import random

# Initialize Faker
fake = Faker()

# Number of rows to generate
num_rows = 10000

# Lists to store the generated data
order_ids = []
products = []
categories = []
unit_prices = []
quantities = []
order_dates = []

# Product categories and their corresponding products
product_data = {
    'Electronics': ['Laptop', 'Monitor', 'Keyboard', 'Mouse', 'Headphones'],
    'Clothing': ['T-Shirt', 'Jeans', 'Jacket', 'Socks', 'Hat'],
    'Home Goods': ['Chair', 'Table', 'Lamp', 'Rug', 'Vase'],
    'Books': ['Fiction', 'Non-Fiction', 'Sci-Fi', 'Mystery', 'Biography']
}
categories_list = list(product_data.keys())

# Generate the data
for i in range(1, num_rows + 1):
    order_ids.append(i)

    # Choose a random category and then a random product from that category
    category = random.choice(categories_list)
    product = random.choice(product_data[category])
    categories.append(category)
    products.append(product)

    unit_prices.append(round(random.uniform(10, 500), 2))
    quantities.append(random.randint(1, 10))
    order_dates.append(fake.date_between(start_date='-1y', end_date='today'))

# Create the DataFrame
data = {
    'OrderID': order_ids,
    'Product': products,
    'Category': categories,
    'UnitPrice': unit_prices,
    'Quantity': quantities,
    'OrderDate': order_dates
}
df = pd.DataFrame(data)

# Calculate total sales
df['TotalSales'] = df['UnitPrice'] * df['Quantity']

# Create a pivot table
pivot_table = df.pivot_table(
    values='TotalSales',
    index='Category',
    columns='Product',
    aggfunc='sum',
    fill_value=0
)

# Create a new Excel file and add the data and pivot table to it
writer = pd.ExcelWriter('excel-dashboard/dashboard.xlsx', engine='xlsxwriter')
df.to_excel(writer, sheet_name='Sales Data', index=False)
pivot_table.to_excel(writer, sheet_name='Pivot Table')
writer.close()
