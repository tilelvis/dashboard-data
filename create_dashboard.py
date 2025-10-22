import pandas as pd

# Create a DataFrame with the sample data
data = {
    'Region': ['North', 'North', 'South', 'South', 'East', 'East', 'West', 'West'],
    'Product': ['Laptop', 'Monitor', 'Laptop', 'Monitor', 'Laptop', 'Monitor', 'Laptop', 'Monitor'],
    'Sales': [10000, 5000, 15000, 8000, 12000, 6000, 20000, 10000]
}
df = pd.DataFrame(data)

# Create a pivot table
pivot_table = df.pivot_table(
    values='Sales',
    index='Region',
    columns='Product',
    aggfunc='sum'
)

# Create a new Excel file and add the pivot table to it
writer = pd.ExcelWriter('excel-dashboard/dashboard.xlsx', engine='xlsxwriter')
pivot_table.to_excel(writer, sheet_name='Pivot Table')
writer.close()
