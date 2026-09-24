import pandas as pd

data = {
    "Transaction_ID": [101, 102, 103, 104, 105, 106, 107, 108],
    "Region": ["North", "South", "East", "West", "North", "South", "East", "West"],
    "Sales_Rep": ["Amit", "Priya", "Rahul", "Neha", "Amit", "Priya", "Rahul", "Neha"],
    "Category": ["Electronics", "Furniture", "Electronics", "Furniture", "Furniture", "Electronics", "Furniture", "Electronics"],
    "Units_Sold": [12, 5, 25, 8, 15, 30, 4, 18],
    "Unit_Price": [150.0, 450.0, 80.0, 300.0, 200.0, 120.0, 500.0, 95.0]
}

df = pd.DataFrame(data)
df.to_excel("Raw_Sales_Data.xlsx", sheet_name="Sales_Data", index=False)
print("Created 'Raw_Sales_Data.xlsx'")