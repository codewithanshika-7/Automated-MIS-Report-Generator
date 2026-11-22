import pandas as pd

# Sample raw sales data with intentional duplicates and missing values
sales_data = {
    "Date": ["2026-11-20", "2026-11-20", "2026-11-21", "2026-11-21", "2026-11-22", "2026-11-22", "2026-11-22", "2026-11-22"],
    "Region": ["East", "East", "North", "West", "East", "South", "East", "North"],
    "Product": ["NeuroCare 50", "NeuroCare 50", "CardioPlus", "NeuroCare 50", "DermaSafe", "CardioPlus", None, "OrthoRelief"],
    "Units_Sold": [120, 120, 85, 95, 150, None, 60, 110],
    "Unit_Price": [250, 250, 400, 250, 180, 400, 250, 300]
}

# Sample inventory stock data
inventory_data = {
    "Product": ["NeuroCare 50", "CardioPlus", "DermaSafe", "OrthoRelief"],
    "Current_Stock": [45, 320, 30, 500],
    "Reorder_Level": [100, 150, 80, 200]
}

with pd.ExcelWriter("raw_pharma_data.xlsx", engine="openpyxl") as writer:
    pd.DataFrame(sales_data).to_excel(writer, sheet_name="Raw_Sales", index=False)
    pd.DataFrame(inventory_data).to_excel(writer, sheet_name="Inventory", index=False)

print("Sample raw data created: raw_pharma_data.xlsx")