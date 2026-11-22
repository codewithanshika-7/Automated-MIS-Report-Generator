import pandas as pd

# 1. Read raw multi-sheet Excel data
input_file = "raw_pharma_data.xlsx"
sales_df = pd.read_excel(input_file, sheet_name="Raw_Sales")
inventory_df = pd.read_excel(input_file, sheet_name="Inventory")

# 2. Data Cleaning & Validation
# Drop exact duplicate rows and rows missing critical fields
sales_clean = sales_df.drop_duplicates().dropna(subset=["Product", "Units_Sold"]).copy()
sales_clean["Units_Sold"] = sales_clean["Units_Sold"].astype(int)

# Compute transaction revenue
sales_clean["Total_Revenue"] = sales_clean["Units_Sold"] * sales_clean["Unit_Price"]

# 3. MIS Aggregations & Business Logic
# Region-wise sales summary
region_summary = (
    sales_clean.groupby("Region", as_index=False)
    .agg(Total_Units=("Units_Sold", "sum"), Total_Revenue=("Total_Revenue", "sum"))
    .sort_values(by="Total_Revenue", ascending=False)
)

# Product-wise performance summary
product_summary = (
    sales_clean.groupby("Product", as_index=False)
    .agg(Total_Units=("Units_Sold", "sum"), Total_Revenue=("Total_Revenue", "sum"))
    .sort_values(by="Total_Revenue", ascending=False)
)

# Low-stock reorder alerts
inventory_df["Stock_Status"] = inventory_df.apply(
    lambda row: "LOW STOCK - REORDER" if row["Current_Stock"] < row["Reorder_Level"] else "OK",
    axis=1
)

# 4. Export formatted multi-tab MIS Excel report
output_file = "Automated_MIS_Report.xlsx"
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    region_summary.to_excel(writer, sheet_name="Region_Summary", index=False)
    product_summary.to_excel(writer, sheet_name="Product_Performance", index=False)
    inventory_df.to_excel(writer, sheet_name="Inventory_Alerts", index=False)
    sales_clean.to_excel(writer, sheet_name="Cleaned_Sales_Data", index=False)

print(f"Automated MIS Report generated successfully: {output_file}")