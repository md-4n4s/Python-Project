import pandas as pd
from pathlib import Path

# ====================
# Data Cleaning
# ====================

df = pd.read_excel("products.xlsx")

df = df.drop_duplicates(subset="name", keep="first")

# Remove rows where any data is missing
df = df.dropna()

# Clean and standardize name
df["name"] = df["name"].str.strip().str.title()

df.to_excel("clean_products.xlsx", index=False)

# ====================
# Manage multiple files
# ====================

folder = Path("sample_sales_data")

excel_files = folder.glob("*.xlsx")

all_data = []
for excel_file in excel_files:
    df = pd.read_excel(excel_file)
    all_data.append(df)

# Combine all data
combined_df = pd.concat(all_data, ignore_index=True)

# Drop empty rows
combined_df = combined_df.dropna()

# Calculate sales by Category
summary = (
    combined_df.groupby("Category", as_index=False)["Total Sales ($)"].sum().sort_values("Total Sales ($)", ascending=False)
)

# Save everything to one Excel workbook
with pd.ExcelWriter("Sales_Report.xlsx", engine="openpyxl") as writer:
    combined_df.to_excel(writer, sheet_name = "All Data", index=False)
    summary.to_excel(writer, sheet_name = "Summary", index=False)