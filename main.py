import pandas as pd

# Read the Excel File
df = pd.read_excel("products.xlsx")

# Remove duplicate rows based on name
df = df.drop_duplicates(subset="name", keep="first")

# Remove rows where any data is missing
df = df.dropna(subset=["name", "review", "price"])

# Clean and standardize name
df["name"] = df["name"].str.strip().str.title()

# Save the clean Excel file
df.to_excel("products.xlsx", index=False)