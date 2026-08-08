# Sales Report Generator

A Python script that reads multiple Excel sales-data files from a folder, combines them into a single dataset, calculates totals and category summaries, and produces a formatted Excel report (`Sales_Report.xlsx`) complete with styled headers, borders, auto-filters, and a bar chart.

## What It Does

1. Scans a folder (`sample_sales_data/`) for all `.xlsx` files.
2. Validates that each file contains the required columns.
3. Combines all files into a single dataset, tagging each row with its source file.
4. Removes duplicate items (by `Item ID`) and rows with missing values.
5. Converts numeric columns and calculates a `Total Cost $` column.
6. Builds a summary of total `Quantity Sold` per `Category`.
7. Writes everything to `Sales_Report.xlsx` with two sheets: `All Data` and `Summary`.
8. Applies formatting (colored headers, borders, column auto-width, frozen header row, auto-filters).
9. Adds a bar chart to the `Summary` sheet showing quantity sold by category.

## Requirements

- Python 3.8+
- Packages:
  ```bash
  pip install pandas openpyxl
  ```

## Project Structure

```
project/
├── main.py          # Main script (your code)
├── README.md                 # This file
└── sample_sales_data/        # Folder containing input Excel files
    ├── file1.xlsx
    ├── file2.xlsx
    └── ...
```

## Required Input Format

Every `.xlsx` file inside `sample_sales_data/` must contain these columns (exact names):

| Column            | Type   | Notes                          |
|-------------------|--------|--------------------------------|
| Item ID           | text   | Used to drop duplicate items   |
| Item Name         | text   |                                |
| Category          | text   | Used for the summary/chart     |
| Region            | text   |                                |
| Quantity Sold     | number | Converted to numeric           |
| Unit Price ($)    | number | Converted to numeric           |

If a file is missing any of these columns, the script raises a `ValueError` and stops.

## Usage

1. Place all raw sales Excel files inside a folder named `sample_sales_data` in the same directory as the script.
2. Run the script:
   ```bash
   python main.py
   ```
3. Find the generated report as `Sales_Report.xlsx` in the same directory.

## Output

`Sales_Report.xlsx` contains:

- **All Data** — the combined, cleaned dataset from every input file, with an added `Source File` and `Total Cost $` column.
- **Summary** — total `Quantity Sold` per `Category`, sorted in descending order, plus a bar chart visualizing the totals.

Both sheets have:
- Bold white text on a dark blue header row
- Thin borders on all cells
- Auto-sized columns
- A frozen header row
- Auto-filter enabled on all columns
