# Symbol-wise-Turnover-Analysis-in-NEPSE-
# I have scrapped the complete floor sheet of companies traded in NEPSE in the last 30 days and analyzed the datewise turnover of each script symbol. 

import pandas as pd
import os

# Directory containing the input CSV files
input_directory = r"C:\Users\croje\OneDrive\0000.Daily Trade Data"

# Directory for the output Excel file
output_directory = r"C:\Users\croje\OneDrive\0000.Daily Trade Data\7 days Trading Analysis"
os.makedirs(output_directory, exist_ok=True)  # Create the directory if it doesn't exist

# List of CSV files to process
csv_files = [
    "04_01_2025.csv",
    "04_02_2025.csv",
    "04_03_2025.csv",
    "04_07_2025.csv",
    "04_08_2025.csv",
    "04_09_2025.csv",
    "04_10_2025.csv",
    "04_13_2025.csv",
    "04_15_2025.csv",
    "04_16_2025.csv",
    "04_17_2025.csv",
    "04_20_2025.csv",
    "04_21_2025.csv",
    "04_22_2025.csv",
    "04_23_2025.csv",
    "04_24_2025.csv",
    "04_27_2025.csv",
    "04_28_2025.csv",
    "04_29_2025.csv",
    "04_30_2025.csv",
    "05_04_2025.csv",
    "05_05_2025.csv",
    "05_06_2025.csv",
    "05_07_2025.csv",
    "05_08_2025.csv",
    "05_11_2025.csv",
    "05_13_2025.csv",
    "05_14_2025.csv",
    "05_18_2025.csv",
    "05_19_2025.csv",
    "05_20_2025.csv",
    "05_25_2025.csv",
    "06_09_2025.csv",
    "06_10_2025.csv",
    "06_11_2025.csv",
    "06_12_2025.csv",
    "06_15_2025.csv",
    "06_16_2025.csv",
    "06_17_2025.csv",
    "06_18_2025.csv",

]

# DataFrame to hold consolidated data
consolidated_df = pd.DataFrame()

# Process each CSV file
for file_name in csv_files:
    file_path = os.path.join(input_directory, file_name)
    if os.path.exists(file_path):
        df = pd.read_csv(file_path, usecols=["Symbol", "Amount", "Quantity"])  # Explicitly read necessary columns
        df["CSV File"] = file_name  # Add a column for the file name (date)
        consolidated_df = pd.concat([consolidated_df, df], ignore_index=True)
    else:
        print(f"File not found: {file_path}")

# Ensure column names are consistent and relevant columns exist
if "Symbol" in consolidated_df.columns and "Amount" in consolidated_df.columns and "Quantity" in consolidated_df.columns:
    # Pivot table to arrange data by Symbol and Date (CSV File)
    pivot_amount = consolidated_df.pivot_table(index="Symbol", columns="CSV File", values="Amount", aggfunc="sum", fill_value=0)
    pivot_quantity = consolidated_df.pivot_table(index="Symbol", columns="CSV File", values="Quantity", aggfunc="sum", fill_value=0)

    # Sort columns by date order
    pivot_amount = pivot_amount.reindex(columns=csv_files, fill_value=0)
    pivot_quantity = pivot_quantity.reindex(columns=csv_files, fill_value=0)

    # Combine the data into a single DataFrame
    final_df = pivot_amount.copy()
    final_df.columns = [f"Amount ({col})" for col in pivot_amount.columns]
    quantity_columns = [f"Quantity ({col})" for col in pivot_quantity.columns]

    for q_col, q_data in zip(quantity_columns, pivot_quantity.T.values):
        final_df[q_col] = q_data

    # Determine the latest date from the file names
    latest_date = max(csv_files).replace(".csv", "")

    # Save the result to a new Excel file
    output_file = os.path.join(output_directory, f"7_days_trading_analysis_{latest_date}.xlsx")
    final_df.to_excel(output_file, index_label="Symbol")

    print(f"Detailed trading data saved to: {output_file}")
else:
    print("Required columns (Symbol, Amount, Quantity) not found in the consolidated data.")
