import pandas as pd
import os

def inspect_excel(filepath):
    print(f"Inspecting: {filepath}")
    try:
        xls = pd.ExcelFile(filepath)
        print("Sheet names:", xls.sheet_names)
        for sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet, nrows=5)
            print(f"\n--- Sheet: {sheet} ---")
            print("Columns:", list(df.columns))
            print("First 2 rows:")
            print(df.head(2).to_string())
    except Exception as e:
        print(f"Error reading file: {e}")

input_dir = r"c:\Users\mohamed.hassan\OneDrive - Next Structural Integrity\3_Projects\Henry Song's files - HS_Feeder Project\20. Automation\ratchet\data\input"
files = ["temp_test_model_1.xlsx"]

if files:
    inspect_excel(os.path.join(input_dir, files[0]))
else:
    print("No Excel files found in input directory.")
