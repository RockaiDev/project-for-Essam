
import pandas as pd

def check_structure():
    xls = pd.ExcelFile('BulkContractorDetailedInvoice.xlsx')
    sheet = xls.sheet_names[0]
    df = pd.read_excel(xls, sheet_name=sheet)
    
    # Print first 25 rows content to spot Contract
    print("--- First 25 Rows Dump ---")
    for i, row in df.head(25).iterrows():
        print(f"Row {i}:")
        for idx, x in enumerate(row.values):
             if pd.notna(x):
                 print(f"  [{idx}] {str(x).strip()}")
    print("--------------------------")

    # Find header
    header_row_idx = None
    for i, row in df.iterrows():
        vals = [str(x).lower().strip() for x in row.values if pd.notna(x)]
        if 'description' in vals and ('qty' in vals or 'unit' in vals):
            header_row_idx = i
            break
            
    if header_row_idx is None:
        print("Header not found")
        return

    print(f"Header found at row {header_row_idx}")
    header_row = df.iloc[header_row_idx]
    
    # Print all non-null columns in header
    print("Columns found in Excel:")
    mapped_cols = []
    for idx, val in enumerate(header_row):
        if pd.notna(val):
            print(f" - Col {idx}: {val}")
            mapped_cols.append(idx)
            
    # Now check what our script maps
    # Our script maps: Description, Qty, Total, Discount, Debit, Date, Unit
    
    print("\nSample Data Row (Header + 1):")
    data_row = df.iloc[header_row_idx + 1]
    for idx in mapped_cols:
        print(f" - Col {idx} ({header_row[idx]}): {data_row[idx]}")

    print("\n--- Searching for Grand Total Row ---")
    for i in range(header_row_idx + 1, len(df)):
        row = df.iloc[i]
        row_vals = [str(x).strip().lower() for x in row.values if pd.notna(x)]
        if 'grand total' in ' '.join(row_vals):
            print(f"Found Grand Total at Row {i}:")
            for idx, val in enumerate(row.values):
                if pd.notna(val):
                    print(f"  Col {idx}: {val}")
            break
    print("-" * 50)

if __name__ == "__main__":
    check_structure()
