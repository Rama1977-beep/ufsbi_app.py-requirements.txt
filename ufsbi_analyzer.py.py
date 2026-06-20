import pandas as pd
import os

print("=" * 50)
print("UFSBI ANALYZER")
print("=" * 50)

# Check files exist

healthy_file = "healthy.xlsx"
fault_file = "fault.xlsx"

if not os.path.exists(healthy_file):
    print(f"ERROR : {healthy_file} not found")
    exit()

if not os.path.exists(fault_file):
    print(f"ERROR : {fault_file} not found")
    exit()

try:

    healthy_df = pd.read_excel(healthy_file)
    fault_df = pd.read_excel(fault_file)

    print("\nHealthy File Loaded Successfully")
    print("Rows :", len(healthy_df))
    print("Columns :", len(healthy_df.columns))

    print("\nFault File Loaded Successfully")
    print("Rows :", len(fault_df))
    print("Columns :", len(fault_df.columns))

    print("\nHealthy File Columns:")
    print(list(healthy_df.columns))

    print("\nFault File Columns:")
    print(list(fault_df.columns))

except Exception as e:
    print("\nERROR READING FILE")
    print(str(e))