import os
import re
import glob
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

def generate_dirty_data():
    """Generates intentional messy mock datasets simulating production system bugs."""
    print("[+] Simulating real-world dirty corporate datasets...")
    os.makedirs("raw_data_inputs", exist_ok=True)
    
    # Dataset 1: User Transactions (Duplicates, Bad Casings, Extra Spaces)
    data_1 = {
        "Transaction_ID": [1001, 1002, 1002, 1003, 1004, np.nan, 1005],
        "Customer_Name": [" Alice Smith  ", "bob jones", "bob jones", "CHRIS T.", "  Dana W.", "Unidentified", "Ethan Hunt"],
        "Amount_USD": ["$250.00", "120", "120", "$450.50", "MISSING", "300.00", "-50.00"],
        "Region": ["North", "north", "north", "South", "East", "West", "UNKNOWN"]
    }
    pd.DataFrame(data_1).to_csv("raw_data_inputs/regional_sales_north.csv", index=False)
    
    # Dataset 2: Legacy Subscriptions (Date inconsistencies, Null IDs)
    data_2 = {
        "Transaction_ID": [1006, 1007, 1008, 1009, np.nan],
        "Customer_Name": ["Fiona Gallagher", "George Costanza", "Hannah Baker", "Ian Malcolm", "Unknown Guest"],
        "Amount_USD": ["105.00", "$99.99", "310.00", "0.00", "$15.00"],
        "Region": ["South", "West", "East", "North", "South"]
    }
    pd.DataFrame(data_2).to_csv("raw_data_inputs/regional_sales_south.csv", index=False)
    print("[✓] Messy mock data pipelines generated successfully inside 'raw_data_inputs/' folder.")

def automated_cleaning_pipeline():
    """Reads files sequentially, scrubs formats, strips strings, drops null keys, and unifies frameworks."""
    print("\n[+] Executing Algorithmic Data Scrubbing Engine...")
    all_files = glob.glob("raw_data_inputs/*.csv")
    
    if not all_files:
        raise FileNotFoundError("Error: Target files missing in source space.")
        
    compiled_df = []
    
    for file_path in all_files:
        print(f" Processing: {file_path}")
        df = pd.read_csv(file_path)
        
        # 1. Row Cleanse: Drop instances missing vital Transaction ID anchors
        df = df.dropna(subset=["Transaction_ID"])
        df["Transaction_ID"] = df["Transaction_ID"].astype(int)
        
        # 2. String Cleanse: Clean whitespace padding and standardize letter casing structures
        df["Customer_Name"] = df["Customer_Name"].astype(str).str.strip().str.title()
        df["Region"] = df["Region"].astype(str).str.strip().str.upper()
        df["Region"] = df["Region"].replace({"NORTH": "NORTH", "SOUTH": "SOUTH", "EAST": "EAST", "WEST": "WEST"})
        
        # 3. Numeric Parse: Extract valid numeric digits from currency annotations
        def parse_amount(val):
            val_str = str(val).strip().upper()
            if "MISSING" in val_str or not val_str:
                return np.nan
            cleaned = re.sub(r'[^\d\.\-]', '', val_str) # Strip currency tags
            try:
                num = float(cleaned)
                return num if num >= 0 else np.nan # Drop non-sensical negative business costs
            except ValueError:
                return np.nan
                
        df["Amount_USD"] = df["Amount_USD"].apply(parse_amount)
        compiled_df.append(df)
        
    # Standardize distinct column matrices into a single baseline table
    master_df = pd.concat(compiled_df, ignore_index=True)
    
    # 4. Handle remaining missing metrics using localized rolling median values
    median_amount = master_df["Amount_USD"].median()
    master_df["Amount_USD"] = master_df["Amount_USD"].fillna(median_amount)
    
    # 5. Row Cleanse: Deduplicate transactional data records
    initial_count = len(master_df)
    master_df = master_df.drop_duplicates(subset=["Transaction_ID"], keep="first")
    print(f"[✓] Deduplication dropped {initial_count - len(master_df)} trailing duplicate entries.")
    
    os.makedirs("cleaned_reports", exist_ok=True)
    master_df.to_csv("cleaned_reports/master_clean_transactions.csv", index=False)
    print("[✓] Processed database stored at 'cleaned_reports/master_clean_transactions.csv'")
    return master_df

def generate_automated_reports(df):
    """Aggregates operational charts, calculates business totals, and saves visual summaries."""
    print("\n[+] Generating Visual Summaries & Executive Analytics Reports...")
    
    # Aggregation Summary Calculations
    regional_metrics = df.groupby("Region")["Amount_USD"].agg(["sum", "mean", "count"]).reset_index()
    regional_metrics.columns = ["Region", "Total_Revenue_USD", "Average_Ticket_Size_USD", "Transaction_Volume"]
    
    # Save a clean Excel Workbook Summary Report
    report_path = "cleaned_reports/executive_summary_report.xlsx"
    with pd.ExcelWriter(report_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name="Master Transactions Data", index=False)
        regional_metrics.to_excel(writer, sheet_name="Regional KPI Aggregations", index=False)
        
    print(f"[✓] Multi-Sheet Excel Document bundled at '{report_path}'")
    
    # Construct Descriptive Visual Analytics Chart
    plt.figure(figsize=(8, 5))
    colors = ["#4EA8DE", "#56CFE1", "#72EFDD", "#80FFDB"][:len(regional_metrics)]
    plt.bar(regional_metrics["Region"], regional_metrics["Total_Revenue_USD"], color=colors, edgecolor='grey', width=0.6)
    
    plt.title("Automated Reporting: Revenue Metrics Across Regions", fontsize=13, fontweight='bold', pad=15)
    plt.xlabel("Business Regions", fontsize=11)
    plt.ylabel("Total Revenue (USD)", fontsize=11)
    plt.grid(axis='y', linestyle='--', alpha=0.5)
    
    plt.tight_layout()
    chart_path = "cleaned_reports/regional_revenue_summary.png"
    plt.savefig(chart_path, dpi=300)
    plt.close()
    print(f"[✓] Analytics Chart exported to '{chart_path}'")
    print("\n=======================================================")
    print(" PIPELINE EXECUTION SUMMARY: SUCCESSFUL")
    print("=======================================================")

if __name__ == "__main__":
    generate_dirty_data()
    cleaned_dataframe = automated_cleaning_pipeline()
    generate_automated_reports(cleaned_dataframe)
