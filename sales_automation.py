# =====================================
# DAILY SALES REPORT AUTOMATION
# Author: Prafull Wahatule
# =====================================

import pandas as pd
import os
import glob
import matplotlib.pyplot as plt
from datetime import datetime
import sys

# -----------------------------
# CONFIG
# -----------------------------
DATA_FOLDER = "data/"
OUTPUT_FOLDER = "output/"
REPORT_NAME = "Daily_Sales_Report.xlsx"
CHART_NAME = "revenue_chart.png"

os.makedirs(OUTPUT_FOLDER, exist_ok=True)

print("\n🚀 Sales Automation Started")
print("⏰ Run Time:", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

# -----------------------------
# STEP 1: PICK LATEST FILE
# -----------------------------
files = glob.glob(os.path.join(DATA_FOLDER, "*.xlsx"))

if not files:
    print("❌ No Excel files found in data folder")
    sys.exit()

latest_file = max(files, key=os.path.getctime)
print(f"📂 Processing file: {latest_file}")

# -----------------------------
# STEP 2: LOAD DATA (SAFE)
# -----------------------------
try:
    df = pd.read_excel(latest_file)
    print("✅ Data loaded successfully")
except Exception as e:
    print("❌ Failed to load Excel file:", e)
    sys.exit()

# -----------------------------
# STEP 3: VALIDATE COLUMNS
# -----------------------------
required_columns = {'OrderID', 'Date', 'Product', 'Category', 'Quantity', 'Price'}

missing_cols = required_columns - set(df.columns)
if missing_cols:
    print(f"❌ Missing columns: {missing_cols}")
    sys.exit()

# -----------------------------
# STEP 4: DATA CLEANING
# -----------------------------
df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
df.dropna(subset=['OrderID', 'Category', 'Quantity', 'Price'], inplace=True)

df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce')
df['Price'] = pd.to_numeric(df['Price'], errors='coerce')

df.dropna(subset=['Quantity', 'Price'], inplace=True)

df['Revenue'] = df['Quantity'] * df['Price']
print("🧹 Data cleaned successfully")

# -----------------------------
# STEP 5: KPI CALCULATIONS
# -----------------------------
total_revenue = df['Revenue'].sum()
total_orders = df['OrderID'].nunique()
avg_order_value = total_revenue / total_orders if total_orders != 0 else 0

summary = (
    df.groupby('Category', as_index=False)
      .agg(Revenue=('Revenue', 'sum'),
           Orders=('OrderID', 'nunique'))
)

print("📊 KPIs calculated")

# -----------------------------
# STEP 6: SAVE EXCEL REPORT
# -----------------------------
report_path = os.path.join(OUTPUT_FOLDER, REPORT_NAME)

try:
    with pd.ExcelWriter(report_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Raw_Data", index=False)
        summary.to_excel(writer, sheet_name="Revenue_By_Category", index=False)

    print("📄 Excel report generated:", report_path)
except Exception as e:
    print("❌ Failed to generate Excel report:", e)
    sys.exit()

# -----------------------------
# STEP 7: CREATE CHART
# -----------------------------
try:
    plt.figure(figsize=(7, 4))
    plt.bar(summary['Category'], summary['Revenue'])
    plt.title("Revenue by Category")
    plt.xlabel("Category")
    plt.ylabel("Revenue")
    plt.tight_layout()

    chart_path = os.path.join(OUTPUT_FOLDER, CHART_NAME)
    plt.savefig(chart_path)
    plt.close()

    print("📈 Chart created:", chart_path)
except Exception as e:
    print("❌ Failed to create chart:", e)

# -----------------------------
# STEP 8: FINAL KPI SUMMARY
# -----------------------------
print("\n========= KPI SUMMARY =========")
print(f"Total Revenue       : ₹{total_revenue:,.2f}")
print(f"Total Orders        : {total_orders}")
print(f"Average Order Value : ₹{avg_order_value:,.2f}")
print("================================")

print("\n✅ Daily Sales Automation Completed Successfully")
