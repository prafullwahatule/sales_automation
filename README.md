# 📊 Daily Sales Report Automation using Python

## 🧠 Project Overview

This project demonstrates a **real-life end-to-end data automation pipeline** built using **Python**. The automation is designed to run **daily without manual intervention** and is suitable for enterprise environments.

The script automatically:

* Picks the latest sales Excel file
* Cleans and validates data
* Calculates key business KPIs
* Generates an Excel report
* Creates visual charts
* Can be scheduled to run daily using **Windows Task Scheduler**

This project reflects **industry-standard automation practices** commonly used by Data Analysts.

---

## 🎯 Business Problem

Manual daily sales reporting is time-consuming, error-prone, and inefficient. Analysts often need to:

* Open Excel files daily
* Clean raw data
* Calculate KPIs
* Create reports and charts

This project eliminates all manual steps by automating the entire workflow.

---

## 🚀 Solution

A Python-based automation pipeline that:

1. Detects the latest sales Excel file
2. Performs data cleaning and validation
3. Calculates KPIs
4. Generates formatted Excel reports
5. Creates revenue visualizations
6. Runs automatically on a daily schedule using **Windows Task Scheduler**

---

## 🛠️ Tools & Technologies Used

* **Python 3.13**
* **pandas** – Data cleaning & analysis
* **matplotlib** – Automated visualizations
* **openpyxl** – Excel report generation
* **Windows Task Scheduler** – Scheduling automation

---

## 📁 Project Structure

```
sales_automation/
│
├── data/
│   └── sales_2026_01_27.xlsx
│
├── output/
│   ├── Daily_Sales_Report.xlsx
│   └── revenue_chart.png
│
├── automation.py
├── requirements.txt
└── README.md
```

---

## 📊 Input Data Format

**File Name Pattern:** `sales_YYYY_MM_DD.xlsx`

### Required Columns

| Column Name | Description             |
| ----------- | ----------------------- |
| OrderID     | Unique order identifier |
| Date        | Order date              |
| Product     | Product name            |
| Category    | Product category        |
| Quantity    | Units sold              |
| Price       | Price per unit          |

---

## 📈 KPIs Calculated

* **Total Revenue**
* **Total Orders**
* **Average Order Value (AOV)**
* **Revenue by Category**

---

## 📄 Output Generated

### 1️⃣ Excel Report – `Daily_Sales_Report.xlsx`

* **Sheet 1:** Raw_Data
* **Sheet 2:** Revenue_By_Category

### 2️⃣ Chart – `revenue_chart.png`

* Bar chart showing revenue distribution by category

---

## ▶ How to Run the Project

### 1️⃣ Install Dependencies

```bash
pip install -r requirements.txt
```

### 2️⃣ Run the Script Manually

```bash
python automation.py
```

---

## ⏰ Daily Scheduling Using Windows Task Scheduler

The automation can run **daily without manual intervention** using Windows Task Scheduler.

### Steps:

1. Open **Task Scheduler** (`taskschd.msc`) on Windows.
2. Click **Create Basic Task** and provide a name & description.
3. Trigger → Select **Daily** and set the time.
4. Action → **Start a Program**

   * Program/script: `C:\Users\Prafull Wahatule\AppData\Local\Programs\Python\Python313\python.exe`
   * Add arguments: `automation.py`
   * Start in: `C:\Users\Prafull Wahatule\Desktop\sales_automation`
5. Check **Open the Properties dialog box for this task when I click Finish** to configure advanced settings:

   * Run whether user is logged on or not
   * Run with highest privileges
6. Save the task. Task will now execute the Python script automatically at the scheduled time.

This ensures **consistent daily reporting** with zero manual effort.

---

## 🧠 Key Features

✔ Dynamic file detection  
✔ Robust data validation  
✔ Error handling & safety checks  
✔ Scheduler-ready Python script  
✔ Industry-level automation design

---

## 📌 Future Enhancements

* Email report automation
* SQL database integration
* Logging to file
* Power BI dataset auto-refresh
* Cloud deployment

---

## 👤 Author

**Prafull Wahatule**
Data Analyst

---

⭐ If you find this project useful, feel free to star the repository!
