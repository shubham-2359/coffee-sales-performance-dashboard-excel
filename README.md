# ☕ Coffee Sales Performance Dashboard

<p align="center">
  <strong>Turning Coffee Shop Transactions into Actionable Business Insights ☕📊</strong>
  <br/>
  An interactive Microsoft Excel dashboard that analyzes sales performance, customer ordering patterns, product demand, and store-level results to support data-driven business decisions.
</p>

---

## 📸 Dashboard Preview

<p align="center">
  <img src="https://github.com/shubham-2359/coffee-sales-performance-dashboard-excel/blob/main/Coffee-Sales-Dashboard.png" alt="Coffee Sales Performance Dashboard" width="100%"/>
</p>

<p align="center">
  <em>Interactive dashboard with KPI cards, sales charts, and dynamic filters.</em>
</p>

---

## 📑 Table of Contents

- [🎯 Project Objective](#-project-objective)
- [💼 Business Problems](#-business-problems)
- [📂 Dataset Overview](#-dataset-overview)
- [🛠️ Tools & Technologies](#️-tools--technologies)
- [⚙️ Project Workflow](#️-project-workflow)
  - [🧹 Data Cleaning & Transformation](#-data-cleaning--transformation)
  - [📐 Measure Creation](#-measure-creation)
  - [🎛️ Interactive Filters](#️-interactive-filters)
  - [📊 Dashboard Visualizations](#-dashboard-visualizations)
- [🔍 Key Insights](#-key-insights)
- [💡 Business Recommendations](#-business-recommendations)
- [🚀 Future Scope](#-future-scope)
- [👨‍💻 Author](#-author)

---

## 🎯 Project Objective

The primary objective of this project is to transform raw coffee shop transaction data into an interactive sales performance dashboard using Microsoft Excel.

The dashboard provides a consolidated view of key performance indicators, store performance, hourly ordering trends, product category demand, and revenue contribution.

**The project aims to:**

- 📈 Monitor overall sales performance through key business KPIs.
- 🏪 Compare revenue and order volume across store locations.
- ⏰ Identify peak ordering hours and daily demand patterns.
- ☕ Discover top-selling product categories and revenue-generating products.
- 📦 Understand product-size preferences and improve inventory planning.
- 💡 Generate actionable insights to support operational and commercial decisions.

---

## 💼 Business Problems

The dashboard addresses practical business questions faced by coffee shop managers and decision-makers.

| Business Question | Dashboard Analysis | Business Value |
|---|---|---|
| 💰 How much revenue does the business generate? | Total Revenue KPI | Monitor overall sales performance. |
| 🧾 How many transactions are completed? | Total Orders KPI | Understand transaction volume. |
| 🛒 How many items are sold? | Quantity Sold KPI | Support demand and inventory planning. |
| 💳 What is the average revenue per order? | Average Order Value KPI | Evaluate opportunities to increase spending per transaction. |
| ⏰ When is customer demand highest? | Quantity Ordered by Hour | Optimize staffing and product preparation. |
| 🏪 Which store performs best? | Store Performance: Orders & Revenue | Compare locations and investigate performance gaps. |
| ☕ Which categories sell the most? | Quantity Sold by Product Category | Prioritize inventory for high-demand categories. |
| 🏆 Which products generate the most revenue? | Top 5 Product Types by Revenue | Focus on important revenue contributors. |
| 📏 Which product sizes are most popular? | Quantity Sold by Product Size | Improve product mix and size-level planning. |
| 📅 How does performance vary by weekday? | Sales Performance by Day of Week | Support staffing and promotional planning. |

---

## 📂 Dataset Overview

📌 **Data Source:** Coffee Shop Sales dataset published by [Maven Analytics](https://mavenanalytics.io/).

The supplied Excel workbook is organized into three worksheets covering the data preparation and analysis process.

| Worksheet | Description |
|---|---|
| 📄 Unclean Dataset | Original transaction-level data before cleaning and transformation. |
| 🧹 Clean Dataset | Structured and cleaned records prepared for analysis. |
| 📊 Excel Measures & Pivot Tables | KPI calculations and PivotTable summaries supporting the dashboard visualizations. |

### 🔑 Key Data Fields

- **Transaction ID:** Unique identifier for each transaction.
- **Transaction Date & Time:** Supports hourly, daily, and monthly analysis.
- **Transaction Quantity:** Number of items sold.
- **Store ID & Store Location:** Enables store-level performance comparisons.
- **Product ID:** Identifies individual products.
- **Unit Price:** Selling price per item.
- **Product Category:** Groups products into categories such as Coffee, Tea, Bakery, and Drinking Chocolate.
- **Product Type:** Identifies specific product varieties.
- **Product Details / Size:** Supports product-level analysis and size classification.

The cleaned dataset includes standardized size labels and analytical fields such as Revenue, Month Name, and Day Name.

> 🗓️ **Reporting scope:** The dashboard includes month filters from January through June. Results should be interpreted according to the available dataset and selected filters.

---

## 🛠️ Tools & Technologies

| Technology | Application |
|---|---|
| **Microsoft Excel** | Dashboard development, calculations, and reporting. |
| **Power Query Editor** | Data cleaning, transformation, and standardization. |
| **PivotTables** | Aggregating sales data by product, store, hour, and day. |
| **PivotCharts & Excel Charts** | Visualizing sales trends and performance comparisons. |
| **Excel Data Model / DAX Measures** | Calculating KPIs and analytical metrics. |
| **Slicers / Report Filters** | Enabling interactive analysis by month and day of the week. |

---

## ⚙️ Project Workflow

### 🧹 Data Cleaning & Transformation

**Tool: Power Query Editor**

#### 1️⃣ Data Import & Connection
- Open Microsoft Excel and navigate to **Data → Get Data → From File → From Excel Workbook**.
- Select the source workbook, `Coffee Shop Sales Datadet.xlsx`.
- Choose the required transaction table in the Navigator window.
- Click **Transform Data** to open Power Query Editor.

#### 2️⃣ Data Inspection
- Reviewed columns, sample records, and assigned data types.
- Inspected transaction IDs, dates, times, quantities, prices, store locations, and product details.
- Checked product categories and product types for inconsistencies.

#### 3️⃣ Product Size Standardization

Standardized abbreviated product sizes into readable labels:

| Original Code | Standardized Value |
|---|---|
| `LG` | Large |
| `RG` | Regular |
| `SM` | Small |

Records with missing or undefined size information were retained under an appropriate category, such as **Not Defined**, where applicable.

#### 4️⃣ Missing Values & Data Types
- Reviewed null values, blank entries, and inconsistent text.
- Applied appropriate cleaning and standardization.
- Assigned suitable data types to numeric, date, time, and text columns.
- Ensured quantities and prices could be aggregated correctly.

#### 5️⃣ Derived Columns

Created or validated the following analytical fields:

- **Revenue:** Transaction quantity multiplied by unit price.
- **Size:** Standardized product size.
- **Month Name:** Derived from the transaction date.
- **Day Name:** Derived from the transaction date.

#### 6️⃣ Load the Cleaned Data
- Selected **Home → Close & Load** to load the transformed data into Excel.
- Used the cleaned dataset as the source for PivotTables, measures, and dashboard charts.

### 📐 Measure Creation

The following measures support the dashboard's key performance indicators.

| Measure | Formula | Purpose |
|---|---|---|
| 💰 **Total Revenue** | `SUM(Table1[Revenue])` | Calculates total revenue. |
| 🧾 **Total Orders** | `DISTINCTCOUNT(Table1[transaction_id])` | Counts unique transactions. |
| 🛒 **Items Sold** | `SUM(Table1[transaction_qty])` | Calculates total items sold. |
| 💳 **Average Order Value** | `DIVIDE([Total Revenue], [Total Orders], 0)` | Calculates average revenue per order. |

**Formula notes:**
- Replace `Table1` with the actual table name in your Excel Data Model if necessary.
- These formulas use DAX-style measure syntax.
- The Average Order Value measure returns `0` when there are no orders.

### 🎛️ Interactive Filters

The dashboard includes two interactive filters:

- 📅 **Day Name:** Filters sales performance by Sunday, Monday, Tuesday, Wednesday, Thursday, Friday, or Saturday.
- 🗓️ **Month Name:** Filters results by the available months, from January through June.

These filters allow users to explore different time periods and examine changes in business performance.

### 📊 Dashboard Visualizations

| Visualization | Chart Type | Purpose |
|---|---|---|
| 💰 Total Revenue | KPI Card | Displays revenue for the selected data. |
| 🧾 Total Orders | KPI Card | Shows the number of unique transactions. |
| 🛒 Quantity Sold | KPI Card | Displays the total quantity sold. |
| 💳 Average Order Value | KPI Card | Shows average revenue per order. |
| ⏰ Quantity Ordered by Hour | Line Chart | Identifies peak hours and changes in demand. |
| 🏪 Store Performance: Orders & Revenue | Clustered Column Chart | Compares orders and revenue across three store locations. |
| 📏 Quantity Sold by Product Size | Pie Chart | Shows the distribution of sales across size groups. |
| ☕ Quantity Sold by Product Category | Horizontal Bar Chart | Ranks categories by total quantity sold. |
| 🏆 Top 5 Product Types by Revenue | Column Chart | Highlights the five leading product types by revenue. |
| 📅 Sales Performance by Day of Week | Clustered Column Chart | Compares order volume and revenue across weekdays. |

---

## 🔍 Key Insights

*Insights are based on the dashboard values shown. Results may vary depending on the selected filters.*

### 1. 💰 Overall Sales Performance

The dashboard reports approximately $698.8K in total revenue, 149.1K orders, 214.5K quantity sold and an average order value of $4.69. 
These KPIs provide a clear overview of the coffee shop's overall sales performance.

### 2. ⏰ Morning Hours Drive Demand

Ordering activity rises sharply after 6:00 AM and reaches its highest point at approximately **10:00 AM**. Demand then declines and remains relatively steady through much of the afternoon before dropping toward the end of the trading day.

This pattern highlights the importance of preparing staff and inventory for the morning rush.

### 3. 🏪 Store Performance Varies by Location

**Hell's Kitchen** generates the highest revenue at approximately **$236.5K** and records the highest order count with **50,735 orders**. 
This highlights its strong sales performance and provides an opportunity to explore the factors driving its success.

### 4. ☕ Top-Selling Categories and Products

Coffee leads product volume with 89,250 items sold, followed by Tea with 69,737 items. 
Barista Espresso leads the top five product types by revenue at $91,406.20.

### 5. 📏 Product Size Distribution

Regular and Large sizes lead with 70,186 and 68,540 items, respectively. 
The 54,321 items classified as Not Defined highlight an opportunity to improve data quality.

### 6. 📅 Day-Level Performance Can Guide Planning

The dashboard shows broadly consistent order and revenue levels throughout the week. 
Comparing exact daily totals can help identify relatively stronger and weaker trading days for staffing and promotional planning.

---

## 💡 Business Recommendations

### 1. ⏰ Optimize Staff Scheduling
Schedule sufficient baristas and counter staff during the morning rush, especially between **8:00 AM and 10:00 AM**. 
Review staffing requirements during slower periods to balance service quality and labor costs.

### 2. 📦 Improve Inventory Planning
Prioritize stock availability for Coffee and Tea products because they lead quantity sold. 
Use store-level and time-based demand patterns to reduce stockouts and avoid excess inventory.

### 3. 🏪  Improve Store Performance
Analyze differences in revenue and order volume across locations to identify opportunities to increase average order value and sales.

### 4. 🏆 Promote High-Revenue Products
Monitor the availability and performance of Barista Espresso and other leading product types. 
Test complementary products, relevant add-ons, or bundles to encourage higher order values while maintaining customer satisfaction.

### 5. 🧹 Improve Product-Size Data Quality
Investigate the large number of items classified as Not Defined. 
Standardize product-size information during data entry or transformation to improve product-mix analysis and inventory planning.

### 6. 🎯 Use Day-Level Trends for Promotions
Identify relatively slower trading days and test targeted promotions. 
Compare results with a baseline to determine whether the campaigns generate additional revenue rather than simply discounting existing purchases.

---

## 🚀 Future Scope

- 📊 **Advanced Analytics:** Identify sales trends and customer purchasing patterns.
- 📈 **Sales Forecasting:** Predict future sales using historical data.
- 🏪 **Store Performance:** Explore more insights to compare store performance.
- 🎯 **Enhanced Dashboard:** Add more interactive filters and visualizations.
- ⚡ **Business Optimization:** Improve inventory planning and staff allocation based on peak sales hours.
- 🛠️ **Advanced Tools:** Explore Power BI and SQL for deeper analysis and reporting.

---

## 👨‍💻 Author 
Shubham Maurya      
🎯 Aspiring Data Analyst passionate about transforming raw data into meaningful insights and supporting data-driven decisions.

💻 Skills Used: Microsoft Excel | Power Query | Data Cleaning | PivotTables | Data Visualization | Sales Analysis
<p>
  <a href="https://www.linkedin.com/in/shubham-maurya-a53b57c59/">
    <img src="https://img.shields.io/badge/LinkedIn-Connect-0A66C2?style=for-the-badge&logo=linkedin&logoColor=white" alt="LinkedIn Profile"/>
  </a>
</p>

---

<p align="center">
  ☕ <strong>Thank you for visiting this project!</strong> 📊
  <br/>
  <em>Turning data into insights, and insights into better business decisions.</em>
</p>
