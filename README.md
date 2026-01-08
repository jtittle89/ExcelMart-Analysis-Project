# ExcelMart-Analysis-Project
Sales Performance Excel Analysis Project

📌 Project Overview

This project demonstrates an end-to-end advanced Excel analytics solution built to analyze sales performance for a fictional retail company, ExcelMart. The goal was to showcase professional-grade Excel capabilities commonly used by data analysts, including data modeling, automation, dashboarding, and business insights.

The solution integrates multiple datasets, automates data refresh and reporting, and delivers an interactive executive-level dashboard.

🧩 Data Sources

The project uses five relational datasets:

Sales_Data.csv – Transaction-level sales data

Products.csv – Product categories and pricing

Stores.csv – Store locations and regions

Salespeople.csv – Sales representative details

Customers.csv – Customer demographics and segments

All datasets are loaded and transformed using Power Query.

🛠 Tools & Technologies

Microsoft Excel

Power Query – Data cleaning, transformations, and automated refresh

Power Pivot & Data Model – Star schema modeling and relationships

DAX Measures – KPI calculations and dynamic aggregations

PivotTables & PivotCharts

Conditional Formatting

VBA Macros – Automation and PDF export

Dashboard Design & Reporting

🔄 Data Modeling

A star schema was implemented:

Fact Table: Sales_Data

Dimension Tables: Products, Stores, Salespeople, Customers

Relationships were created using primary and foreign keys to enable scalable analysis and efficient filtering across visuals.

📈 Key KPIs & Metrics

Total Sales Revenue

Total Units Sold

Average Order Value (AOV)

Profit & Profit Margin

Sales by Product, Store, Region, and Customer Segment

Time-based performance trends

All KPIs are built using DAX measures to ensure dynamic recalculation based on slicers and filters.

📊 Dashboard Features

Interactive KPI cards

Slicers for time period, region, store, product, and customer segment

Trend analysis visuals (monthly & yearly)

Conditional formatting for performance insights

Executive-friendly, single-page layout optimized for PDF export

<img width="1470" height="759" alt="image" src="https://github.com/user-attachments/assets/bcc342cd-b78e-4223-92f4-a21da2542f9e" />


⚙️ Automation (VBA)

The project includes VBA automation to:

Refresh all Power Query connections and PivotTables

Export the dashboard as a single-page, landscape PDF

Save outputs directly to the project folder

Display user confirmation messages after export

This automation reduces manual effort and ensures reports are always generated using the most current data.

<img width="754" height="413" alt="image" src="https://github.com/user-attachments/assets/57422ba5-15f0-45c6-9e4e-7a1442f65221" />


🎯 Results & Business Value

Reduced manual reporting effort through automated refresh and export

Enabled fast, self-service analysis with interactive dashboards

Delivered a scalable Excel model suitable for executive reporting

Demonstrated advanced Excel skills aligned with real-world analyst roles

🚀 How to Use This Project

Download the Excel workbook and CSV datasets

Open the workbook and refresh data using the Refresh Data macro

Interact with the dashboard using slicers

Export the dashboard using the Export to PDF macro
