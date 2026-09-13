**E-Commerce Sales Performance Analysis (2022–2024)**

Project Overview
An end-to-end analytics project on 3 years of e-commerce order data — from raw data cleaning through Advanced Excel analysis to a Power BI dashboard design, closing with data-driven business recommendations. Built to demonstrate both technical analytics skill and business interpretation, not just dashboard aesthetics.

Business Problem
Management needs visibility into where sales and profit are coming from, which categories/regions/products are under- or over-performing, and why profitability weakened in 2024 despite three years of sales growth.

**Dataset**

Source: data/ecommerce_sales_data.xlsx
Size: 349 orders × 7 fields (Order Date, Product Name, Category, Region, Quantity, Sales, Profit)
Coverage: Jan 2022 – Dec 2024 | 10 products | 3 categories | 4 regions
Data quality: No missing values, no duplicate rows, no negative or invalid Sales/Profit values — the dataset arrived clean, so this project focuses on structuring and Deriving. 


**Tools & Technologies**

Excel (Advanced Formulas, Pivot-style Summaries, Conditional Formatting, Data Validation) · Power BI (Data Modeling, DAX) · Business Analysis

**Data Preparation**

* Structured the raw data as a formal Excel Table (Raw_Data)
* Built a Cleaned_Data table linked by formula (not hardcoded copies) so it stays in sync with the source
* Added derived columns: Year, Quarter, Month Name, Profit Margin %, Avg Unit Price, Order Value Tier (Low/Medium/High)
* Applied data validation dropdowns on Category and Region to prevent future entry errors
* Applied conditional formatting (color scale on margin, data bars on sales) to surface outliers visually.


Advanced Excel Analysis
Technique	Where used	Why
SUMIFS / COUNTIFS	Category, Region, Year, Product summary tables	Aggregate sales/profit/orders by dimension without manual pivoting
IFS (nested logic)	Order Value Tier column	Classify orders into Low/Medium/High bands
INDEX / MATCH	KPI_Dashboard "Best-Selling Product" / "Lowest-Margin Product"	Dynamic lookups that update automatically if the data changes
Date/Text functions (YEAR, MONTH, TEXT, ROUNDUP)	Quarter, Month Name columns	Enable time-based grouping
Excel Tables	Raw_Data, Cleaned_Data, Product_Summary	Auto-expanding structured ranges for formulas and charts
Conditional Formatting	Margin %, Sales columns	Instantly flag low-margin rows/products
Data Validation	Category, Region columns	Enforce consistent category/region entry
Native charts	Sales-by-Product bar chart, Yearly trend line chart	Visual summary directly in Excel

File: **excel/Ecommerce_Sales_Analysis.xlsx** — all KPIs are live formulas (SUMIFS/COUNTIFS/INDEX-MATCH)
Power BI Dashboard


**Key KPIs**


KPI- Values
Total Sales- ₹10,49,470
Total Profit- ₹1,86,845.53
Average Profit Margin- 17.8%
Total Orders- 349
2023 → 2024 Profit Change‑ 12.5%


**Key Insights**


* Sales dipped 2.4% in 2024 while profit fell 12.5% — margin compressed from 18.82% to 16.87%
* Q4 2024 sales dropped 43% quarter-over-quarter after three quarters of growth
* Electronics (47.5%) and Accessories (45.4%) drive 93% of revenue; Office is a minor 7.1%
* Office has the weakest margin (16.46%) of the three categories
* North region leads on both sales and margin; West has strong sales but the lowest margin (16.56%)
* South's Office sales (₹5,954) are roughly a quarter of every other region — a likely stocking gap
* Monitor has the best product margin (20.69%); Smartwatch has the weakest (15.21%)
* Headphones is the top-selling product; Printer is the weakest


**Business Recommendations**

* Investigate the 2024 margin compression — audit discounting, supplier cost, or shipping cost changes that outpaced the 2.4% sales dip to explain the 12.5% profit fall.
* Root-cause the Q4 2024 sales drop before the next Q4 — check inventory, marketing spend, and competitor activity in that window specifically, since it breaks an otherwise steady upward trend.
* Fix the South Office gap — confirm Office SKUs are actively stocked/listed in South; this looks like a low-cost, high-confidence fix (~₹15K–18K recoverable).
* Rebalance category investment toward Electronics and Accessories, and keep Office lean rather than expanding it further given its low share and weak margin.
* Shift promotional focus toward higher-margin products (e.g., Monitor) and review costing/pricing on low-margin products (e.g., Smartwatch) to lift blended margin without needing extra volume.


**Technical Skills Demonstrated**

* Advanced Excel (SUMIFS, COUNTIFS, INDEX-MATCH, IFS, Data Validation, Conditional Formatting, Excel Tables)
* Data Cleaning & Data Management
* KPI Development
* Power BI Data Modeling & DAX
* Dashboard Design
* Business Analysis & Insight Generation
* Data Visualization

**Executive overview**

<img width="1252" height="637" alt="image" src="https://github.com/user-attachments/assets/e7836369-6604-4333-a98d-ff7fb93755bc" />



**Detailed analysis**










