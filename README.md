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


**Advanced Excel Analysis**

Technique	              
SUMIFS / COUNTIFS	      
IFS (nested logic)	    
INDEX / MATCH	          
Date/Text functions (YEAR, MONTH, TEXT, ROUNDUP)    
Excel Tables	          
Conditional Formatting	Margin %, Sales columns	          
Data Validation	        
Native charts

Where Used

Category, Region, Year, Product summary tables

Order Value Tier column

KPI_Dashboard "Best-Selling Product" / "Lowest-Margin Product"

Quarter, Month Name columns

Raw_Data, Cleaned_Data, Product_Summary

Margin %, Sales columns

Category, Region columns

Sales-by-Product bar chart, Yearly trend line chart

Why

Aggregate sales/profit/orders by dimension without manual pivoting

Classify orders into Low/Medium/High bands

Dynamic lookups that update automatically if the data changes

Enable time-based grouping

Auto-expanding structured ranges for formulas and charts

Instantly flag low-margin rows/products

Enforce consistent category/region entry

Visual summary directly in Excel


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

<img width="1470" height="901" alt="image" src="https://github.com/user-attachments/assets/0ade3e10-cd10-46bc-afc6-15be739e4cc5" />




**Power BI Dashboard — Design Specification**

1. Data Model

Column	            Type
Order Date	        Date
Product Name	      Text
Category	          Text
Region	            Text
Quantity	          Whole number
Sales	              Currency
Profit	            Currency


2. Power Query (M) — Data Prep Steps

* Source → ecommerce_sales_data.xlsx
* Change types: Order Date → Date, Quantity → Whole Number, Sales/Profit → Decimal
* Trim/Clean Category, Region, Product Name (defensive step — no dirty text found, but standard practice)
* Verify no nulls or duplicates (confirmed clean in source analysis)
* Load to model

3. Calculated Columns

Order Year = YEAR(Sales[Order Date])
Order Quarter = "Q" & ROUNDUP(MONTH(Sales[Order Date])/3, 0)
Order Value Tier = IF(Sales[Sales] < 1000, "Low", IF(Sales[Sales] < 4000, "Medium", "High"))


4. Core DAX Measures


Total Sales = SUM(Sales[Sales])
Total Profit = SUM(Sales[Profit])
Total Orders = COUNTROWS(Sales)
Profit Margin % = DIVIDE([Total Profit], [Total Sales], 0)
Avg Order Value = DIVIDE([Total Sales], [Total Orders], 0)

Sales PY = CALCULATE([Total Sales], SAMEPERIODLASTYEAR('Date'[Date]))
Sales YoY % = DIVIDE([Total Sales] - [Sales PY], [Sales PY], 0)

Sales QoQ % =
VAR CurrQ = [Total Sales]
VAR PrevQ = CALCULATE([Total Sales], DATEADD('Date'[Date], -1, QUARTER))
RETURN DIVIDE(CurrQ - PrevQ, PrevQ, 0)

% of Total Sales = DIVIDE([Total Sales], CALCULATE([Total Sales], ALL(Sales)), 0)

Rank by Sales (Product) = RANKX(ALL(Sales[Product Name]), [Total Sales], , DESC)





**Business Insights**

Dataset: 349 orders across 10 products, 3 categories, and 4 regions (Jan 2022 – Dec 2024). Total Sales: ₹10,49,470 | Total Profit: ₹1,86,845.53 | Average Profit Margin: 17.8%

1. Growth stalled and margin compressed in 2024 despite three years of expansion

Sales grew from ₹2,63,624 (2022) to ₹3,97,752 (2023, +50.9%), but slipped to ₹3,88,094 in 2024 (‑2.4%). Profit fell harder — from ₹74,861.8 to ₹65,481.1 (‑12.5%) — pulling the average margin down from 18.82% to 16.87%. Driver: Costs or discounting grew faster than sales in 2024, since the sales decline (‑2.4%) is far smaller than the profit decline (‑12.5%). Impact: The business is trading margin for volume — a warning sign that profitability, not just revenue, needs monitoring going into 2025.

2. Q4 2024 shows a sharp, isolated demand drop

Quarterly sales climbed steadily through 2024 (₹92,524 → ₹1,07,811 → ₹1,19,536 in Q1–Q3) before crashing to ₹68,223 in Q4 — a 43% quarter-over-quarter fall and the weakest quarter since Q3 2022. Profit margin in the same quarter fell to 16.62%. Driver: This break in an otherwise rising trend points to a Q4-specific event — inventory shortage, reduced marketing spend, or a competitive/seasonal disruption — rather than a gradual decline. Impact: If unaddressed, this could signal a recurring seasonal risk heading into the next Q4.


3. Electronics and Accessories jointly drive 93% of revenue; Office is a minor, low-priority category

Electronics contributed ₹4,98,091 (47.5% of sales) and Accessories ₹4,76,710 (45.4%), while Office totaled just ₹74,669 (7.1%) across only 31 orders. Driver: Office products (mainly Printers) have both far fewer orders and lower average order value (₹2,409 vs. ₹3,310 for Accessories). Impact: Category strategy, inventory investment, and marketing budget should be weighted toward Electronics and Accessories, with Office kept lean rather than expanded.

4. Electronics and Accessories carry near-identical margins, but Office lags behind

Margins are close for the two lead categories — Electronics 17.92%, Accessories 17.89% — while Office trails at 16.46%, the weakest of the three. Driver: Printer-heavy Office orders have a lower average unit price and thinner markup than Electronics/Accessories items like Monitors (20.69% margin) and Smartphones (19.85%). Impact: Without a margin improvement (pricing or supplier renegotiation), Office is not a segment worth growing further.

5. North region leads on both scale and profitability; West lags on margin despite strong sales

North posted the highest total sales (₹2,88,713) and the best margin (18.65%). West generated similar-scale sales (₹2,80,039, 2nd highest) but the lowest margin of any region (16.56%). Driver: West's Electronics sales are the highest of any region (₹1,38,449), but that volume isn't translating into proportional profit — suggesting higher discounting or higher fulfillment cost in that region. Impact: West needs a margin/cost review even though its top-line performance looks healthy — a case where revenue numbers alone would mislead management.


6. Office category has a striking regional gap: South is nearly absent

South's Office sales are just ₹5,954, compared to ₹21,900–₹23,900 in every other region — roughly one-quarter of the norm, despite South performing normally in the other two categories. Driver: Likely a distribution/listing gap (Office products may not be well-stocked or promoted in South) rather than a genuine demand difference, since South's Accessories and Electronics sales are in line with other regions. Impact: A quick, low-cost fix (ensure Office SKUs are actively listed/stocked in South) could recover an estimated ₹15,000–₹18,000 in incremental sales to bring South to parity.


7. Monitor is the most profitable product per rupee of sales; Smartwatch is the least

Monitor carries the highest margin at 20.69% (₹23,413.6 profit on ₹1,13,151 sales), while Smartwatch is lowest at 15.21% (₹13,154 profit on ₹86,480 sales) — a 5.5 percentage-point gap between the best and worst product margins. Driver: Products differ meaningfully in profitability even within the same broad category (both are Electronics/Accessories-adjacent), meaning category-level averages can mask product-level problems. Impact: Margin-focused promotion (push Monitors, re-price or re-negotiate Smartwatch costs) could lift blended margin without needing new sales volume.


8. Headphones is the top-selling product; Printer is the weakest

Headphones leads all products with ₹1,47,540 in sales across 41 orders, while Printer trails at ₹74,669 across 31 orders — consistent with Office being the smallest category overall. Driver: Headphones benefits from being a high-frequency, broadly-appealing Accessories item, while Printer is a low-frequency, single-purchase item. Impact: Confirms that assortment depth (more SKUs like Headphones) matters more for growth than trying to scale a naturally low-repeat category like printers.


  









