# BIX Business Performance Analytics Dashboard 
  
## Project Overview
This project features a **dynamic, interactive Excel dashboard** designed to analyze retail sales performance across stores, time periods, and customer segments. Transforming raw transactional data into actionable insights, the solution relies on a **normalized data model** in **Power Pivot** that interconnects a 20,000+ row fact table with specific customer, product, store, and custom date dimensions. Built entirely within Microsoft Excel using **Power Query** for data extraction and advanced **DAX measures** for calculation, the dashboard visualizes critical KPIs—including revenue, profit margins, targets, and refund trends—to support data-driven decision-making.

---
## Problem Statement  
The business required an interactive dashboard to consolidate and analyze transactional data for **revenue tracking**, **profit optimization**, **customer segmentation**, and **operational KPIs** across stores, timeframes, and demographics—but lacked a scalable Excel solution.   

---

## Project Objectives

* **Data Transformation & Modeling:** Normalize and model raw transactional data (capable of handling up to 10M+ rows) to ensure seamless performance and stability without lag.
* **Multi-View Architecture:** Develop **three interconnected dashboards**—Store Performance, Time Frame Analysis, and Profit View—to provide a holistic view of the business.
* **Dynamic Analysis & Segmentation:** Enable granular slicing by time, store, product, and salesperson, alongside customer demographic segmentation (age, gender), to identify **growth opportunities** and **underperforming segments**.
* **Visual Performance Tracking:** Implement dynamic KPIs with visual storytelling elements (conditional formatting, icons, arrows) to compare actuals vs. targets.
* **Key Metric Monitoring:** Track and visualize high-impact metrics, specifically monitoring **Revenue Growth (46%)**, **Profit Margins (42.8%)**, and **Refund Rates (8%)** to drive data-led decision-making.
  
**Technologies Used:**  
- **Excel 365/2021+**: Power Query (ETL), Power Pivot (modeling), DAX (measures).  
- **Visualization**: Charts, conditional formatting, icons/shapes (Flaticon), Zebra add-in.  
- **Automation**: VBA macros (slicer toggle), form controls (group boxes/option buttons).  

---

## Data Set Details  
- **Size**: 20,000+ rows in fact table (transactions); total ~25,000 rows across 5 tables.  
- **Tables**:
  | Table | Columns | Key Fields |
  |-------|---------|------------|
  | **Fact (Transactions)** | 9 | TransactionID, CustomerID, ProductID, StoreID, SalespersonID, Date, QuantitySold, ReturnQuantity  
  | **Customers** | 5 | CustomerID, FirstName, LastName, DOB, Gender  
  | **Products** | 4 | ProductID, ProductName, SalesPrice, CostPrice  
  | **Salespersons** | 3 | SalespersonID, FirstName, LastName  
  | **Stores** | 3 | StoreID, StoreName, Location  
  | **Date** (Custom) | 12 | Date, Year, Month, Quarter, Weekday, Weekend, etc.  
- **Time Period**: Multi-year dataset with quarterly/monthly granularity.  
---

## Key Analyses Performed
1. **Revenue vs Target Variance**: By Store/Month/Quarter (absolute + % with up/down indicators).   
2. **Profitability Metrics**: Margin %, Cost of Goods Sold, MoM growth.   
3. **Return & Refund Analysis**: Total Refund, Refund Rate (8% overall), Return Quantity %.   
4. **Customer Segmentation**: Age groups (0-20, 21-30, etc.), Gender profit %, Top/Bottom 5 customers.   
5. **Time Intelligence**: Weekend/Weekday revenue %, Quarter trends, # Products Sold (distinct).   
6. **Product Performance**: Top products by Profit/Quantity, overall # Products/Transactions.
7. **Store Performance**: Revenue/profit by location with variance % and top/bottom rankings.
---

## key Steps 


---

### Data Cleaning & Preparation (Power Query)  
<table>
  <tr>
    <td>
    <td style="text-align: center;">
      <em>Data cleaning & preparation in power query</em><br />
      <img src="images/data cleaning & preparation in power query.png" alt="data cleaning & preparation in power query" width="500" height="1000" /></td>
    <td>
    <td style="text-align: center;">
      <em>Date table creation in power query</em><br />
      <img src="images/date table creation in power query.png" alt="date table creation in power query.png" width="500" height="400" /></td>
  </tr>
</table>
I Structured ETL pipeline to clean and transform raw CSV data into a relational model. I implemented the following steps sequentially:

1. **Source Import**: Loaded 5 CSV files (transactions, customers, products, salespersons, stores) via **Get Data > From Text/CSV**.  
2. **Header Promotion**: Promoted first row to column headers across all queries.  
3. **Data Type Changes**: Auto-detected and set types (e.g., Date to Date, QuantitySold/ReturnQuantity to Whole Number, IDs to Text).  
4. **Name Merging**: Combined FirstName + " " + LastName into **FullName** columns for customers and salespersons.  
5. **Age Derivation**: Added **Age** column: $$Age = \frac{TODAY() - DOB}{365.25}$$ (rounded to nearest whole number).  
6. **Duplicate Removal**: Eliminated rows based on unique **TransactionID** in fact table.  
7. **Date Dimension Creation**: Generated custom calendar table from min/max dates:
   - Extracted **Year**, **Month Name**, **Month Number**, **Quarter**, **Weekday**, **Weekend** flag.
   - Added sorting columns (e.g., MonthSort = 1-12).
<table>
  <tr>
    <td><img src="images/only connection 1.png" alt="Queries set to **Connection Only**" width="300" height="300" /></td>
    <td><img src="images/only connection 2.png" alt="Queries set to **Connection Only**" width="300" height="300" /></td>
  </tr>
</table>

8. **Load Optimization**: Queries set to **Connection Only** (except calculations table) for Power Pivot efficiency.  

> I used Power Query to handle the heavy lifting of data transformation before loading it into Power Pivot.  

 

### Establishing  Power Pivot Relationships (Data modeling)  
**Star Schema** design in Power Pivot for efficient querying and slicing.  

### Model Structure  
- **Fact Table**: Transactions (20K+ rows) as central hub.  
- **Dimension Tables**: Customers, Products, Salespersons, Stores, Date (one-to-many from fact).  

### Relationships Established  
1. **Fact[CustomerID] → Customers[CustomerID]** (Many:1, single direction).  
2. **Fact[ProductID] → Products[ProductID]** (Many:1).  
3. **Fact[StoreID] → Stores[StoreID]** (Many:1).  
4. **Fact[SalespersonID] → Salespersons[SalespersonID]** (Many:1).  
5. **Fact[Date] → Date[Date]** (Many:1, active for time intelligence).  

**Additional Table**: Blank "Calculations" table to house all DAX measures (no relationships needed).  

| Relationship | Cardinality | Cross-Filter | Purpose |
|--------------|-------------|--------------|---------|
| Fact → Customers | Many:1 | Single | Customer segmentation  
| Fact → Products | Many:1 | Single | Product profitability  
| Fact → Date | Many:1 | Both | Time-based slicing  

**DAX Integration**: Measures reference related tables via **RELATED()** and **USERELATIONSHIP()** for context-aware calculations.
 

### Creating DAX Measures  
**Core Calculations** (stored in blank "calculations" table):  

| Measure | Formula | Purpose |
|---------|---------|---------|
| **Total Revenue** | $$SUMX(FactTable, FactTable[QuantitySold] \times RELATED(ProductTable[SalesPrice]))$$   | Calculates total sales revenue by multiplying quantity sold with sales price across transactions.  
| **COGS** | $$SUMX(FactTable, FactTable[QuantitySold] \times RELATED(ProductTable[CostPrice]))$$   | Computes total cost of goods sold using quantity and cost price.  
| **Profit Margin** | $$[Total Revenue] - [COGS]$$   | Determines gross profit by subtracting COGS from revenue.  
| **Profit %** | $$DIVIDE([Profit Margin], [Total Revenue])$$   | Calculates profit margin as percentage of revenue.  
| **# Transactions** | $$COUNTROWS(FactTable)$$   | Counts total number of transactions.  
| **Total Refund** | $$SUMX(FactTable, FactTable[ReturnQuantity] \times RELATED(ProductTable[SalesPrice]))$$   | Sums refund value based on returned quantity and sales price.  
| **Refund Rate** | $$DIVIDE([Total Refund], [Total Revenue])$$   | Measures refund rate as percentage of total revenue.  
| **# Products Sold** | $$DISTINCTCOUNT(FactTable[ProductID])$$   | Counts unique products sold (excludes unsold inventory).   
| **# Customers** | $$DISTINCTCOUNT(FactTable[CustomerID])$$   | Counts unique customers with purchase history.  
| **Return Rate** | $$DIVIDE(SUM(FactTable[ReturnQuantity]), SUM(FactTable[QuantitySold]))$$   | Calculates ratio of returned items to total sold.  
| **Average Customer Age** | $$AVERAGE(Customers[Age])$$   | Computes average age of purchasing customers.   

**Notes**:  
Custom formats: e.g., $$[<=1000000]\$0.0,,"M";[<=1000]\$0.0,"K";\$0.0$$ for abbreviated millions/Ks.  

---
## Pivot Table Creation & Implementation  

I developed **12+ interconnected PivotTables** as the analytical foundation of the dashboard, leveraging the **Power Pivot data model** to enable dynamic cross-table analysis without complex formulas.   

### Implementation Workflow  

#### **1. Foundation Setup**  
For each dashboard section, I inserted PivotTables directly from the data model (Insert → PivotTable → From Data Model), which automatically inherited all 6 tables and their relationships.   I  placed source PivotTables in different sheets(e.g. analysis 1, analysis 2) to maintain clean dashboard visuals.  

#### **2. Some of the Dashboard-Specific PivotTables**  

**Store  Dashboard** :  
| PivotTable | Rows | Values | Purpose |
|------------|------|--------|---------|
| Store Revenue | StoreName | Total Revenue, Target | Source for revenue bars  
| Store Variance | StoreName | Variance % | Dynamic variance labels  

**Time Frame Dashboard** :  
| PivotTable | Rows/Columns | Values | Purpose |
|------------|--------------|--------|---------|
| Monthly Trends | Month (Columns) | Total Revenue | Line chart with markers  
| Quarter Analysis | Quarter | Revenue, Average | Benchmark comparison  
| Weekday Split | Weekday | Revenue % of Total | Waffle chart source  

**Profit View Dashboard** :  
| PivotTable | Rows | Values | Purpose |
|------------|------|--------|---------|
| Customer Ranking | Customer Name (Top 5) | Profit Margin | Switchable top/bottom bars  
| Age Demographics | Age Group | Profit %, Avg Age | Demographic analysis  
| Product Performance | Product Name (Top 5) | Profit/Quantity | Category breakdown  


#### **3. Advanced Value Settings**  
I configured each measure through Value Field Settings with custom calculations and formatting.   For variance analysis, I used "Show Values As → % Difference From" to calculate month-over-month changes.   Number formats were customized to display abbreviated values (e.g., $5.4M instead of $5,400,000).  

#### **4. Layout Optimization**  
I standardized all PivotTables with compact layout, removed subtotals and grand totals, and disabled banded rows for cleaner visuals.   Each PivotTable was named descriptively (e.g., "pt_StoreRevenue") for easy reference in VBA macros.  


#### **5. Interactive Slicer System**  
I created slicers for Month, Category, and Gender dimensions to enable dynamic filtering.   Each slicer was styled with white fill and bold selection for professional appearance.   Using Report Connections, I synchronized all slicers across the 12+ PivotTables and charts, ensuring one-click filtering across the entire dashboard.   I also implemented a VBA macro to toggle slicer visibility for presentation mode.  


#### **6. PivotChart Integration**  
From each PivotTable, I created corresponding PivotCharts (Insert → PivotChart → Clustered Column).   I enhanced charts with data labels sourced from helper columns using formulas like =TEXT([@Variance],"0.0%"), added targets on secondary axes, and applied gradient fills with the Zebra add-in.    The charts automatically refresh when users interact with slicers, providing instant visual feedback.  

**Key Achievement**: The data model relationships enabled one slicer to filter the fact table and all dimension tables simultaneously, eliminating the need for complex VLOOKUP formulas.

---

## Dashboard Visualizations  

<img src="images/store view db.png" alt="Store Dashboard — Store view (Image 1)" width="800" />

### Store Dashboard (Part 1)  
- **Zebra BI Combo Chart**: Revenue vs Target bars by Store + variance % (up/down arrows).
- **Dynamic KPIs**: Total Revenue, Target, Variance (IF/TEXT for ±12% ↑ with emojis).
- **Month Slicer**: 4-column, custom styled (bottom border on selected).

<img src="images/time frame db.png" alt="Time Frame Dashboard — Time Frame view (Image 2)" width="800" />

### Time Frame Dashboard (Part 2)  
- **Trend Chart**: Smoothed Revenue/Target lines + MoM variance.
- **Waffle Charts**: Weekend/Weekday revenue % (10x10 grids).
- **Quarter Combo**:  Columns + lines + % diff, highlights above average.   
- Variance waterfall (invert if negative, conditional colors).

<img src="images/profit view.png" alt="Profit View Dashboard — Profit View (Image 3)" width="800" />

### Profit View Dashboard (Part 3)  
- **Customer Switcher**: Top/Bottom 5 via Option Buttons (Profit/Revenue).
- **Age Buckets**: Column chart (0-20, 21-30, etc.) + average age.
- **Gender Waffle**: Icon-based (square/male, circle/female).
- **Product KPIs**: Return Rate, # Products, Top Products toggle.
- Dynamic captions (TEXTJOIN + IF for context like "Top 5 Profitable Products: 100/600 Customers").
  
**Design Elements**: Gradient shapes (RGB: 31-140-179 theme), icons (Flaticon PNGs recolored), navigation hyperlinks, toggle VBA for slicers (AI-enhanced macro).   

---

## Design & UX Principles  
- **Color Consistency**: Blue gradient theme (RGB 31-140-179), white/grey backgrounds for readability.  
- **Hierarchy**: Bold KPIs top-center, slicers left, charts right; navigation tabs bottom.   
- **Interactivity**: One-click filters, hover effects, no-scroll layout (fit to screen).  
- **Accessibility**: High contrast, alt-text icons, logical tab order.  
- **Minimalism**: Hide unused elements, dynamic titles/captions for context.  
---
## Key Features:
- **Dynamic KPIs**: Total revenue ($$\$5.4M$$), COGS, profit margin (42.8%), refund rate (8%), and targets with variance indicators (e.g., +46% growth).   
- **Multi-tab Views**: Store analysis, time frame trends, profit views with customer/product breakdowns.   
- **Interactive Controls**: Slicers (month, category), option buttons (top/bottom 5), combo boxes, and hide/show filters via VBA macros.   
- **Custom Visuals**: Waffle charts (weekday revenue split), gradient-filled trends, zebra add-in bars, and conditional formatting with icons/arrows.   
- **Data Model**: Star schema with 1:many relationships (fact → dimensions).  
- **DAX Measures**: 10+ custom metrics (e.g., Profit % = $$DIVIDE([Profit], [Revenue])$$, formatted as %).     
- **Advanced Charts**: Waffle grids (SEQUENCE/conditional formatting), gradient trends, zebra bars (add-in), pie for categories.   
- **Dynamic Elements**: Captions (TEXTJOIN/IF), variance labels ($$TEXT(variance,"+0.0\% \u2191")$$).
- **Navigation**: Hyperlinked buttons between 3 dashboards; paste-as-picture for locked KPIs.    

---

## Key Insights & Business Value  

**BIX slightly outperformed its goal, delivering $5.4M revenue vs a $5.3M target (+3.7% variance) while maintaining a 42.18% profit margin.**

- **Profit structure:** COGS totals $3.1M, resulting in $2.3M profit (profit margin value).

### 1. Revenue & Time-Based Trends
- **Weekday Dominance:** Revenue is disproportionately concentrated on weekdays (**86%**) versus weekends (**14%**), indicating a massive opportunity to boost Saturday/Sunday engagement.
- **Seasonal Volatility:** Performance fluctuates significantly, with strong peaks in **August (+12.5%)** and **November (+11.9%)**, contrasted by dips in **July (-5.0%)** and **December (-3.4%)**.
- **Quarterly Stability:** Despite monthly swings, quarterly revenue remains consistent, hovering near the **$1.36M average** across all four quarters.

### 2. Store Performance Variance
- **Store performance reveals a massive 55.7% execution gap between top locations like Lee-Myers (+31.1%) and underperformers like Novak PLC (-24.6%), indicating that revenue variance is driven by specific operational inconsistencies rather than market-wide issues.**

### 3. Customer & Product Profitability
- **Demographic Focus:** Profitability is driven by older demographics, with the **51+ age group** contributing the most profit (≈ $554K), followed by the **41–50 group** (≈ $504K). Gender contribution is effectively balanced (51.5% Male vs. 48.5% Female).
- **Category Hierarchy:** **Soft Drinks** are the primary profit engine ($718K), generating nearly double the profit of the next best category, **Sports Drinks** ($417K).
- **Operational Leakage:** A high **return rate of 8.03%** and **refund rate of 8.05%** indicates potential quality control or customer satisfaction issues that are eroding margins.

  
**Skills Demonstrated**:Data cleaning ,Data preparation, Data modeling,Power Pivot, Dashboarding, Power Query ETL, DAX , Advanced Charting (waffle/gradients), Form controls/VBA, Custom Formatting, Data Story telling. 

---

## Project Outcome  
- **Business Impact**: Identified top stores/customers (e.g., 41-50 age group), optimized inventory (low-return products), tracked 46% growth vs. targets.   
- **Scalable Solution**: Handles large datasets dynamically; exportable to Power BI if needed.      

---


