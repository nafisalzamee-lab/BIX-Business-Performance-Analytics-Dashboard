#Excel Business Performance Analytics Dashboard  


## Project Overview  
This dashboard transforms raw transactional data into actionable insights through a **normalized data model** in **Power Pivot**, featuring multiple interconnected tables (fact table with 20,000+ rows, customer, product, sales, store, and custom date dimensions).   
---
**Problem Statement**  
The business required an interactive dashboard to consolidate and analyze transactional data for **revenue tracking**, **profit optimization**, **customer segmentation**, and **operational KPIs** across stores, timeframes, and demographics—but lacked a scalable Excel solution.   

---

## Project Objectives  
- Transform raw transactional data into an interactive dashboard for **business decision-making**.  
- Enable dynamic slicing by time, customer, product, store, and salesperson to identify **growth opportunities** and **underperformers**.   
- Provide visual KPIs with targets/variances to track **revenue growth (46%)**, **profit margins (42.8%)**, and **refund rates (8%)**.   

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
1. **Revenue & Profit Trends**: Monthly/quarterly revenue vs. target, MoM growth, weekday/weekend splits (73% weekday revenue).   
2. **Store Performance**: Revenue/profit by location with variance % and top/bottom rankings.   
3. **Customer Segmentation**: Profit by age group (41-50 peak), gender, top/bottom 5 customers.   
4. **Product Insights**: Top products by profit/quantity, category distribution, return rates (<10% target).   
5. **Operational KPIs**: Refund rate (8%), # transactions/customers/products, avg. order value.   
---

## Data Processing & Modeling  

### Key Cleaning Steps (Power Query)  
1. **Import & Promote Headers**: Loaded CSVs, promoted first row to headers.  
2. **Data Type Detection**: Auto-changed columns (e.g., Date to Date, Quantity to Whole Number).  
3. **Merge Names**: Appended FirstName + " " + LastName → FullName for customers/salespersons.  
4. **Age Calculation**: $$age = \frac{TODAY() - DOB}{365.25}$$ (rounded).  
5. **Remove Duplicates**: Eliminated based on TransactionID.  
6. **Date Table Creation**: Calendar table with extracted Year/Month/Quarter/Weekday; added sorting columns (MonthNumber, DayNumber).   
7. **Load Strategy**: Tables as "Connection Only" except calculations table.  


---

### Data Preparation (Power Query)  
Structured ETL pipeline to clean and transform raw CSV data into a relational model.  

1. **Source Import**: Loaded 5 CSV files (transactions, customers, products, salespersons, stores) via **Get Data > From Text/CSV**.  
2. **Header Promotion**: Promoted first row to column headers across all queries.  
3. **Data Type Changes**: Auto-detected and set types (e.g., Date to Date, QuantitySold/ReturnQuantity to Whole Number, IDs to Text).  
4. **Name Merging**: Combined FirstName + " " + LastName into **FullName** columns for customers and salespersons.  
5. **Age Derivation**: Added **Age** column: $$Age = \frac{TODAY() - DOB}{365.25}$$ (rounded to nearest whole number).  
6. **Duplicate Removal**: Eliminated rows based on unique **TransactionID** in fact table.  
7. **Date Dimension Creation**: Generated custom calendar table from min/max dates:
   - Extracted **Year**, **Month Name**, **Month Number**, **Quarter**, **Weekday**, **Weekend** flag.
   - Added sorting columns (e.g., MonthSort = 1-12).   
8. **Load Optimization**: Queries set to **Connection Only** (except calculations table) for Power Pivot efficiency.  

> "Power Query handles the heavy lifting of data transformation before loading into Power Pivot."  

 

### 2.Establishing  Power Pivot Relationships   
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
 

### 3. DAX Measures  
**Core Calculations** (stored in blank "calculations" table):  
1. **Total Revenue**: $$SUMX(FactTable, FactTable[QuantitySold] \times RELATED(ProductTable[SalesPrice]))$$  
2. **COGS**: Similar SUMX with cost price.  
3. **Profit Margin**: $$[Total Revenue] - [COGS]$$  
4. **Profit %**: $$DIVIDE([Profit Margin], [Total Revenue])$$  
5. **# Transactions**: $$COUNTROWS(FactTable)$$  
6. **Total Refund**: $$SUMX(FactTable, FactTable[ReturnQuantity] \times RELATED(ProductTable[SalesPrice]))$$  
7. **Refund Rate**: $$DIVIDE([Total Refund], [Total Revenue])$$  
8. **# Products Sold**: $$DISTINCTCOUNT(FactTable[ProductID])$$  
9. **Return Rate**: $$DIVIDE(SUM(FactTable[ReturnQuantity]), SUM(FactTable[QuantitySold]))$$  
10. **# Customers**: $$DISTINCTCOUNT(FactTable[CustomerID])$$  

Custom formats: e.g., $$[<=1000000]\$0.0,,"M";[<=1000]\$0.0,"K";\$0.0$$ for abbreviated millions/Ks.  

---

## Dashboard Visualizations  

### Store Dashboard (Part 1)  
- Revenue vs. target bars (Zebra add-in) by store with variance % and arrows (IF formulas + TEXT for $$+12.0\% \uparrow$$).   
- Month slicer with custom formatting (no borders, white fill, bold selected).  

### Time Frame Dashboard (Part 2)  
- Revenue/target trends (smoothed lines + markers) with top-2 highlights (LARGE function).   
- Variance waterfall (invert if negative, conditional colors).  
- Waffle charts: Revenue by weekday/weekend (1-100 grid + conditional formatting/icons).   
- Quarter revenue vs. average line + MoM % change (overlaps, data labels from cells).   

### Profit View Dashboard (Part 3)  
- Switchable top/bottom 5: Customers/locations by profit (option buttons + group box, IF/XLOOKUP).   
- Age group profit bars + average line.   
- Gender waffle (SEQUENCE/SORT for 10x10 grid, icons, conditional rules).   
- Top products: Profit/quantity switch (group box), category pie.   
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
**Key Features:**
- **Dynamic KPIs**: Total revenue ($$\$5.4M$$), COGS, profit margin (42.8%), refund rate (8%), and targets with variance indicators (e.g., +46% growth).   
- **Multi-tab Views**: Store analysis, time frame trends, profit views with customer/product breakdowns.   
- **Interactive Controls**: Slicers (month, category), option buttons (top/bottom 5), combo boxes, and hide/show filters via VBA macros.   
- **Custom Visuals**: Waffle charts (weekday revenue split), gradient-filled trends, zebra add-in bars, and conditional formatting with icons/arrows.   
-**Data Model**: Star schema with 1:many relationships (fact → dimensions).  
-**DAX Measures**: 10+ custom metrics (e.g., Profit % = $$DIVIDE([Profit], [Revenue])$$, formatted as %).     
-**Advanced Charts**: Waffle grids (SEQUENCE/conditional formatting), gradient trends, zebra bars (add-in), pie for categories.   
-**Dynamic Elements**: Captions (TEXTJOIN/IF), variance labels ($$TEXT(variance,"+0.0\% \u2191")$$), hyperlinks.   

---

## Key Insights & Business Value  
- **Performance Trends**: Q2/Q4 exceed average revenue; weekdays drive 73% sales.   
- **Customer Focus**: 41-50 age group most profitable (avg. age 45); top locations for targeted marketing.  
- **Product Optimization**: Low refund (8%), focus inventory on top profitable items (e.g., via return rate <10%).   
- **Scalability**: Handles large datasets; fully dynamic for real-time filtering.

**Skills Demonstrated**: Power Query ETL, DAX modeling, advanced charting (waffle/gradients), form controls/VBA, custom formatting. 
---



## Project Outcome  
- **Business Impact**: Identified top stores/customers (e.g., 41-50 age group), optimized inventory (low-return products), tracked 46% growth vs. targets.   
- **Scalable Solution**: Handles large datasets dynamically; exportable to Power BI if needed.  
- **Portfolio Value**: Demonstrates end-to-end analytics (ETL → viz) in Excel. Built via Data with Decision's 3-part tutorial series.    

---


