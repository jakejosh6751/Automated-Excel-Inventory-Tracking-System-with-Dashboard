# Automated Excel Inventory Tracking System with Dashboard

### About this project
This project is an automated inventory tracking system built in Excel with VBA, Power Query, Power Pivot, Pivot tables, and dashboard. It streamlines the management of product purchases, sales, and stock levels. The system automatically updates KPIs, inventory status, and dashboard when new transactions are recorded. It is designed for small to medium-sized businesses to track inventory movement, monitor stock availability, and evaluate sales performance without relying on external software.

### Database Modelling
Two *Excel formatted* tables are created using sheets as the database:
- **Products** table defines product details; ProductID *(Primary Key)*, ProductName, Category, UnitCostPrice, UnitSalesPrice, ReorderLevel, and Discontinued *status*.
- **Transactions** table records transaction details; TransactionID *(Primary Key)*, Date, ProductID *(Foreign Key to Products table)*, Quantity, TransactionType (Purchase, Sale, or Damage), and Note *(for remark, especially when TransactionType is Damage)*.

Using VLOOKUP, UnitCostPrice and UnitSalesPrice are copied from the Products table to the Transactions table base on ProductID. This is to ease data modelling calculations. 

*Products* table (snippet):

*Transactions* table (snippet):

### Data Entry

### Data Extraction & Transformation
**Data transformation** is achieved through formulas, pivot tables, and macros that automatically:
  - Match transactions to products.
  - Update available stock based on purchases and sales.
  - Compute cost of goods sold (COGS) and revenue.
  - Refresh dashboards upon data entry.

### Data Modelling
The system uses a relational model connecting:
- **Products Table** (dimension) → Product attributes.
- **Transactions Table** (fact) → Purchases and sales.
- **KPIs Table** (calculated fact table) → Net sales, inventory value, COGS, turnover, and ratios.

Relationships are maintained via **ProductID** as a key.

### Visualization
The project includes an interactive Excel dashboard with the following pages:
1. **Inventory Overview** – Summary of stock availability, inventory value, and movement.
2. **KPI Dashboard** – Tracks key metrics such as Net Sales, COGS, Inventory Turnover Ratio, and Inventory to Sales Ratio.
3. **Top 10 Analysis** –
    - Products by revenue.
    - Products by stock value.
4. **Monthly Sales Trend** – Quantity sold per month.

### Key Insights
From the sample dataset:
- Some products like **Cheddar Crackers** and **Deodorant** generate high sales revenue but also tie up significant stock value.
- Inventory Turnover Ratios range between **1.0 – 1.2**, indicating that stock is moving but at a moderate pace.
- Several products maintain high **available stock relative to sales**, suggesting potential overstocking.

### Recommendation
- **Optimize stock levels** by reducing over-purchased items and prioritising fast-moving products.
- **Reorder point adjustments** should be set based on actual sales velocity instead of fixed thresholds.
- Introduce **damage/loss tracking** for better cost accuracy.
- Consider transitioning the system into **Power BI or SQL-based architecture** for scalability as transaction volume grows.

### Additional Project Images
