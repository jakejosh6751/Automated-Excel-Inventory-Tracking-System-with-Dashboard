# Automated Excel Inventory Tracking System with Dashboard

### About this project
This project is an automated inventory tracking system built in Excel with macros, Power Query, and dashboards. It streamlines the management of product purchases, sales, and stock levels. The system automatically updates KPIs, inventory status, and dashboards when new transactions are entered. It is designed for small to medium-sized businesses to track inventory movement, monitor stock availability, and evaluate sales performance without relying on external software.

### Data Extraction & Transformation
- Products table defines product details such as product ID, name, category, unit cost price, unit sales price, reorder levels, and discontinued status.
- Transactions table records purchases, sales, and other movement types with date, product, quantity, cost, and sales details.
- Data transformation is achieved through formulas, pivot tables, and macros that automatically:
  - Match transactions to products.
  - Update available stock based on purchases and sales.
  - Compute cost of goods sold (COGS) and revenue.
  - Refresh dashboards upon data entry.

### Data Modelling
The system uses a relational model connecting:
- **Products Table** (dimension) → Product attributes.
- **Transactions Table** (fact) → Purchases and sales.
- **KPIs Table** (calculated fact table) → Net sales, inventory value, COGS, turnover, and ratios.
- 
Relationships are maintained via **ProductID** as a key.

### Visualization

### Key Insights

### Recommendation

#### Additional Project Images
