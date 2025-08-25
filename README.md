# Automated Excel Inventory Tracking System with Dashboard
![Inventory Management Dashboard.jpg](https://github.com/jakejosh6751/Automated-Excel-Inventory-Tracking-System-with-Dashboard/blob/main/Inventory%20Management%20Dashboard.jpg)
*short navigation animation*:

### About this project
This project is an automated inventory tracking system built in Excel with VBA, Power Query, PowerPivot, Pivot tables, and dashboard. It streamlines the management of product purchases, sales, and stock levels. The system automatically updates KPIs, inventory status, and dashboard when new transactions are recorded. It is designed for small to medium-sized businesses to track inventory movement, monitor stock availability, and evaluate sales performance without relying on external software.

### Database Modelling
The system uses a relational model. Two *Excel formatted* tables are created using sheets as the database:
- **Products (Dim)** defines product details;
- **Transactions (Facts)** records transaction details;
- **Calendar (Dim)** holds daily calendar dates.

*Data Dictionary*:
![Data Dictionary.png](https://github.com/jakejosh6751/Automated-Excel-Inventory-Tracking-System-with-Dashboard/blob/main/Data%20Dictionary.png)

Using **VLOOKUP**, UnitCostPrice and UnitSalesPrice are copied from the Products table to the Transactions table base on ProductID. This is to ease data modelling calculations.

### Data Entry
The system features an **Excel UserForm** that simplifies recording new transactions into the **Transactions** table.

*Data Entry UserForm*:
![Data Entry UserForm.png](https://github.com/jakejosh6751/Automated-Excel-Inventory-Tracking-System-with-Dashboard/blob/main/Data%20Entry%20UserForm.png)

To reduce errors and speed up entry, key fields are automated and streamlined:
- **Product Name**: Selected from a drop-down list of available products.
- **Quantity**: Defaults to 1 in a numeric box, with spin buttons to adjust up or down.
- **Transaction Type**: Defaults to “Sale” (the most common transaction), with other options available such as Purchase or Damage.
- **Date**: Automatically set to the current date through a VBA formula, with the option to override if necessary.
- **Note**: An optional text field, mainly used for recording remarks on damaged items.

This user-friendly entry system ensures consistency, accuracy, and efficiency in updating inventory records.

### Data Extraction & Transformation
Dataset from the Products and Transactions table are extracted into Power Query. Data types are adjusted.

A **Calendar** table is created in Power Query using **List.Dates** to cover the desired duration. In this case, *List.Dates(#date(2025, 1, 1), 365, #duration(1, 0, 0, 0))*. Additional columns *(Month, Month Name, Week of Year)* are added to enable date filters for months and weeks in visuals when needed.

All three tables *(Products, Transactions, and Calendar)* are loaded to the Data Model (PowerPivot).

### Data Modelling
The data model adopts a star schema structure (the most efficient for a data model) comprising a single fact table (Transactions) related to two dimension tables (Products & Calendar). Relationships are maintained via **ProductID** (Primary Key for the Products table) and **Date** (Primary Key for the Calendar table).

*Schema Diagram*:
![Schema Diagram.png](https://github.com/jakejosh6751/Automated-Excel-Inventory-Tracking-System-with-Dashboard/blob/main/Schema%20Diagram.png)

*Key Measures*:
![Key Metrics.png](https://github.com/jakejosh6751/Automated-Excel-Inventory-Tracking-System-with-Dashboard/blob/main/Key%20Metrics.png)

### Visualization
The project features an **Inventory Overview Sheet** highlighting key metrics with colour formatting to emphasize important insights, and an interactive **Inventory Dashboard** with slicers (filters) to explore data by months, product categories, and products.

*Inventory Overview Sheet*:
![Inventory Overview.png](https://github.com/jakejosh6751/Automated-Excel-Inventory-Tracking-System-with-Dashboard/blob/main/Inventory%20Overview.png)

*Inventory Dashboard*:
![Inventory Management Dashboard.jpg](https://github.com/jakejosh6751/Automated-Excel-Inventory-Tracking-System-with-Dashboard/blob/main/Inventory%20Management%20Dashboard.jpg)

### Key Insights
**1.** Overall Inventory & Sales Performance
    - The business currently holds 334 units in stock worth ₦336,500.
    - Across the period, 237 sales transactions generated ₦953,710 in revenue.
    - The inventory turnover ratio is 4.95, showing that stock is being cycled moderately — about five times in the period.
2. Sales Trends & Seasonality
    - Sales quantities reveal clear seasonal patterns.
      - Demand peaks in April (189 units) and June (186 units).
      - Lowest activity occurred in February (74 units).
    - This suggests Q2 (Apr–Jun) is a high-demand period that requires stronger inventory readiness.







### Recommendation
- **Optimize stock levels** by reducing over-purchased items and prioritising fast-moving products.
- **Reorder point adjustments** should be set based on actual sales velocity instead of fixed thresholds.
- Introduce **damage/loss tracking** for better cost accuracy.
- Consider transitioning the system into **Power BI or SQL-based architecture** for scalability as transaction volume grows.
