# Automated Excel Inventory Tracking System with Dashboard
![Inventory Management Dashboard.jpg](https://github.com/jakejosh6751/Automated-Excel-Inventory-Tracking-System-with-Dashboard/blob/main/Inventory%20Management%20Dashboard.jpg)
*short navigation animation*:

### About this project
This project is an automated inventory tracking system built in Excel with VBA, Power Query, PowerPivot, Pivot tables, and dashboard. It streamlines the management of product purchases, sales, and stock levels. The system automatically updates KPIs, inventory status, and dashboard when new transactions are recorded. It is designed for small to medium-sized businesses to track inventory movement, monitor stock availability, and evaluate sales performance without relying on external software.

### Database Modelling
The system uses a relational model. Two *Excel formatted* tables are created using sheets as the database:
- **Products (Dimension)** table defines product details;
- **Transactions (Facts)** table records transaction details.

Table attributes are described in the *Data Dictionary* below:
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

*Schema Diagram View*:

Key measures were created as described in the metrics dictionary below;

### Visualization
The project includes an interactive Excel dashboard with the following pages:
1. **Inventory Dashboard** –
   - **KPIs** – Tracks key metrics such as Net Sales, COGS, Inventory Turnover Ratio, and Inventory to Sales Ratio.
   - **Top 10 Analysis** –
       - Products by revenue.
       - Products by stock value.
   - **Monthly Sales Trend** – Quantity sold per month.

2. **Inventory Overview** – Summary of stock availability, inventory value, and movement.....stock status







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
