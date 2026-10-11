# Data Dictionary

All data is a synthetic 2024 dataset for a multi-store grocery retailer with 15 physical stores across 8 districts plus an online store.

## Fact tables

### Fact_Sales — 2,350 rows
One row per in-store transaction or online order. Built by appending `Clean_InStore` and `Clean_Online` in Power Query.

| Column | Type | Description |
|---|---|---|
| SKU | Text | Product key → `Dim_Products[SKU]` |
| Qty | Whole number | Units sold |
| StoreID | Text | Store key → `Dim_Stores[StoreID]`. Online orders use `ONLINE` |
| UnitPrice | Decimal | Selling price per unit (৳) |
| DiscountPct | Decimal | Discount as a fraction (0.10 = 10%) |
| Date | Date | Transaction date → `Dim_Calendar[Date]` |
| Amount | Decimal | Net revenue = `Qty × UnitPrice × (1 − DiscountPct)` |
| Channel | Text | `InStore` or `Online` |
| CustomerID | Text | Customer key → `Dim_Customers[CustomerID]` (online only; blank for in-store) |

### Fact_Inventory — 8,634 rows
Weekly stock-on-hand snapshot per store and SKU (8 weeks × 15 stores × 72 SKUs), unpivoted from a wide layout.

| Column | Type | Description |
|---|---|---|
| SKU | Text | → `Dim_Products[SKU]` |
| Date | Date | Week-ending date of the snapshot → `Dim_Calendar[Date]` |
| ReorderLevel | Whole number | Stock level at which the SKU should be reordered |
| StoreID | Text | → `Dim_Stores[StoreID]` |
| StockQty | Whole number | Units on hand at week end |

### Fact_Forecast — 180 rows
Monthly forecast vs. actual demand per store (15 stores × 12 months).

| Column | Type | Description |
|---|---|---|
| StoreID | Text | → `Dim_Stores[StoreID]` |
| MonthDate | Date | First day of the month → `Dim_Calendar[Date]` |
| YearMonth | Text | e.g. `2024-01` |
| ForecastQty | Whole number | Forecast units |
| ActualQty | Whole number | Actual units sold |

## Dimension tables

### Dim_Products — 72 rows
`SKU` (key), `ProductName`, `Category` (8 categories), `SubCategory`, `Brand`, `PrivateLabel`, `PackSize`, `UnitCost`, `UnitPrice`

### Dim_Stores — 16 rows
`StoreID` (key), `StoreName`, `District`, `Region`, `Format` (Supermarket / Express / E-commerce), `SizeSqft`, `OpenDate`.
15 physical stores plus 1 `ONLINE` record, which was added so online sales link to a real store row (see README → Challenges).

### Dim_Customers — 200 rows
`CustomerID` (key), `CustomerName`, `City`, `RegistrationDate`, `Segment`, `PreferredChannel`

### Dim_Calendar — 366 rows
One row per day of 2024: `Date` (key), `Year`, `Quarter`, `MonthNum`, `MonthName`, `YearMonth`, `Day`, `WeekdayName`, `IsWeekend`, `WeekOfYear`

## Relationships (all one-to-many, single direction)

| Fact (many side) | Dimension (one side) | Key |
|---|---|---|
| Fact_Sales | Dim_Products | SKU |
| Fact_Sales | Dim_Stores | StoreID |
| Fact_Sales | Dim_Customers | CustomerID |
| Fact_Sales | Dim_Calendar | Date |
| Fact_Inventory | Dim_Products | SKU |
| Fact_Inventory | Dim_Stores | StoreID |
| Fact_Inventory | Dim_Calendar | Date |
| Fact_Forecast | Dim_Stores | StoreID |
| Fact_Forecast | Dim_Calendar | MonthDate → Date |
