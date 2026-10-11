# Build Guide: Recreate This Project From Scratch

This guide walks through building the dashboard step by step in Excel. Follow it to rebuild the project yourself or to understand how each part works.

**You need:** Excel for Microsoft 365 or Excel 2016+ on Windows, with **Power Pivot** enabled (File → Options → Add-ins → COM Add-ins → *Microsoft Power Pivot for Excel*).

---

## Phase 1: Prepare the raw data

The workbook starts with three raw exports that are deliberately messy, plus the reference tables:

| Sheet | Content | Excel table name |
|---|---|---|
| Raw_InStore_Sales | POS transactions | `tbl_Raw_InStore` |
| Raw_Online_Sales | Online order lines | `tbl_Raw_Online` |
| Raw_Inventory_Snapshot | Weekly stock, one column per store (wide) | `tbl_Raw_Inventory` |
| Dim_Products / Dim_Stores / Dim_Customers / Dim_Calendar | Reference tables | same as sheet name |
| Demand_Forecast | Monthly forecast vs actual | `Fact_Forecast` |

Convert each range to a table (**Ctrl + T**) and give it the name above (Table Design → Table Name).

---

## Phase 2: Clean and shape with Power Query

Load each table through **Data → From Table/Range**. The full M code for every query is in [`/power-query`](../power-query).

### Query 1: Clean_InStore ([code](../power-query/01_Clean_InStore.pq))
Problems in the raw data and how each one is fixed:

| Problem | Fix |
|---|---|
| Blank / garbage rows | Filter out rows with no `TransactionID` |
| Stray spaces and hidden characters in `StoreCode`, `SKU` | `Text.Trim` + `Text.Clean` |
| Lower-case store codes (`shw-cha09`) | `Text.Upper` |
| Four different date formats (`d/M/yyyy`, `M/d/yyyy`, `d-MMM-yy`, `yyyy-MM-dd`) | Custom column that tries each format in turn |
| Duplicate rows | Remove Duplicates |
| Missing discount | Replace `null` with `0` |
| Inconsistent payment labels (`cash`, `CASH`, `bKash`, `MFS`…) | Standardise to `Cash`, `Card`, `Mobile Banking` |
| No revenue column | Add `Amount = Qty × UnitPrice × (1 − DiscountPct)` |
| Missing channel flag | Add `Channel = "InStore"` |

Rename `StoreCode` → `StoreID`. **Result: 1,300 clean rows.**

### Query 2: Clean_Online ([code](../power-query/02_Clean_Online.pq))

| Problem | Fix |
|---|---|
| `UnitPrice` stored as text with a currency symbol (`৳400.20`) | Keep only digits and `.`, convert to number |
| Mixed date-time formats (ISO `2024-03-30T…` and `30-03-2024 4:45 PM`) | Parse the first 10 characters with the right format |
| Payment method variants (`bkash`, `BKASH`, `cod`, `COD`…) | Standardise to `bKash`, `Nagad`, `Cash on Delivery`, `Card` |
| District and delivery status casing / spaces | `Text.Proper` + `Text.Trim` |
| Missing `CustomerID` | Replace with `GUEST` |
| Discount stored as a whole number (10 = 10%) | Divide by 100 |
| No physical store | Set `StoreID = "Online"` (matches the `ONLINE` row in Dim_Stores) |

Add `Amount` and `Channel = "Online"`. **Result: 1,050 clean rows.**

### Query 3: Fact_Sales ([code](../power-query/03_Fact_Sales.pq))
**Append** Clean_InStore + Clean_Online (Home → Append Queries), keep only the 9 shared columns, and set data types. **Result: 2,350 rows.**

### Query 4: Fact_Inventory ([code](../power-query/04_Fact_Inventory.pq))
The raw snapshot has one column per store (`SHW-DHA01`, `SHW-DHA02`…).
1. Select `Category`, `SKU`, `ProductName`, `WeekEnding`, `ReorderLevel` → **Transform → Unpivot Other Columns**.
2. Rename `Attribute` → `StoreID`, `Value` → `StockQty`, `WeekEnding` → `Date`.
3. Remove `Category` and `ProductName` (they come from Dim_Products).

**Result: 8,634 rows** in long format.

### Queries 5–9: dimensions and forecast
`Dim_Products`, `Dim_Stores`, `Dim_Customers`, `Dim_Calendar` and `Fact_Forecast` only need correct data types.

**Load settings:** Home → Close & Load To… → **Only Create Connection** + tick **Add this data to the Data Model** for the 7 model tables. Clean_InStore and Clean_Online can also be loaded to sheets so you can inspect them.

---

## Phase 3: Build the Data Model (Power Pivot)

Open **Power Pivot → Manage → Diagram View** and drag the key from each fact table onto its dimension:

![Data model](../images/data_model.png)

| From (many) | To (one) | Key |
|---|---|---|
| Fact_Sales | Dim_Products / Dim_Stores / Dim_Customers / Dim_Calendar | SKU / StoreID / CustomerID / Date |
| Fact_Inventory | Dim_Products / Dim_Stores / Dim_Calendar | SKU / StoreID / Date |
| Fact_Forecast | Dim_Stores / Dim_Calendar | StoreID / MonthDate → Date |

Then mark the calendar: select Dim_Calendar → **Design → Mark as Date Table** → `Date`.

> **Why Fact_Forecast links to Dim_Calendar (not a separate month table):** forecasts are stored on the 1st of each month. Linking to the same daily calendar as sales means both share one month axis and one set of date filters, which the *Sales vs Forecast* chart relies on. It's a valid one-to-many link because `MonthDate` repeats only on the fact side.

---

## Phase 4: Write the DAX measures

In the Power Pivot calculation area (below each table), add the measures from [`/dax/measures.dax`](../dax/measures.dax):

- **Sales:** Total Sales, Total Units Sold, Transaction Count, AOV
- **Inventory:** Total Stock Qty, Avg Weekly Stock, Current Stock, Reorder Point, Stock Health Ratio
- **Forecast:** Total Forecast Qty, Total Actual Qty, Forecast Variance

---

## Phase 5: Calculation engine (`Pivot_Engine` sheet)

Insert → PivotTable → **From Data Model**, once for each visual and KPI:

| Pivot | Rows | Values |
|---|---|---|
| KPI block (cells B4:B7) | — | Total Sales, Total Units Sold, Transaction Count, AOV |
| Monthly trend | Dim_Calendar[MonthName] | Total Sales, Total Forecast Qty |
| Channel split | Fact_Sales[Channel] | Total Sales |
| Top districts | Dim_Stores[District] (Top 5 filter) | Total Sales |
| Top categories | Dim_Products[Category] | Total Sales |

Format the KPI cells explicitly, because linked text boxes inherit the cell's format:
- Revenue: `"৳ "#,##0.00,,"M"`
- AOV: `"৳ "#,##0.00`
- Counts: `#,##0`

---

## Phase 6: Dashboard layout (`Dashboard` sheet)

1. **Canvas:** View → untick *Gridlines* and *Headings*.
2. **Colour palette:** Crimson `#C8102E` (primary), Dark Slate `#1E293B` (secondary), Light Slate `#E2E8F0` (borders), White `#FFFFFF` (cards).
3. **Layout (F-pattern):** a slim control panel on the left, KPI cards across the top, charts in the centre, and insights at the bottom right.
4. **KPI cards:** Insert → Shapes → Rounded Rectangle. Add a text box, click into the formula bar, and type `=Pivot_Engine!$B$4` (B5, B6 and B7 for the others).
5. **Charts:** create PivotCharts from the engine pivots:
   - Monthly Sales vs Forecast: combo chart, sales as columns, forecast as a line on the secondary axis
   - Channel split: donut chart
   - Top 5 districts and top categories: horizontal bar charts
6. **Slicers:** Insert → Slicer for **Channel**, **District** and **Category**. For each one, right-click → **Report Connections** → tick **every** pivot, so one click filters the whole dashboard.

---

## Phase 7: Inventory report (`Inventory_Reporting` sheet)

A PivotTable from the Data Model with `Dim_Products[SKU]`, `[ProductName]` and `[Category]` on rows, and **Current Stock**, **Reorder Point** and **Stock Health Ratio** as values. Use conditional formatting on the ratio: anything **below 1.0** needs replenishment.

---

## Phase 8: Final checks

- **Data → Refresh All** runs without errors.
- Every slicer changes every chart and KPI.
- The KPI cards show: **৳ 3.56M**, **13,034** units, **2,350** transactions, **৳ 1,515.65** AOV (all filters cleared).
