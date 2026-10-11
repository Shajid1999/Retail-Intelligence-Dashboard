// Power Query (M) — Fact_Sales
// Exported from the workbook. Open via Data > Queries & Connections > right-click Fact_Sales > Edit > Advanced Editor.

let
    Source = Table.Combine({Clean_InStore, Clean_Online}),
    #"Removed Other Columns" = Table.SelectColumns(Source,{"SKU", "Qty", "StoreID", "UnitPrice", "DiscountPct", "Date", "Amount", "Channel", "CustomerID"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Removed Other Columns",{{"Amount", type number}, {"UnitPrice", type number}, {"DiscountPct", type number}, {"Qty", Int64.Type}, {"Date", type date}})
in
    #"Changed Type"
