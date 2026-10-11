// Power Query (M) — Fact_Inventory
// Exported from the workbook. Open via Data > Queries & Connections > right-click Fact_Inventory > Edit > Advanced Editor.

let
    Source = Excel.CurrentWorkbook(){[Name="tbl_Raw_Inventory"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Category", type text}, {"SKU", type text}, {"ProductName", type text}, {"WeekEnding", type date}, {"ReorderLevel", Int64.Type}, {"SHW-DHA01", Int64.Type}, {"SHW-DHA02", Int64.Type}, {"SHW-DHA03", Int64.Type}, {"SHW-DHA04", Int64.Type}, {"SHW-GAZ05", Int64.Type}, {"SHW-GAZ06", Int64.Type}, {"SHW-NAR07", Int64.Type}, {"SHW-NAR08", Int64.Type}, {"SHW-CHA09", Int64.Type}, {"SHW-CHA10", Int64.Type}, {"SHW-CHA11", Int64.Type}, {"SHW-CUM12", Int64.Type}, {"SHW-SYL13", Int64.Type}, {"SHW-KHU14", Int64.Type}, {"SHW-RAJ15", Int64.Type}}),
    #"Unpivoted Other Columns" = Table.UnpivotOtherColumns(#"Changed Type", {"Category", "SKU", "ReorderLevel", "ProductName", "WeekEnding"}, "Attribute", "Value"),
    #"Renamed Columns" = Table.RenameColumns(#"Unpivoted Other Columns",{{"Attribute", "StoreID"}, {"Value", "StockQty"}, {"WeekEnding", "Date"}}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Renamed Columns",{{"StockQty", Int64.Type}, {"ReorderLevel", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type1",{"Category", "ProductName"})
in
    #"Removed Columns"
