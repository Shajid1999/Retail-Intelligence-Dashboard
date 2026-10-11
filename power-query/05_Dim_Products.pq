// Power Query (M) — Dim_Products
// Exported from the workbook. Open via Data > Queries & Connections > right-click Dim_Products > Edit > Advanced Editor.

let
    Source = Excel.CurrentWorkbook(){[Name="Dim_Products"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"SKU", type text}, {"ProductName", type text}, {"Category", type text}, {"SubCategory", type text}, {"Brand", type text}, {"PrivateLabel", type text}, {"PackSize", type text}, {"UnitCost", type number}, {"UnitPrice", type number}})
in
    #"Changed Type"
