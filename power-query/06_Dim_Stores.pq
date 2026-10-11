// Power Query (M) — Dim_Stores
// Exported from the workbook. Open via Data > Queries & Connections > right-click Dim_Stores > Edit > Advanced Editor.

let
    Source = Excel.CurrentWorkbook(){[Name="Dim_Stores"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"StoreID", type text}, {"StoreName", type text}, {"District", type text}, {"Region", type text}, {"Format", type text}, {"SizeSqft", Int64.Type}, {"OpenDate", type datetime}})
in
    #"Changed Type"
