// Power Query (M) — Dim_Customers
// Exported from the workbook. Open via Data > Queries & Connections > right-click Dim_Customers > Edit > Advanced Editor.

let
    Source = Excel.CurrentWorkbook(){[Name="Dim_Customers"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"CustomerID", type text}, {"CustomerName", type text}, {"City", type text}, {"RegistrationDate", type datetime}, {"Segment", type text}, {"PreferredChannel", type text}})
in
    #"Changed Type"
