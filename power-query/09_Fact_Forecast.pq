// Power Query (M) — Fact_Forecast
// Exported from the workbook. Open via Data > Queries & Connections > right-click Fact_Forecast > Edit > Advanced Editor.

let
    Source = Excel.CurrentWorkbook(){[Name="Fact_Forecast"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"StoreID", type text}, {"MonthDate", type datetime}, {"YearMonth", type date}, {"ForecastQty", Int64.Type}, {"ActualQty", Int64.Type}})
in
    #"Changed Type"
