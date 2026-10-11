// Power Query (M) — Dim_Calendar
// Exported from the workbook. Open via Data > Queries & Connections > right-click Dim_Calendar > Edit > Advanced Editor.

let
    Source = Excel.CurrentWorkbook(){[Name="Dim_Calendar"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Date", type date}, {"Year", Int64.Type}, {"Quarter", type text}, {"MonthNum", Int64.Type}, {"MonthName", type text}, {"YearMonth", type date}, {"Day", Int64.Type}, {"WeekdayName", type text}, {"IsWeekend", type text}, {"WeekOfYear", Int64.Type}})
in
    #"Changed Type"
