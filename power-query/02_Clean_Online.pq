// Power Query (M) — Clean_Online
// Exported from the workbook. Open via Data > Queries & Connections > right-click Clean_Online > Edit > Advanced Editor.

let
    Source = Excel.CurrentWorkbook(){[Name="tbl_Raw_Online"]}[Content],
    #"Replaced Value" = Table.TransformColumns(Source, {{"UnitPrice", each Number.From(Text.Select(Text.From(_), {"0".."9", "."})), type number}}),
    #"Added Custom" = Table.AddColumn(#"Replaced Value", "Date_Clean", each let
    raw = Text.Trim(Text.From([OrderDateTime])),
    datePart = Text.Start(raw, 10)
in
    if Text.Contains(raw, "T") then
        Date.FromText(datePart, [Format="yyyy-MM-dd"])
    else
        Date.FromText(datePart, [Format="dd-MM-yyyy"])),
    #"Replaced Value1" = Table.ReplaceValue(#"Added Custom","bkash","bKash",Replacer.ReplaceText,{"PaymentMethod"}),
    #"Replaced Value2" = Table.ReplaceValue(#"Replaced Value1","BKASH","bKash",Replacer.ReplaceText,{"PaymentMethod"}),
    #"Replaced Value3" = Table.ReplaceValue(#"Replaced Value2","nagad","Nagad",Replacer.ReplaceText,{"PaymentMethod"}),
    #"Replaced Value4" = Table.ReplaceValue(#"Replaced Value3","COD","Cash on Delivery",Replacer.ReplaceText,{"PaymentMethod"}),
    #"Replaced Value5" = Table.ReplaceValue(#"Replaced Value4","cod","Cash on Delivery",Replacer.ReplaceText,{"PaymentMethod"}),
    #"Replaced Value6" = Table.ReplaceValue(#"Replaced Value5","card","Card",Replacer.ReplaceText,{"PaymentMethod"}),
    #"Capitalized Each Word" = Table.TransformColumns(#"Replaced Value6",{{"DeliveryDistrict", Text.Proper, type text}}),
    #"Trimmed Text" = Table.TransformColumns(#"Capitalized Each Word",{{"DeliveryDistrict", Text.Trim, type text}}),
    #"Capitalized Each Word1" = Table.TransformColumns(#"Trimmed Text",{{"DeliveryStatus", Text.Proper, type text}}),
    #"Replaced Value7" = Table.ReplaceValue(#"Capitalized Each Word1",null,"GUEST",Replacer.ReplaceValue,{"CustomerID"}),
    #"Removed Duplicates" = Table.Distinct(#"Replaced Value7"),
    #"Added Custom1" = Table.AddColumn(#"Removed Duplicates", "Amount", each [Qty] * [UnitPrice] * (1 - ([DiscountPct] / 100))),
    #"Added Custom2" = Table.AddColumn(#"Added Custom1", "Channel", each "Online"),
    #"Divided Column" = Table.TransformColumns(#"Added Custom2", {{"DiscountPct", each _ / 100, type number}}),
    #"Added Custom3" = Table.AddColumn(#"Divided Column", "StoreID", each null),
    #"Removed Columns" = Table.RemoveColumns(#"Added Custom3",{"OrderDateTime"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns",{{"Date_Clean", "Date"}}),
    #"Replaced Value8" = Table.ReplaceValue(#"Renamed Columns",null,"",Replacer.ReplaceValue,{"StoreID"}),
    #"Replaced Value9" = Table.ReplaceValue(#"Replaced Value8","","Online",Replacer.ReplaceValue,{"StoreID"})
in
    #"Replaced Value9"
