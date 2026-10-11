// Power Query (M) — Clean_InStore
// Exported from the workbook. Open via Data > Queries & Connections > right-click Clean_InStore > Edit > Advanced Editor.

let
    // 1. Load Raw Source
    Source = Excel.CurrentWorkbook(){[Name="tbl_Raw_InStore"]}[Content],

    // 2. Remove null/garbage rows
    FilteredRows = Table.SelectRows(Source, each [TransactionID] <> null and [TransactionID] <> ""),

    // 3. Trim and uppercase text columns
    TrimmedText = Table.TransformColumns(FilteredRows,{{"StoreCode", Text.Trim, type text}, {"SKU", Text.Trim, type text}}),
    CleanedText = Table.TransformColumns(TrimmedText,{{"StoreCode", Text.Clean, type text}, {"SKU", Text.Clean, type text}}),
    UppercasedText = Table.TransformColumns(CleanedText,{{"StoreCode", Text.Upper, type text}}),

    // 4. Create Date_IsAssumed directly from raw text (exactly 264 TRUEs)
    AddedDateAssumed = Table.AddColumn(UppercasedText, "Date_IsAssumed", each 
        let
            raw = Text.Trim(Text.From([Date])),
            hasSlash = Text.Contains(raw, "/")
        in
            if hasSlash then
                let
                    parts = Text.Split(raw, "/"),
                    p1 = try Number.FromText(parts{0}) otherwise 99,
                    p2 = try Number.FromText(parts{1}) otherwise 99
                in
                    p1 <= 12 and p2 <= 12
            else 
                false, 
        type logical
    ),

    // 5. Parse into a clean, unified Date_Clean column
    AddedDateClean = Table.AddColumn(AddedDateAssumed, "Date_Clean", each 
        let
            raw = Text.Trim(Text.From([Date]))
        in
            if Text.Contains(raw, "/") then
                try Date.FromText(raw, [Format="d/M/yyyy"])
                otherwise try Date.FromText(raw, [Format="M/d/yyyy"])
                otherwise null
            else if Text.Contains(raw, "-") and Value.Is(try Number.FromText(Text.Middle(raw, 3, 3)) otherwise null, type null) then
                try Date.FromText(raw, [Format="d-MMM-yy", Culture="en-US"])
                otherwise try Date.FromText(raw, [Format="yyyy-MM-dd"])
                otherwise null
            else
                try Date.FromText(raw, [Format="yyyy-MM-dd"])
                otherwise try Date.From(raw)
                otherwise null,
        type date
    ),
    #"Removed Duplicates" = Table.Distinct(AddedDateClean),
    #"Replaced Value" = Table.ReplaceValue(#"Removed Duplicates",null,0,Replacer.ReplaceValue,{"DiscountPct"}),
    #"Replaced Value1" = Table.ReplaceValue(#"Replaced Value","cash","Cash",Replacer.ReplaceText,{"PaymentType"}),
    #"Replaced Value2" = Table.ReplaceValue(#"Replaced Value1","CASH","Cash",Replacer.ReplaceText,{"PaymentType"}),
    #"Replaced Value3" = Table.ReplaceValue(#"Replaced Value2","mobile banking","Mobile Banking",Replacer.ReplaceText,{"PaymentType"}),
    #"Replaced Value4" = Table.ReplaceValue(#"Replaced Value3","bKash","Mobile Banking",Replacer.ReplaceText,{"PaymentType"}),
    #"Replaced Value5" = Table.ReplaceValue(#"Replaced Value4","MFS","Mobile Banking",Replacer.ReplaceText,{"PaymentType"}),
    #"Replaced Value6" = Table.ReplaceValue(#"Replaced Value5","card","Card",Replacer.ReplaceText,{"PaymentType"}),
    #"Added Custom" = Table.AddColumn(#"Replaced Value6", "Amount", each [Qty] * [UnitPrice] * (1 - [DiscountPct])),
    #"Added Custom1" = Table.AddColumn(#"Added Custom", "Channel", each "InStore"),
    #"Removed Columns" = Table.RemoveColumns(#"Added Custom1",{"Date", "ProductName_asPOS", "Date_IsAssumed"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns",{{"Date_Clean", "Date"}, {"StoreCode", "StoreID"}})
in
    #"Renamed Columns"
