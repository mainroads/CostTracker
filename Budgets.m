let
    Source = SharePoint.Files("https://mainroads.sharepoint.com/teams/MR-30000597-MEBD-PRJ-Commercial/", [ApiVersion = 15]),
    #"Filtered Rows" = Table.SelectRows(Source, each Text.Contains([Name], "-BudgetData")),
    #"Filtered Hidden Files1" = Table.SelectRows(#"Filtered Rows", each [Attributes]?[Hidden]? <> true),
    #"Invoke Custom Function1" = Table.AddColumn(#"Filtered Hidden Files1", "Transform File (6)", each #"Transform File (6)"([Content])),
    #"Renamed Columns1" = Table.RenameColumns(#"Invoke Custom Function1", {"Name", "Source.Name"}),
    #"Removed Other Columns1" = Table.SelectColumns(#"Renamed Columns1", {"Source.Name", "Transform File (6)"}),
    #"Expanded Table Column1" = Table.ExpandTableColumn(#"Removed Other Columns1", "Transform File (6)", Table.ColumnNames(#"Transform File (6)"(#"Sample File (5)"))),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded Table Column1",{{"Source.Name", type text}, {"ProjectNumber", Int64.Type}, {"FinancialYear", type text}, {"BudgetValue", type number}, {"Column4", type any}, {"Column5", type any}, {"Column6", type any}, {"Column7", type any}, {"Column8", type any}, {"Column9", type any}, {"Column10", type any}, {"Column11", type any}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"Column4", "Column5", "Column6", "Column7", "Column8", "Column9", "Column10", "Column11", "Source.Name"}),
    #"Removed Blank Rows" = Table.SelectRows(#"Removed Columns", each not List.IsEmpty(List.RemoveMatchingItems(Record.FieldValues(_), {"", null}))),
    #"Changed Type1" = Table.TransformColumnTypes(#"Removed Blank Rows",{{"ProjectNumber", type text}}),
    #"Inserted Text Before Delimiter" = Table.AddColumn(#"Changed Type1", "Text Before Delimiter", each Text.Combine({"1/7/", Text.Start([FinancialYear], 4)}), type text),
    #"Changed Type3" = Table.TransformColumnTypes(#"Inserted Text Before Delimiter",{{"Text Before Delimiter", type date}}),
    #"Renamed Columns" = Table.RenameColumns(#"Changed Type3",{{"Text Before Delimiter", "Date"}}),
    #"Changed Type2" = Table.TransformColumnTypes(#"Renamed Columns",{{"BudgetValue", Currency.Type}}),
    bufferMeSideways = Table.Buffer(#"Changed Type2")
in
    bufferMeSideways
