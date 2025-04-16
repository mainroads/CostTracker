let
    Source = SharePoint.Files("https://mainroads.sharepoint.com/teams/MR-30000597-MEBD-PRJ-Commercial", [ApiVersion = 15]),
    ShowXLSX = Table.SelectRows(Source, each ([Extension] = ".xlsx")),
    #"Filtered Rows" = Table.SelectRows(ShowXLSX, each Text.Contains([Name], "PendingTx")),
    #"Filtered Hidden Files1" = Table.SelectRows(#"Filtered Rows", each [Attributes]?[Hidden]? <> true),
    #"Invoke Custom Function1" = Table.AddColumn(#"Filtered Hidden Files1", "Transform File (4)", each #"Transform File (4)"([Content])),
    #"Renamed Columns1" = Table.RenameColumns(#"Invoke Custom Function1", {"Name", "Source.Name"}),
    #"Removed Other Columns1" = Table.SelectColumns(#"Renamed Columns1", {"Source.Name", "Transform File (4)"}),
    #"Expanded Table Column1" = Table.ExpandTableColumn(#"Removed Other Columns1", "Transform File (4)", Table.ColumnNames(#"Transform File (4)"(#"Sample File (4)"))),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded Table Column1",{{"Source.Name", type text}, {"ProjectNumber", Int64.Type}, {"Supplier", type text}, {"Line Description", type text}, {"Amount", Int64.Type}, {"Invoice Number", type text}, {"IncurredDate", type date}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"Source.Name"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns",{{"Line Description", "Tx Description"}, {"ProjectNumber", "Project Number"}}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Renamed Columns",{{"Project Number", type text}, {"Amount", Currency.Type}}),
    #"Renamed Columns2" = Table.RenameColumns(#"Changed Type1",{{"IncurredDate", "Date"}}),
    #"Filtered Rows1" = Table.SelectRows(#"Renamed Columns2", each [Project Number] <> null and [Project Number] <> ""),
    #"Merged Queries" = Table.NestedJoin(#"Filtered Rows1", {"Invoice Number"}, Tx, {"Invoice Number"}, "Tx", JoinKind.LeftAnti),
    #"Removed Columns1" = Table.RemoveColumns(#"Merged Queries",{"Tx"}),
    bufferMeSideways = Table.Buffer(#"Removed Columns1")
in
    bufferMeSideways