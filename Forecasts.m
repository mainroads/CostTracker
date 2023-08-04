let
    Source = SharePoint.Files("https://mainroads.sharepoint.com/teams/MR-30000597-MEBD-PRJ-Commercial", [ApiVersion = 15]),
    #"Filtered Rows" = Table.SelectRows(Source, each ([Extension] = ".xlsx" or [Extension] = ".XLSX")),
    #"Filtered Rows1" = Table.SelectRows(#"Filtered Rows", each Text.Contains([Name], "-Forecasts")),
    #"Filtered Hidden Files1" = Table.SelectRows(#"Filtered Rows1", each [Attributes]?[Hidden]? <> true),
    #"Invoke Custom Function1" = Table.AddColumn(#"Filtered Hidden Files1", "Transform File (5)", each #"Transform File (5)"([Content])),
    #"Renamed Columns1" = Table.RenameColumns(#"Invoke Custom Function1", {"Name", "Source.Name"}),
    #"Removed Other Columns1" = Table.SelectColumns(#"Renamed Columns1", {"Source.Name", "Transform File (5)"}),
    #"Expanded Table Column1" = Table.ExpandTableColumn(#"Removed Other Columns1", "Transform File (5)", Table.ColumnNames(#"Transform File (5)"(#"Sample File (3)"))),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded Table Column1",{{"Source.Name", type text}, {"ForecastID", Int64.Type}, {"ProjectNumber", Int64.Type}, {"Date", type date}, {"Amount", type number}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"Source.Name", "ForecastID"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Removed Columns",{{"Amount", Currency.Type}, {"Contingency", Currency.Type}}),
    bufferMeSideways = Table.Buffer(#"Changed Type1"),
    #"Changed Type2" = Table.TransformColumnTypes(bufferMeSideways,{{"Escalation", Currency.Type}, {"Other", Currency.Type}, {"StaffCosts", Currency.Type}})
in
    #"Changed Type2"
