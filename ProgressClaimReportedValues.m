let
    Source = SharePoint.Files("https://mainroads.sharepoint.com/teams/mr-30000597-mebd-con-Commercial", [ApiVersion = 15]),
    #"Filtered Rows" = Table.SelectRows(Source, each Text.StartsWith([Name], "MEBD-C15121-O-CO-PC")),
    #"Filtered Rows1" = Table.SelectRows(#"Filtered Rows", each ([Extension] = ".xlsm" or [Extension] = ".XLSM")),
    #"Filtered Rows2" = Table.SelectRows(#"Filtered Rows1", each not Text.Contains([Name], "Review")),
    #"Filtered Rows3" = Table.SelectRows(#"Filtered Rows2", each ([Name] <> "MEBD-C15121-O-CO-PC-00NN Progress Claim for Works in Month Year.xlsm")),
    #"Sorted Rows" = Table.Sort(#"Filtered Rows3",{{"Folder Path", Order.Ascending}}),
    #"Filtered Hidden Files1" = Table.SelectRows(#"Sorted Rows", each [Attributes]?[Hidden]? <> true),
    // Directly extract and promote headers from 'Payment Calcs' sheet in each file
    #"Get Payment Calcs Sheet" = Table.AddColumn(#"Filtered Hidden Files1", "ExtractedData", each let
        wb = Excel.Workbook([Content], null, true),
        sheet = wb{[Item="Payment Calcs", Kind="Sheet"]}[Data],
        promoted = Table.PromoteHeaders(sheet, [PromoteAllScalars=true])
    in promoted),
    #"Renamed Columns1" = Table.RenameColumns(#"Get Payment Calcs Sheet", {"Name", "Source.Name"}),
    #"Removed Other Columns1" = Table.SelectColumns(#"Renamed Columns1", {"Source.Name", "ExtractedData"}),
    #"Expanded Table Column1" = Table.ExpandTableColumn(#"Removed Other Columns1", "ExtractedData", Table.ColumnNames(#"Get Payment Calcs Sheet"{0}[ExtractedData]), Table.ColumnNames(#"Get Payment Calcs Sheet"{0}[ExtractedData])),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded Table Column1",{{"Source.Name", type text}, {"MANDURAH ESTUARY BRIDGE DUPLICATION (MEBD) PROJECT", type any}, {"Column2", type text}, {"Column3", type any}, {"Column4", type any}, {"Column5", type any}, {"Column6", type any}, {"Column7", type any}, {"Column8", type any}, {"Column9", type any}, {"PAYMENT CERTIFICATE No. 1", type text}, {"Column11", type number}, {"Column12", type any}, {"Column13", type any}, {"Column14", type any}, {"Column15", type text}, {"Column16", type text}, {"Column17", type text}, {"Column18", type text}, {"Column19", type any}, {"Column20", type text}}),
    #"Filtered Rows7" = Table.SelectRows(#"Changed Type", each ([Column2] = "TOTAL (Excludes GST)")),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows7",{"PAYMENT CERTIFICATE No. 1", "Column11", "Column12", "Column13", "Column14", "Column15", "Column16", "Column17", "Column18", "Column19", "Column20", "MANDURAH ESTUARY BRIDGE DUPLICATION (MEBD) PROJECT"}),
    #"Filtered Rows4" = Table.SelectRows(#"Removed Columns", each ([Column8] <> null)),
    #"Filtered Rows5" = Table.SelectRows(#"Filtered Rows4", each true),
    #"Removed Columns2" = Table.RemoveColumns(#"Filtered Rows5",{"Column2"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns2",{{"Column3", "Revised Contract Sum"}}),
    #"Changed Type3" = Table.TransformColumnTypes(#"Renamed Columns",{{"Revised Contract Sum", Currency.Type}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type3", "Start Date", each let
    SourceText = [Source.Name],
    MonthText = Text.BetweenDelimiters(SourceText, "in ", " "),
    YearText = Text.Middle(SourceText, Text.PositionOf(SourceText, MonthText) + Text.Length(MonthText) + 1, 4),
    MonthNumber = Date.Month(Date.FromText("1 " & MonthText & " " & YearText))
  in
    #date(Number.FromText(YearText), MonthNumber, 1)),
    #"Changed Type1" = Table.TransformColumnTypes(#"Added Custom",{{"Start Date", type date}}),
    #"Added Custom1" = Table.AddColumn(#"Changed Type1", "Finish Date", each Date.EndOfMonth([Start Date])),
    #"Changed Type2" = Table.TransformColumnTypes(#"Added Custom1",{{"Finish Date", type date}}),
    #"Added Custom2" = Table.AddColumn(#"Changed Type2", "PC Number", each let
    SourceText = [Source.Name],
    PCText = Text.BetweenDelimiters(SourceText, "PC-", " "),
    PCNumber = "PC" & Text.End(PCText, 2)
  in
    PCNumber),
    #"Removed Columns1" = Table.RemoveColumns(#"Added Custom2",{"Column4", "Column5"}),
    #"Renamed Columns2" = Table.RenameColumns(#"Removed Columns1",{{"Column6", "Cumulative Progress"}}),
    #"Changed Type4" = Table.TransformColumnTypes(#"Renamed Columns2",{{"Cumulative Progress", Percentage.Type}}),
    #"Removed Columns3" = Table.RemoveColumns(#"Changed Type4",{"Column7", "Column8"}),
    #"Renamed Columns3" = Table.RenameColumns(#"Removed Columns3",{{"Column9", "Cumulative Claim Value"}}),
    #"Changed Type5" = Table.TransformColumnTypes(#"Renamed Columns3",{{"Cumulative Claim Value", Currency.Type}}),
    #"Inserted Text Before Delimiter" = Table.AddColumn(#"Changed Type5", "Text Before Delimiter", each Text.BeforeDelimiter([Source.Name], " "), type text),
    #"Renamed Columns4" = Table.RenameColumns(#"Inserted Text Before Delimiter",{{"Text Before Delimiter", "PC ID"}})
in
    #"Renamed Columns4"