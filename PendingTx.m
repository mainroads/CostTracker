let
    Source = SharePoint.Files("https://mainroads.sharepoint.com/teams/MR-30000597-MEBD-PRJ-Commercial", [ApiVersion = 15]),
    ShowXLSX = Table.SelectRows(Source, each ([Extension] = ".xlsx")),
    #"Filtered Rows" = Table.SelectRows(ShowXLSX, each Text.Contains([Name], "PendingTx")),
    #"Filtered Hidden Files1" = Table.SelectRows(#"Filtered Rows", each [Attributes]?[Hidden]? <> true),
    #"Filtered Rows2" = Table.SelectRows(#"Filtered Hidden Files1", each not Text.Contains([Folder Path], "/ss")),
    
    // For each file, extract the first sheet
    #"Added Custom" = Table.AddColumn(#"Filtered Rows2", "Excel Sheets", each Excel.Workbook([Content])),
    #"Expanded Excel Sheets" = Table.ExpandTableColumn(#"Added Custom", "Excel Sheets", {"Name", "Data", "Kind"}, {"SheetName", "Data", "Kind"}),
    #"Filtered Sheet Rows" = Table.SelectRows(#"Expanded Excel Sheets", each ([Kind] = "Sheet")),
    
    // Group by file name and take the first sheet for each
    #"Grouped Rows" = Table.Group(#"Filtered Sheet Rows", {"Name"}, {
        {"First Sheet", each Table.FirstN(_, 1), type table [Name=text, SheetName=text, Data=table, Kind=text]}
    }),
    
    // Expand the first sheet data
    #"Expanded First Sheet" = Table.ExpandTableColumn(#"Grouped Rows", "First Sheet", {"SheetName", "Data"}, {"SheetName", "Data"}),
    #"Expanded Data" = Table.ExpandTableColumn(#"Expanded First Sheet", "Data", Table.ColumnNames(#"Expanded First Sheet"[Data]{0})),
    
    // Ensure we have the headers in the first row
    #"Promoted Headers" = Table.PromoteHeaders(#"Expanded Data", [PromoteAllScalars=true]),
    
    // Transform data types to match original code
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers", {
        {"ProjectNumber", Int64.Type}, 
        {"Supplier", type text}, 
        {"Line Description", type text}, 
        {"Amount", Int64.Type}, 
        {"Invoice Number", type text}, 
        {"IncurredDate", type date}
    }),
    
    // Apply the same column transformations as in the original code
    #"Renamed Columns" = Table.RenameColumns(#"Changed Type", {
        {"Line Description", "Tx Description"}, 
        {"ProjectNumber", "Project Number"}
    }),
    #"Changed Type1" = Table.TransformColumnTypes(#"Renamed Columns", {
        {"Project Number", type text}, 
        {"Amount", Currency.Type}
    }),
    #"Renamed Columns2" = Table.RenameColumns(#"Changed Type1", {{"IncurredDate", "Date"}}),
    #"Filtered Rows1" = Table.SelectRows(#"Renamed Columns2", each [Project Number] <> null and [Project Number] <> ""),
    #"Merged Queries" = Table.NestedJoin(#"Filtered Rows1", {"Invoice Number"}, Tx, {"Invoice Number"}, "Tx", JoinKind.LeftAnti),
    #"Removed Columns1" = Table.RemoveColumns(#"Merged Queries", {"Tx"}),
    bufferMeSideways = Table.Buffer(#"Removed Columns1"),
    columnsToRemove = List.Intersect({Table.ColumnNames(bufferMeSideways), {"30000597-PendingTx.xlsx", "PendingTx"}}),
    #"Removed Columns" = Table.RemoveColumns(bufferMeSideways, columnsToRemove)
in
    #"Removed Columns"