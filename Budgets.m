let
    Source = SharePoint.Files("https://mainroads.sharepoint.com/teams/MR-30000597-MEBD-PRJ-Commercial/", [ApiVersion = 15]),
    #"Filtered Rows" = Table.SelectRows(Source, each Text.Contains([Name], "-BudgetData")),
    #"Filtered Hidden Files1" = Table.SelectRows(#"Filtered Rows", each [Attributes]?[Hidden]? <> true),
    
    // For each file, extract the first sheet
    #"Added Custom" = Table.AddColumn(#"Filtered Hidden Files1", "Excel Sheets", each Excel.Workbook([Content])),
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
        {"FinancialYear", type text}, 
        {"BudgetValue", type number}
    }),
    
    // Remove any extra columns that might be present in the Excel file
    #"Selected Columns" = Table.SelectColumns(#"Changed Type", {"ProjectNumber", "FinancialYear", "BudgetValue"}),
    #"Removed Blank Rows" = Table.SelectRows(#"Selected Columns", each not List.IsEmpty(List.RemoveMatchingItems(Record.FieldValues(_), {"", null}))),
    #"Changed Type1" = Table.TransformColumnTypes(#"Removed Blank Rows", {{"ProjectNumber", type text}}),
    #"Inserted Text Before Delimiter" = Table.AddColumn(#"Changed Type1", "Text Before Delimiter", each Text.Combine({"1/7/", Text.Start([FinancialYear], 4)}), type text),
    #"Changed Type3" = Table.TransformColumnTypes(#"Inserted Text Before Delimiter", {{"Text Before Delimiter", type date}}),
    #"Renamed Columns" = Table.RenameColumns(#"Changed Type3", {{"Text Before Delimiter", "Date"}}),
    #"Changed Type2" = Table.TransformColumnTypes(#"Renamed Columns", {{"BudgetValue", Currency.Type}}),
    bufferMeSideways = Table.Buffer(#"Changed Type2")
in
    bufferMeSideways