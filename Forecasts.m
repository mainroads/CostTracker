let
    Source = SharePoint.Files("https://mainroads.sharepoint.com/teams/MR-30000597-MEBD-PRJ-Commercial", [ApiVersion = 15]),
    #"Filtered Rows" = Table.SelectRows(Source, each ([Extension] = ".xlsx" or [Extension] = ".XLSX")),
    #"Filtered Rows1" = Table.SelectRows(#"Filtered Rows", each Text.Contains([Name], "-Forecasts")),
    #"Filtered Hidden Files1" = Table.SelectRows(#"Filtered Rows1", each [Attributes]?[Hidden]? <> true),
    
    // Import Excel workbooks and extract the first sheet from each
    #"Added Excel Data" = Table.AddColumn(#"Filtered Hidden Files1", "Excel Data", each Excel.Workbook([Content])),
    #"Expanded Excel Data" = Table.ExpandTableColumn(#"Added Excel Data", "Excel Data", {"Name", "Data", "Kind"}, {"Sheet Name", "Sheet Data", "Kind"}),
    #"Filtered Sheets" = Table.SelectRows(#"Expanded Excel Data", each ([Kind] = "Sheet")),
    #"Filtered Rows2" = Table.SelectRows(#"Filtered Sheets", each not Text.Contains([Folder Path], "/ss")),
    
    // Group by file and take the first sheet from each file
    #"Grouped by File" = Table.Group(#"Filtered Rows2", {"Name"}, {
        {"First Sheet", each Table.FirstN(_, 1), type table}
    }),
    #"Expanded First Sheet" = Table.ExpandTableColumn(#"Grouped by File", "First Sheet", {"Sheet Data"}, {"Sheet Data"}),
    
    // Add a step to normalize each sheet's columns before combining
    #"Added Normalized Data" = Table.AddColumn(#"Expanded First Sheet", "Normalized Data", each 
        let
            sheetData = [Sheet Data],
            promotedHeaders = Table.PromoteHeaders(sheetData, [PromoteAllScalars=true]),
            // Define the union of all expected columns
            allExpectedColumns = {
                "ForecastID", "ProjectNumber", "Date", "Amount", "ManagementCosts", 
                "Other", "Other2", "Escalation", "Contingency", "Progress Claim", 
                "isActual"
            },
            existingColumns = Table.ColumnNames(promotedHeaders),
            missingColumns = List.Difference(allExpectedColumns, existingColumns),
            // Add missing columns with null values
            tableWithMissingCols = List.Accumulate(
                missingColumns,
                promotedHeaders,
                (state, current) => Table.AddColumn(state, current, each null)
            ),
            // Reorder columns to match expected order
            reorderedTable = Table.ReorderColumns(tableWithMissingCols, allExpectedColumns)
        in
            reorderedTable
    ),
    
    // Remove the original Sheet Data column and expand the normalized data
    #"Removed Sheet Data" = Table.RemoveColumns(#"Added Normalized Data", {"Sheet Data"}),
    #"Expanded Normalized Data" = Table.ExpandTableColumn(#"Removed Sheet Data", "Normalized Data", 
        {"ForecastID", "ProjectNumber", "Date", "Amount", "ManagementCosts", 
         "Other", "Other2", "Escalation", "Contingency", "Progress Claim", 
         "isActual"}),
    

    bufferMeSideways = Table.Buffer(#"Expanded Normalized Data"),
    #"Changed Type" = Table.TransformColumnTypes(bufferMeSideways,{{"Name", type text}, {"ForecastID", type any}, {"ProjectNumber", type any}, {"Date", type any}, {"Amount", type any}, {"ManagementCosts", type any}, {"Other", type any}, {"Other2", type any}, {"Escalation", type any}, {"Contingency", type any}, {"Progress Claim", type any}, {"isActual", type text}}),
    #"Filtered Rows3" = Table.SelectRows(#"Changed Type", each ([ForecastID] <> "ForecastID")),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows3",{"Progress Claim", "isActual"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Removed Columns",{{"ForecastID", Int64.Type}, {"ProjectNumber", Int64.Type}, {"Date", type date}, {"Amount", Currency.Type}, {"ManagementCosts", Currency.Type}, {"Other", Currency.Type}, {"Other2", Currency.Type}, {"Escalation", Currency.Type}, {"Contingency", Currency.Type}})
in
    #"Changed Type1"