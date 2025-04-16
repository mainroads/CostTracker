let
    // 1. Retrieve files from SharePoint
    Source = SharePoint.Files("https://mainroads.sharepoint.com/teams/MR-30000597-MEBD-PRJ-Commercial/", [ApiVersion = 15]),
    #"Filtered Rows" = Table.SelectRows(Source, each ([Extension] = ".xlsx" or [Extension] = ".XLSX")),
    #"Filtered Rows1" = Table.SelectRows(#"Filtered Rows", each Text.Contains([Name], "-FMS_PA_Report07")),
    #"Filtered Hidden Files1" = Table.SelectRows(#"Filtered Rows1", each [Attributes]?[Hidden]? <> true),

    // 2. Import the Excel workbooks
    #"Imported Excel" = Table.AddColumn(#"Filtered Hidden Files1", "Imported Excel", each Excel.Workbook([Content], true)),
    #"Renamed Columns1" = Table.RenameColumns(#"Imported Excel", {{"Name", "Source.Name"}}),

    // 3. Extract the worksheet "FMS_PA_Report07_Resource" from each workbook
    #"Extracted Resource Sheet" = Table.AddColumn(#"Renamed Columns1", "FMS_PA_Report07_Resource", each 
        let
            wb = [Imported Excel],
            resourceTable = Table.SelectRows(wb, each ([Item] = "FMS_PA_Report07_Resource"))
        in
            if Table.IsEmpty(resourceTable) then null else resourceTable{0}[Data]
    ),
    #"Filtered Hidden Files2" = Table.SelectRows(#"Extracted Resource Sheet", each [Attributes]?[Hidden]? <> true),
    
    // 4. Retain only the columns needed: in this case, the source name and the resource sheet data
    #"Kept Necessary Columns" = Table.SelectColumns(#"Filtered Hidden Files2", {"Source.Name", "FMS_PA_Report07_Resource"}),

    // 5. Define a function to transform the Resource sheet table from each file
    TransformResource = (tbl as nullable table) as table =>
        if tbl = null then
            // Return an empty table if the sheet isn’t found
            #table({}, {})
        else
            let
                // Skip the first 8 rows and promote headers
                Skipped = Table.Skip(tbl, 8),
                Promoted = Table.PromoteHeaders(Skipped, [PromoteAllScalars=true]),
                // Define the union of all expected columns across files.
                ExpectedColumns = {
                    "GL Year", "Period", "Date", "Task No", "Task Desc", "Expenditure Type", 
                    "Expenditure Type Desc", "Resource", "Amount", "Quantity", "Expenditure Item ID", 
                    "Orig Transaction Reference", "Line No", "Document", "Transaction Source", 
                    "GL Batch Name", "Comment", "Vendor Name", "Purchase order no", "Agency Specific Contract", 
                    "GLAcct", "Invoice Num"
                },
                // Function to add any missing expected columns as null
                AddMissingColumns = (t as table, expected as list) as table =>
                    let
                        ExistingColumns = Table.ColumnNames(t),
                        MissingColumns = List.Difference(expected, ExistingColumns),
                        TableWithAdded = List.Accumulate(
                            MissingColumns,
                            t,
                            (state, current) => Table.AddColumn(state, current, each null)
                        )
                    in
                        // Reorder to match expected column order
                        Table.ReorderColumns(TableWithAdded, expected),
                Normalized = AddMissingColumns(Promoted, ExpectedColumns)
            in
                Normalized,
                
    // 6. Apply the transformation to each file's resource sheet
    #"Transformed Resource" = Table.AddColumn(#"Kept Necessary Columns", "TransformedResource", each TransformResource([FMS_PA_Report07_Resource])),
    #"Removed Original Sheet" = Table.RemoveColumns(#"Transformed Resource",{"FMS_PA_Report07_Resource"}),
    
    // 7. Expand the normalized resource data.
    // This dynamically expands the columns from the first non-null transformed resource.
    ExpandedResource = Table.ExpandTableColumn(
        #"Removed Original Sheet", 
        "TransformedResource", 
        Table.ColumnNames(#"Removed Original Sheet"{0}[TransformedResource])
    ),
    
    // 8. Continue with further transformations as per your original logic
    #"Renamed Columns2" = Table.RenameColumns(ExpandedResource,{{"Source.Name", "FileName"}}),
    #"Inserted Text Before Delimiter" = Table.AddColumn(#"Renamed Columns2", "Project Number", each Text.BeforeDelimiter([FileName], "-"), type text),
    #"Renamed Columns" = Table.RenameColumns(#"Inserted Text Before Delimiter", {
        {"Task No", "Task Number"},
        {"Task Desc", "Task Name"},
        {"Expenditure Type", "Exp Code"},
        {"Expenditure Type Desc", "Exp Name"}
    }),
    BufferMyTable = Table.Buffer(#"Renamed Columns"),
    #"Inserted Merged Column" = Table.AddColumn(BufferMyTable, "Task Description", each Text.Combine({[Task Number], [Task Name]}, " "), type text),
    #"Inserted Merged Column1" = Table.AddColumn(#"Inserted Merged Column", "Expenditure Description", each Text.Combine({[Exp Code], [Exp Name]}, " "), type text),
    #"Sorted Rows" = Table.Sort(#"Inserted Merged Column1", {{"Date", Order.Descending}}),
    #"Removed Columns" = Table.RemoveColumns(#"Sorted Rows",{"Column7", "Column14"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Removed Columns",{{"Amount", Currency.Type}}),
    #"Removed Columns1" = Table.RemoveColumns(#"Changed Type",{"Column16", "Column18", "Quantity", "Column20", "Orig Transaction Reference", "Column23", "Line No", "Transaction Source", "GL Batch Name", "Agency Specific Contract", "GLAcct"}),
    #"Added Vendor" = Table.AddColumn(#"Removed Columns1", "Vendor", each 
        if [Vendor Name] <> null and Text.Trim([Vendor Name]) <> "" then
            [Vendor Name]
        else if Text.Contains([Comment], "WSP", Comparer.OrdinalIgnoreCase) then
            "WSP Australia Pty Ltd"
        else if Text.Contains([Comment], "BG&E", Comparer.OrdinalIgnoreCase) then
            "BG&E Pty Ltd"
        else if Text.Contains([Comment], "Jacobs", Comparer.OrdinalIgnoreCase) then
            "Jacobs Group (Australia) Pty Ltd"
        else if Text.Contains([Comment], "AAJV", Comparer.OrdinalIgnoreCase) then
            "AAJV"
        else if Text.Contains([Exp Name], "Labour", Comparer.OrdinalIgnoreCase) then
            "Staff Costs (Contractor)"
        else if Text.Contains([Exp Name], "Wages & Salaries", Comparer.OrdinalIgnoreCase) then
            "Staff Costs (Main Roads)"
        else if Text.Contains([Exp Name], "Ins-Principal", Comparer.OrdinalIgnoreCase) then
            "Principal Controlled Insurance Policy"
        else
            [Exp Name]
    ),
    #"Replaced Value1" = Table.ReplaceValue(#"Added Vendor","Randstad Pty Ltd","Staff Costs (Contractor)",Replacer.ReplaceText,{"Vendor"}),
    #"Added IsAccrual" = Table.AddColumn(#"Replaced Value1", "IsAccrual", each try if Text.Contains(Text.Lower([Comment]), "accrual") then true else false otherwise false),
    setIsAccrualToBinary = Table.TransformColumnTypes(#"Added IsAccrual", {{"IsAccrual", type logical}}),
    #"Replaced Value" = Table.ReplaceValue(setIsAccrualToBinary,"Contracts-Infra-Road Infrastructure Construction and Maintenance","Georgiou Group Pty Ltd",Replacer.ReplaceText,{"Vendor"}),
    #"Added isProgressClaim" = Table.AddColumn(#"Replaced Value", "IsProgressClaim", each (
        if [Resource] <> null and Text.Contains([Resource], "021") then true
        else if [Resource] = null and [Vendor] <> null and Text.Contains(Text.Lower([Vendor]), "georgiou") then true
        else false
    ), type logical),
    #"Renamed Columns3" = Table.RenameColumns(#"Added isProgressClaim",{{"Invoice Num", "Invoice Number"}}),
    #"Removed Columns2" = Table.RemoveColumns(#"Renamed Columns3",{"Vendor Name", "Purchase order no"}),
    #"Filtered Rows2" = Table.SelectRows(#"Removed Columns2", each ([GL Year] <> "GL Year")),
    #"Changed Type2" = Table.TransformColumnTypes(#"Filtered Rows2",{{"Project Number", Int64.Type}}),
    #"Filtered Rows3" = Table.SelectRows(#"Changed Type2", each ([Date] <> null and [Date] <> "")),
    #"Changed Type3" = Table.TransformColumnTypes(#"Filtered Rows3",{{"Date", type date}}),
    #"Replaced Errors" = Table.ReplaceErrorValues(#"Changed Type3", {{"Amount", 0}}),
    #"Replaced Value2" = Table.ReplaceValue(#"Replaced Errors","Randstand","Staff Costs (Contractor)",Replacer.ReplaceText,{"Vendor"})
in
    #"Replaced Value2"