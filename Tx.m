let
    // Connect to SharePoint
    Source = SharePoint.Files("https://mainroads.sharepoint.com/teams/MR-30000597-MEBD-PRJ-Commercial/", [ApiVersion = 15]),
    #"Filtered Rows" = Table.SelectRows(Source, each ([Extension] = ".xlsx" or [Extension] = ".XLSX")),
    #"Filtered Rows1" = Table.SelectRows(#"Filtered Rows", each Text.Contains([Name], "-FMS_PA_Report07")),
    #"Filtered Hidden Files1" = Table.SelectRows(#"Filtered Rows1", each [Attributes]?[Hidden]? <> true),

    // Import the Excel content for each file
    #"Imported Excel" = Table.AddColumn(#"Filtered Hidden Files1", "Imported Excel", each Excel.Workbook([Content], true)),
    #"Renamed Columns1" = Table.RenameColumns(#"Imported Excel", {{"Name", "Source.Name"}}),

    // Extract the table named "FMS_PA_Report07_Resource" from each workbook
    #"Filtered Resource" = Table.AddColumn(#"Renamed Columns1", "FMS_PA_Report07_Resource", each 
        let
            wb = [Imported Excel],
            resourceTable = Table.SelectRows(wb, each ([Item] = "FMS_PA_Report07_Resource"))
        in
            if Table.IsEmpty(resourceTable) then null else resourceTable{0}[Data]
    ),
    #"Filtered Hidden Files2" = Table.SelectRows(#"Filtered Resource", each [Attributes]?[Hidden]? <> true),

    // Invoke your custom transform function (kept as-is)
    #"Invoke Custom Function1" = Table.AddColumn(#"Filtered Hidden Files2", "Transform File (12)", each #"Transform File (12)"([Content])),
    #"Removed Other Columns1" = Table.SelectColumns(#"Invoke Custom Function1",{"Source.Name", "Transform File (12)"}),
    #"Expanded Table Column1" = Table.ExpandTableColumn(#"Removed Other Columns1", "Transform File (12)", Table.ColumnNames(#"Transform File (12)"(#"Sample File (8)"))),
    #"Changed Type1" = Table.TransformColumnTypes(#"Expanded Table Column1",{
        {"Column1", type any}, {"Column2", type any}, {"Column3", type text}, {"Column4", type text}, {"Column5", type text}, {"Column6", type any}, {"Column7", type any}, {"Column8", type text}, {"Project Detail Transactions", type any}, {"Column10", type text}, {"Column11", type any}, {"Column12", type text}, {"Column13", type text}, {"Column14", type text}, {"Column15", type text}, {"Column16", type any}, {"Column17", type text}, {"Column18", type any}, {"Column19", type text}, {"Column20", type text}, {"Column21", type text}, {"Column22", type any}, {"Column23", type any}, {"Column24", type text}, {"Column25", type text}, {"Column26", type text}, {"Column27", type text}, {"Column28", type text}, {"Column29", type text}, {"Column30", type any}, {"Column31", type any}, {"Column32", type text}, {"Column33", type text}, {"Column34", type text}, {"Column35", type text}, {"Column36", type text}
    }),
    #"Removed Top Rows1" = Table.Skip(#"Changed Type1",8),
    #"Promoted Headers" = Table.PromoteHeaders(#"Removed Top Rows1", [PromoteAllScalars=true]),

    // **Define the expected union of column names and a function to add missing ones dynamically**
    ExpectedColumns = {
        "GL Year", "Period", "Date", "Task No", "Task Desc", "Expenditure Type", 
        "Expenditure Type Desc", "Resource", "Amount", "Quantity", "Expenditure Item ID", 
        "Orig Transaction Reference", "Line No", "Document", "Transaction Source", 
        "GL Batch Name", "Comment", "Vendor Name", "Purchase order no", "Agency Specific Contract", 
        "GLAcct", "Invoice Num"
    },
    AddMissingColumns = (tbl as table, expected as list) as table =>
        let
            ExistingColumns = Table.ColumnNames(tbl),
            MissingColumns = List.Difference(expected, ExistingColumns),
            TableWithAdded = List.Accumulate(
                MissingColumns,
                tbl,
                (state, current) => Table.AddColumn(state, current, each null)
            )
        in
            // Reorder the table to have exactly the expected columns order.
            Table.ReorderColumns(TableWithAdded, expected),

    // **Normalize the promoted table by adding any missing columns**
    #"Normalized Table" = AddMissingColumns(#"Promoted Headers", ExpectedColumns),

    // Continue with your transformations using the normalized table
    #"Renamed Columns2" = Table.RenameColumns(#"Normalized Table",{{"30000597-FMS_PA_Report07.xlsx", "FileName"}}),
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
    #"Removed Columns" = Table.RemoveColumns(#"Sorted Rows",{"Column4", "Column7", "Column8", "Column10", "Column14"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Removed Columns",{{"Amount", Currency.Type}}),
    #"Removed Columns1" = Table.RemoveColumns(#"Changed Type",{"Column12", "Column16", "Column18", "Quantity", "Column20", "Orig Transaction Reference", "Column22", "Column23", "Line No", "Column25", "Transaction Source", "Column27", "GL Batch Name", "Column29", "Column31", "Column32", "Agency Specific Contract", "GLAcct"}),
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
    #"Added isProgressClaim" = Table.AddColumn(setIsAccrualToBinary, "IsProgressClaim", each ([Resource] <> null and Text.Contains([Resource], "021")), type logical),
    #"Renamed Columns3" = Table.RenameColumns(#"Added isProgressClaim",{{"Invoice Num", "Invoice Number"}}),
    #"Removed Columns2" = Table.RemoveColumns(#"Renamed Columns3",{"Vendor Name", "Purchase order no"}),
    #"Filtered Rows2" = Table.SelectRows(#"Removed Columns2", each ([GL Year] <> "GL Year")),
    #"Changed Type2" = Table.TransformColumnTypes(#"Filtered Rows2",{{"Project Number", Int64.Type}}),
    #"Filtered Rows3" = Table.SelectRows(#"Changed Type2", each ([Date] <> null and [Date] <> "")),
    #"Changed Type3" = Table.TransformColumnTypes(#"Filtered Rows3",{{"Date", type date}}),
    #"Replaced Errors" = Table.ReplaceErrorValues(#"Changed Type3", {{"Amount", 0}}),
    #"Replaced Value2" = Table.ReplaceValue(#"Replaced Errors","Randstand","Staff Costs (Contractor)",Replacer.ReplaceText,{"Vendor"}),
    #"Replaced Value" = Table.ReplaceValue(#"Replaced Value2","Contracts-Infra-Road Infrastructure Construction and Maintenance","Georgiou Group Pty Ltd",Replacer.ReplaceText,{"Vendor"})
in
    #"Replaced Value"
