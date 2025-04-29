let
    // 1. Connect to SharePoint
    Source = SharePoint.Files("https://mainroads.sharepoint.com/teams/mr-30000597-mebd-con-Commercial/", [ApiVersion = 15]),

    // 2. Filter for relevant files (ends with PPR, .xlsx extension, case-insensitive)
    #"Filtered Files" = Table.SelectRows(Source, each
        (Text.EndsWith(Text.Lower([Name]), "ppr.xlsx"))
    ),

    // 3. Filter out hidden files
    #"Filtered Hidden Files" = Table.SelectRows(#"Filtered Files", each [Attributes]?[Hidden]? <> true),

    // 4. Function to extract the first sheet from an Excel file content, promoting headers
    GetFirstSheet = (fileContent as binary) as table =>
        let
            // Get workbook structure, use headers, infer types
            Workbook = Excel.Workbook(fileContent, true, true),
            // Get the first row (representing the first sheet/table)
            FirstSheetInfo = Table.FirstN(Workbook, 1){0}?,
            // Extract the data table if the first sheet exists
            SheetData = if FirstSheetInfo = null then null else FirstSheetInfo[Data]
        in
            SheetData,

    // 5. Add a column with the data from the first sheet of each file
    #"Added First Sheet Data" = Table.AddColumn(#"Filtered Hidden Files", "FirstSheetData", each GetFirstSheet([Content]), type table),

    // 6. Remove rows where the first sheet couldn't be extracted (e.g., empty/corrupted file or no sheets)
    #"Filtered Empty Sheets" = Table.SelectRows(#"Added First Sheet Data", each [FirstSheetData] <> null and not Table.IsEmpty([FirstSheetData])),

    // 7. Add Source Filename to each table before combining
    #"Added Source Name" = Table.AddColumn(#"Filtered Empty Sheets", "DataWithSource", each Table.AddColumn([FirstSheetData], "Source.Name", (innerRecord) => [Name])),

    // 8. Combine all tables
    // Check if there are tables to combine to prevent errors
    CombinedData = if Table.IsEmpty(#"Added Source Name") then
                       // Create an empty table with at least the Source.Name column if no files found/processed
                       #table({"Source.Name"}, {})
                   else
                       Table.Combine(#"Added Source Name"[DataWithSource]),

    // 9. Optional: Change type of Source.Name if needed (adjust other types as necessary based on data)
    #"Changed Type Source Name" = Table.TransformColumnTypes(CombinedData,{{"Source.Name", type text}}),

    // Existing steps (adjusting input table from #"Replaced WSP Name" to #"Changed Type Source Name")
    #"Removed Top Rows" = Table.Skip(#"Changed Type Source Name",7),
    #"Promoted Headers" = Table.PromoteHeaders(#"Removed Top Rows", [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Subcontractor name#(lf)(list all engaged PBA subcontractors)", type text}, {"ABN", type any}, {"BSB", type text}, {"Account Number", type any}, {"PPI payment this cycle #(lf)($)", type any}, {"PPI retention this cycle #(lf)($)", type any}, {"Contractor Deposit Instruction payments this cycle#(lf)($)", type any}, {"Contractor Deposit Instruction retention this cycle #(lf)($)", type any}, {"Total subcontractor payments to Date ($)", type any}, {"Total retention held to Date#(lf)($)", type any}, {"Total retention released to Subcontractor to date#(lf)($)", type any}, {"Total retention paid to head contractor for liabilities to date ($)", type any}, {"Failed to claim this cycle? (Y/N)", type text}, {"Disputes / Rights set off ($)", type any}, {"Description / Comments#(lf)(Describe any payment anomalies, oustanding payments, issues etc)", type text}, {"Column16", type any}, {"Column17", type any}, {"Column18", type any}, {"MEBD-C15121-O-CO-PC-0002 PPR.xlsx", type text}, {"Column20", type any}}),
    #"Renamed Columns" = Table.RenameColumns(#"Changed Type",{{"Subcontractor name#(lf)(list all engaged PBA subcontractors)", "Subcontractor Name"}}),

    // MOVED Step & Updated Logic: Standardize Subcontractor Names
    #"Standardized Subcontractor Names" = Table.TransformColumns(#"Renamed Columns", {
        {"Subcontractor Name", each 
            let
                currentValue = _
            in
                if currentValue is text then
                    if Text.Contains(currentValue, "WSP", Comparer.OrdinalIgnoreCase) then
                        "WSP Australia Pty Ltd"
                    else if Text.Contains(currentValue, "BG & E", Comparer.OrdinalIgnoreCase) then
                        "BG&E Pty Ltd"
                    else
                        currentValue // No change
                else
                    currentValue, // Not text, no change
            type text
        }
    }),

    // Remaining steps (adjusting input table)
    #"Filtered Rows" = Table.SelectRows(#"Standardized Subcontractor Names", each ([ABN] <> null and [ABN] <> "151.21 Mandurah Estuary Bridge Duplication" and [ABN] <> "ABN") and ([Subcontractor Name] <> null and [Subcontractor Name] <> "Date of payment cert:" and [Subcontractor Name] <> "Date of payment claim:" and [Subcontractor Name] <> "Date of this report:" and [Subcontractor Name] <> "Scheduled payment date:")),
    #"Changed Type1" = Table.TransformColumnTypes(#"Filtered Rows",{{"ABN", type text}, {"BSB", type text}, {"Account Number", type text}, {"PPI payment this cycle #(lf)($)", Currency.Type}, {"PPI retention this cycle #(lf)($)", Currency.Type}, {"Contractor Deposit Instruction payments this cycle#(lf)($)", Currency.Type}, {"Contractor Deposit Instruction retention this cycle #(lf)($)", Currency.Type}, {"Total subcontractor payments to Date ($)", Currency.Type}, {"Total retention held to Date#(lf)($)", Currency.Type}, {"Total retention released to Subcontractor to date#(lf)($)", Currency.Type}, {"Total retention paid to head contractor for liabilities to date ($)", Currency.Type}, {"Disputes / Rights set off ($)", Currency.Type}}),
    #"Renamed Columns1" = Table.RenameColumns(#"Changed Type1",{{"Description / Comments#(lf)(Describe any payment anomalies, oustanding payments, issues etc)", "Comments"}}),
    #"Removed Columns" = Table.RemoveColumns(#"Renamed Columns1",{"Column16", "Column17", "Column18", "Column20"}),
    #"Inserted Text Before Delimiter" = Table.AddColumn(#"Removed Columns", "Text Before Delimiter", each Text.BeforeDelimiter([#"MEBD-C15121-O-CO-PC-0002 PPR.xlsx"], " "), type text),
    #"Renamed Columns2" = Table.RenameColumns(#"Inserted Text Before Delimiter",{{"Text Before Delimiter", "PC ID"}}),
    #"Removed Columns1" = Table.RemoveColumns(#"Renamed Columns2",{"MEBD-C15121-O-CO-PC-0002 PPR.xlsx"}),
    #"Merged Queries" = Table.NestedJoin(#"Removed Columns1", {"PC ID"}, #"Progress Claim Reported Values", {"PC ID"}, "Progress Claim Reported Values", JoinKind.LeftOuter),
    #"Expanded Progress Claim Reported Values" = Table.ExpandTableColumn(#"Merged Queries", "Progress Claim Reported Values", {"Start Date", "Finish Date"}, {"Start Date", "Finish Date"}),
    #"Renamed Columns3" = Table.RenameColumns(#"Expanded Progress Claim Reported Values",{{"PPI payment this cycle #(lf)($)", "PPI payment this cycle ($)"}, {"PPI retention this cycle #(lf)($)", "PPI retention this cycle ($)"}, {"Contractor Deposit Instruction payments this cycle#(lf)($)", "Contractor Deposit Instruction payments this cycle ($)"}, {"Contractor Deposit Instruction retention this cycle #(lf)($)", "Contractor Deposit Instruction retention this cycle ($)"}}),
    // Add the custom column at the end using the renamed columns
    #"Added Total Payment This Cycle" = Table.AddColumn(
        #"Renamed Columns3", 
        "Total Payment This Cycle ($)", 
        each ([#"PPI payment this cycle ($)"] ?? 0) + ([#"Contractor Deposit Instruction payments this cycle ($)"] ?? 0), 
        Currency.Type
    )
in
    #"Added Total Payment This Cycle" // Output the final table with the new custom column