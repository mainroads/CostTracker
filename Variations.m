let
    Source = SharePoint.Files("https://mainroads.sharepoint.com/teams/mr-30000597-mebd-con-Commercial/", [ApiVersion = 15]),
    // Filter for relevant files
    #"Filtered Rows" = Table.SelectRows(Source, each Text.StartsWith([Name], "MEBD-C15121-O-CO-REG-0001")),
    #"Filtered Hidden Files1" = Table.SelectRows(#"Filtered Rows", each [Attributes]?[Hidden]? <> true),
    // Get file content as tables and combine them
    #"Imported Tables" = Table.AddColumn(#"Filtered Hidden Files1", "Table", each Excel.Workbook([Content])),
    #"Expanded Table" = Table.ExpandTableColumn(#"Imported Tables", "Table", {"Item", "Data"}),
    #"Filtered PV Sheets" = Table.SelectRows(#"Expanded Table", each [Item] = "PV"),
    // Expand only the 'Data' column, keep 'Name' for later
    #"Expanded Data" = Table.TransformColumns(#"Filtered PV Sheets", {"Data", each Table.PromoteHeaders(_)}),
    #"Appended Data" = Table.AddColumn(#"Expanded Data", "AppendedTable", each Table.AddColumn([Data], "Source.Name", (x) => [Name])),
    #"Combined Tables" = Table.Combine(#"Appended Data"[AppendedTable]),
    #"Changed Type" = Table.TransformColumnTypes(#"Combined Tables",{{"Source.Name", type text}, {"Date Raised", type date}, {"Potential Variation #", type text}, {"Description", type text}, {"Closed", type text}, {"Quote value", type any}, {"Contrator Ref #", type text}, {"DoA Approval TRIM #", type text}, {"Variation #", type text}, {"Variation Letter TRIM #", type text}, {"Next Action", type text}, {"Action with", type text}, {"Revised Contract Sum", type number}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"Source.Name"}),
    #"Filtered Rows1" = Table.SelectRows(#"Removed Columns", each ([Date Raised] <> null)),
    #"Renamed Columns" = Table.RenameColumns(#"Filtered Rows1",{{"Potential Variation #", "PV ID"}, {"Variation Letter TRIM #", "TRIM ID"}}),
    #"Prepended PV ID to Description" = Table.AddColumn(#"Renamed Columns", "Description2", each [PV ID] & " " & [Description], type text),
    #"Removed Old Description" = Table.RemoveColumns(#"Prepended PV ID to Description", {"Description"}),
    #"Renamed Description2" = Table.RenameColumns(#"Removed Old Description", {{"Description2", "Description"}}),
    #"Changed Type2" = Table.TransformColumnTypes(#"Renamed Description2",{{"Revised Contract Sum", Currency.Type}}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Changed Type2",{{"Quote value", Currency.Type}}),
    #"Replaced Errors" = Table.ReplaceErrorValues(#"Changed Type1", {{"Quote value", null}}),
    #"Renamed Columns2" = Table.RenameColumns(#"Replaced Errors",{{"Variation #", "Variation ID"}}),
    #"Added Custom" = Table.AddColumn(#"Renamed Columns2", "Status", each if Text.StartsWith([Variation ID], "V") then "Approved" else if Text.StartsWith([Variation ID], "TBA") then "Pending" else "Unapproved"),
    #"Removed Columns1" = Table.RemoveColumns(#"Added Custom",{"Revised Contract Sum"}),
    #"Replaced Errors1" = Table.ReplaceErrorValues(#"Removed Columns1", {{"Status", "Unapproved"}})
in
    #"Replaced Errors1"