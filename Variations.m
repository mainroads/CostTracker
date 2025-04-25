let
    Source = SharePoint.Files("https://mainroads.sharepoint.com/teams/mr-30000597-mebd-con-Commercial/", [ApiVersion = 15]),
    #"Filtered Rows" = Table.SelectRows(Source, each Text.StartsWith([Name], "MEBD-C15121-O-CO-REG-0001")),
    #"Filtered Hidden Files1" = Table.SelectRows(#"Filtered Rows", each [Attributes]?[Hidden]? <> true),
    #"Invoke Custom Function1" = Table.AddColumn(#"Filtered Hidden Files1", "Transform File (10)", each #"Transform File (10)"([Content])),
    #"Renamed Columns1" = Table.RenameColumns(#"Invoke Custom Function1", {"Name", "Source.Name"}),
    #"Removed Other Columns1" = Table.SelectColumns(#"Renamed Columns1", {"Source.Name", "Transform File (10)"}),
    #"Expanded Table Column1" = Table.ExpandTableColumn(#"Removed Other Columns1", "Transform File (10)", Table.ColumnNames(#"Transform File (10)"(#"Sample File (2)"))),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded Table Column1",{{"Source.Name", type text}, {"Date Raised", type date}, {"Potential Variation #", type text}, {"Description", type text}, {"Closed", type text}, {"Quote value", type any}, {"Contrator Ref #", type text}, {"DoA Approval TRIM #", type text}, {"Variation #", type text}, {"Variation Letter TRIM #", type text}, {"Next Action", type text}, {"Action with", type text}, {"Revised Contract Sum", type number}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"Source.Name"}),
    #"Filtered Rows1" = Table.SelectRows(#"Removed Columns", each ([Date Raised] <> null)),
    #"Renamed Columns" = Table.RenameColumns(#"Filtered Rows1",{{"Potential Variation #", "PV ID"}, {"Variation Letter TRIM #", "TRIM ID"}}),
    #"Changed Type2" = Table.TransformColumnTypes(#"Renamed Columns",{{"Revised Contract Sum", Currency.Type}}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Changed Type2",{{"Quote value", Currency.Type}}),
    #"Replaced Errors" = Table.ReplaceErrorValues(#"Changed Type1", {{"Quote value", null}}),
    #"Renamed Columns2" = Table.RenameColumns(#"Replaced Errors",{{"Variation #", "Variation ID"}}),
    #"Added Custom" = Table.AddColumn(#"Renamed Columns2", "Status", each if Text.StartsWith([Variation ID], "V") then "Approved" else if Text.StartsWith([Variation ID], "TBA") then "Pending" else "Unapproved"),
    #"Removed Columns1" = Table.RemoveColumns(#"Added Custom",{"Revised Contract Sum"}),
    #"Replaced Errors1" = Table.ReplaceErrorValues(#"Removed Columns1", {{"Status", "Unapproved"}})
in
    #"Replaced Errors1"