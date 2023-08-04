let
    Source = SharePoint.Files("https://mainroads.sharepoint.com/teams/MR-30000597-MEBD-PRJ-Commercial/", [ApiVersion = 15]),
    #"Filtered Rows" = Table.SelectRows(Source, each ([Extension] = ".xlsx" or [Extension] = ".XLSX")),
    #"Filtered Rows1" = Table.SelectRows(#"Filtered Rows", each Text.Contains([Name], "-FMS_PA_Report07")),
    #"Filtered Hidden Files1" = Table.SelectRows(#"Filtered Rows1", each [Attributes]?[Hidden]? <> true),
    #"Invoke Custom Function1" = Table.AddColumn(#"Filtered Hidden Files1", "Transform File (7)", each #"Transform File (7)"([Content])),
    #"Renamed Columns1" = Table.RenameColumns(#"Invoke Custom Function1", {"Name", "Source.Name"}),
    #"Removed Other Columns1" = Table.SelectColumns(#"Renamed Columns1", {"Source.Name", "Transform File (7)"}),
    #"Expanded Table Column1" = Table.ExpandTableColumn(#"Removed Other Columns1", "Transform File (7)", Table.ColumnNames(#"Transform File (7)"(#"Sample File (6)"))),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded Table Column1",{{"Date", type date}}),
    #"Added Conditional Column1" = Table.AddColumn(#"Changed Type" , "Vendor", each if [Vendor Name] = null then [Task Desc] else [Vendor Name]),
    #"Inserted Text Before Delimiter" = Table.AddColumn(#"Added Conditional Column1", "Project Number", each Text.BeforeDelimiter([Source.Name], "-"), type text),
    #"Removed Columns8" = Table.RemoveColumns(#"Inserted Text Before Delimiter",{"Source.Name"}),
    #"Merged Queries" = Table.NestedJoin(#"Removed Columns8", {"Project Number"}, Projects, {"ProjectNumber"}, "Projects", JoinKind.Inner),
    #"Expanded Projects" = Table.ExpandTableColumn(#"Merged Queries", "Projects", {"ContractNumber"}, {"ContractNumber"}),
    #"Added Conditional Column2" = Table.AddColumn(#"Expanded Projects", "IsProgressClaim", each ([Resource] <> null and Text.Contains([Resource], "021")),type logical),
    #"Changed Type1" = Table.TransformColumnTypes(#"Added Conditional Column2",{{"GL Year", Int64.Type}, {"Period", Int64.Type},{"Amount", Currency.Type}, {"Quantity", type number}, {"Line No", Int64.Type}}),
    #"columns to be text" = Table.SelectRows(Table.Schema(#"Changed Type1"), each ([TypeName] = "Any.Type")),
    setAnyToText = Table.TransformColumnTypes(#"Changed Type1", List.Transform(#"columns to be text"[Name], each {_, type text})),
    #"Renamed Columns" = Table.RenameColumns(setAnyToText,{{"Task No","Task Number"}
,{"Task Desc","Task Name"}
,{"Expenditure Type","Exp Code"}
,{"Expenditure Type Desc","Exp Name"}
,{"Invoice Num","Invoice Number"}
}),


    selectCols = Table.SelectColumns(#"Renamed Columns",{"Date"
,"Task Number"
,"Task Name"
,"Exp Code"
,"Exp Name"
,"Resource"
,"Amount"
,"Comment"
,"Invoice Number"
,"Vendor"
,"Project Number"
,"ContractNumber"
,"IsProgressClaim"
}),

    BufferMyTable = Table.Buffer(selectCols),
    #"Inserted Merged Column" = Table.AddColumn(BufferMyTable, "Task Description", each Text.Combine({[Task Number], [Task Name]}, " "), type text),
    #"Inserted Merged Column1" = Table.AddColumn(#"Inserted Merged Column", "Expenditure Description", each Text.Combine({[Exp Code], [Exp Name]}, " "), type text)
in
    #"Inserted Merged Column1"
