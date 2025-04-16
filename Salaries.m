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
    #"Kept Necessary Columns" = Table.SelectColumns(#"Filtered Hidden Files2", {"Source.Name", "FMS_PA_Report07_Resource"}),

    // 4. Define a function to transform the Resource sheet table from each file
    TransformResource = (tbl as nullable table) as table =>
        if tbl = null then
            #table({}, {})
        else
            let
                Skipped = Table.Skip(tbl, 8),
                Promoted = Table.PromoteHeaders(Skipped, [PromoteAllScalars=true]),
                ExpectedColumns = {
                    "GL Year", "Period", "Date", "Task No", "Task Desc", "Expenditure Type", 
                    "Expenditure Type Desc", "Resource", "Amount", "Quantity", "Expenditure Item ID", 
                    "Orig Transaction Reference", "Line No", "Document", "Transaction Source", 
                    "GL Batch Name", "Comment", "Vendor Name", "Purchase order no", "Agency Specific Contract", 
                    "GLAcct", "Invoice Num"
                },
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
                        Table.ReorderColumns(TableWithAdded, expected),
                Normalized = AddMissingColumns(Promoted, ExpectedColumns)
            in
                Normalized,

    // 5. Apply the transformation to each file's resource sheet
    #"Transformed Resource" = Table.AddColumn(#"Kept Necessary Columns", "TransformedResource", each TransformResource([FMS_PA_Report07_Resource])),
    #"Removed Original Sheet" = Table.RemoveColumns(#"Transformed Resource",{"FMS_PA_Report07_Resource"}),

    // 6. Expand the normalized resource data.
    ExpandedResource = Table.ExpandTableColumn(
        #"Removed Original Sheet", 
        "TransformedResource", 
        Table.ColumnNames(#"Removed Original Sheet"{0}[TransformedResource])
    ),

    #"Changed Type" = Table.TransformColumnTypes(ExpandedResource,{{"Date", type date}}),
    #"Added Conditional Column1" = Table.AddColumn(#"Changed Type" , "Vendor", each if [Vendor Name] = null then [Expenditure Type Desc] else [Vendor Name]),
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
    #"Inserted Merged Column1" = Table.AddColumn(#"Inserted Merged Column", "Expenditure Description", each Text.Combine({[Exp Code], [Exp Name]}, " "), type text),
    #"Filtered Rows3" = Table.SelectRows(#"Inserted Merged Column1", each Text.StartsWith([Comment], "Salary PE")),
    #"Sorted Rows" = Table.Sort(#"Filtered Rows3",{{"Date", Order.Descending}}),
    #"Added Custom" = Table.AddColumn(#"Sorted Rows", "Name", each let
    Comment = [Comment],
    Parts = Text.Split(Comment, " "),
    LastName = Parts{4},
    FirstName = Text.Combine(List.Skip(Parts, 5), " ")
in
    Text.Combine({FirstName, LastName}, " ")),
    #"Replaced Value" = Table.ReplaceValue(#"Added Custom","Adrian Patrick MINOGUE","Adrian MINOGUE",Replacer.ReplaceText,{"Name"}),
    #"Replaced Value1" = Table.ReplaceValue(#"Replaced Value","Ruwani Isurika TENNAKOON","Izzy TENNAKOON",Replacer.ReplaceText,{"Name"}),
    #"Replaced Value2" = Table.ReplaceValue(#"Replaced Value1","Denby Southey ADAMS","Denby ADAMS",Replacer.ReplaceText,{"Name"}),
    #"Replaced Value3" = Table.ReplaceValue(#"Replaced Value2","Andrew David IVES","Andrew IVES",Replacer.ReplaceText,{"Name"}),
    #"Replaced Value4" = Table.ReplaceValue(#"Replaced Value3","Benjamin Joel BEAVIS","Benjamin BEAVIS",Replacer.ReplaceText,{"Name"}),
    #"Replaced Value5" = Table.ReplaceValue(#"Replaced Value4","Ruwani TENNAKOON","Izzy TENNAKOON",Replacer.ReplaceText,{"Name"})
in
    #"Replaced Value5"