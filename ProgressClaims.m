let
    Source = SharePoint.Files("https://mainroads.sharepoint.com/teams/mr-30000597-mebd-con-Commercial", [ApiVersion = 15]),
    #"Filtered Rows" = Table.SelectRows(Source, each Text.StartsWith([Name], "MEBD-C15121-O-CO-PC")),
    #"Filtered Rows1" = Table.SelectRows(#"Filtered Rows", each ([Extension] = ".xlsm" or [Extension] = ".XLSM")),
    #"Filtered Rows2" = Table.SelectRows(#"Filtered Rows1", each not Text.Contains([Name], "Review")),
    #"Filtered Rows3" = Table.SelectRows(#"Filtered Rows2", each ([Name] <> "MEBD-C15121-O-CO-PC-00NN Progress Claim for Works in Month Year.xlsm")),
    #"Sorted Rows" = Table.Sort(#"Filtered Rows3",{{"Folder Path", Order.Ascending}}),
    #"Filtered Hidden Files1" = Table.SelectRows(#"Sorted Rows", each [Attributes]?[Hidden]? <> true),
    #"Invoke Custom Function1" = Table.AddColumn(#"Filtered Hidden Files1", "Transform File (11)", each #"Transform File (11)"([Content])),
    #"Renamed Columns1" = Table.RenameColumns(#"Invoke Custom Function1", {"Name", "Source.Name"}),
    #"Removed Other Columns1" = Table.SelectColumns(#"Renamed Columns1", {"Source.Name", "Transform File (11)"}),
    #"Expanded Table Column1" = Table.ExpandTableColumn(#"Removed Other Columns1", "Transform File (11)", Table.ColumnNames(#"Transform File (11)"(#"Sample File (7)"))),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded Table Column1",{{"Source.Name", type text}, {"MANDURAH ESTUARY BRIDGE DUPLICATION (MEBD) PROJECT", type any}, {"Column2", type text}, {"Column3", type any}, {"Column4", type any}, {"Column5", type any}, {"Column6", type any}, {"Column7", type any}, {"Column8", type any}, {"Column9", type any}, {"PAYMENT CERTIFICATE No. 1", type text}, {"Column11", type number}, {"Column12", type any}, {"Column13", type any}, {"Column14", type any}, {"Column15", type text}, {"Column16", type text}, {"Column17", type text}, {"Column18", type text}, {"Column19", type any}, {"Column20", type text}}),
    #"Renamed Columns" = Table.RenameColumns(#"Changed Type",{{"MANDURAH ESTUARY BRIDGE DUPLICATION (MEBD) PROJECT", "CC"}}),
    #"Removed Columns" = Table.RemoveColumns(#"Renamed Columns",{"Column9", "PAYMENT CERTIFICATE No. 1", "Column11", "Column12", "Column13", "Column14", "Column15", "Column16", "Column17", "Column18", "Column19", "Column20"}),
    #"Filtered Rows4" = Table.SelectRows(#"Removed Columns", each ([Column8] <> null)),
    #"Renamed Columns2" = Table.RenameColumns(#"Filtered Rows4",{{"Column2", "Description"}}),
    #"Filtered Rows5" = Table.SelectRows(#"Renamed Columns2", each true),
    #"Added Custom" = Table.AddColumn(#"Filtered Rows5", "Start Date", each let
    SourceText = [Source.Name],
    MonthText = Text.BetweenDelimiters(SourceText, "in ", " "),
    YearText = Text.Middle(SourceText, Text.PositionOf(SourceText, MonthText) + Text.Length(MonthText) + 1, 4),
    MonthNumber = Date.Month(Date.FromText("1 " & MonthText & " " & YearText))
  in
    #date(Number.FromText(YearText), MonthNumber, 1)),
    #"Changed Type1" = Table.TransformColumnTypes(#"Added Custom",{{"Start Date", type date}}),
    #"Added Custom1" = Table.AddColumn(#"Changed Type1", "Finish Date", each Date.EndOfMonth([Start Date])),
    #"Changed Type2" = Table.TransformColumnTypes(#"Added Custom1",{{"Finish Date", type date}}),
    #"Removed Columns1" = Table.RemoveColumns(#"Changed Type2",{"Column4", "Column6"}),
    #"Renamed Columns3" = Table.RenameColumns(#"Removed Columns1",{{"Column5", "% Complete"}}),
    #"Changed Type3" = Table.TransformColumnTypes(#"Renamed Columns3",{{"Column3", Currency.Type}}),
    #"Renamed Columns4" = Table.RenameColumns(#"Changed Type3",{{"Column3", "Contract Value"}}),
    #"Replaced Errors" = Table.ReplaceErrorValues(#"Renamed Columns4", {{"Contract Value", null}}),
    #"Removed Columns2" = Table.RemoveColumns(#"Replaced Errors",{"Column7"}),
    #"Renamed Columns5" = Table.RenameColumns(#"Removed Columns2",{{"Column8", "Claim Value"}}),
    #"Changed Type4" = Table.TransformColumnTypes(#"Renamed Columns5",{{"Claim Value", Currency.Type}}),
    #"Replaced Value" = Table.ReplaceValue(#"Changed Type4", each [CC], each if [CC] = 1.1000000000000001 then if Text.Contains([Description], "insurance", Comparer.OrdinalIgnoreCase) then "1.1" else "1.10" else [CC], Replacer.ReplaceValue, {"CC"}),
    #"Replaced Value1" = Table.ReplaceValue(#"Replaced Value","1.1100000000000001","1.11",Replacer.ReplaceValue,{"CC"}),
    #"Changed Type5" = Table.TransformColumnTypes(#"Replaced Value1",{{"CC", type text}}),
    #"Replaced Value2" = Table.ReplaceValue(#"Changed Type5","1.1100000000000001","1.11",Replacer.ReplaceText,{"CC"}),
    #"Replaced Value3" = Table.ReplaceValue(#"Replaced Value2","1.1200000000000001","1.12",Replacer.ReplaceText,{"CC"}),
    #"Replaced Value4" = Table.ReplaceValue(#"Replaced Value3","1.1299999999999999","1.13",Replacer.ReplaceText,{"CC"}),
    #"Replaced Value5" = Table.ReplaceValue(#"Replaced Value4","1.1399999999999999","1.14",Replacer.ReplaceText,{"CC"}),
    #"Replaced Value6" = Table.ReplaceValue(#"Replaced Value5","1.1499999999999999","1.15",Replacer.ReplaceText,{"CC"}),
    #"Replaced Value7" = Table.ReplaceValue(#"Replaced Value6","1.1599999999999999","1.16",Replacer.ReplaceText,{"CC"}),
    #"Replaced Value8" = Table.ReplaceValue(#"Replaced Value7","2.2000000000000002","2.2",Replacer.ReplaceText,{"CC"}),
    #"Replaced Value9" = Table.ReplaceValue(#"Replaced Value8","2.2999999999999998","2.3",Replacer.ReplaceText,{"CC"}),
    #"Replaced Value10" = Table.ReplaceValue(#"Replaced Value9","4.0999999999999996","4.1",Replacer.ReplaceText,{"CC"}),
    #"Replaced Value11" = Table.ReplaceValue(#"Replaced Value10","4.4000000000000004","4.4",Replacer.ReplaceText,{"CC"}),
    #"Replaced Value12" = Table.ReplaceValue(#"Replaced Value11","4.9000000000000004","4.9",Replacer.ReplaceText,{"CC"}),
    #"Replaced Value13" = Table.ReplaceValue(#"Replaced Value12","5.0999999999999996","5.1",Replacer.ReplaceText,{"CC"}),
    #"Replaced Value14" = Table.ReplaceValue(#"Replaced Value13","8.1999999999999993","8.2",Replacer.ReplaceText,{"CC"}),
    #"Replaced Value15" = Table.ReplaceValue(#"Replaced Value14","8.3000000000000007","8.3",Replacer.ReplaceText,{"CC"}),
    #"Replaced Value16" = Table.ReplaceValue(#"Replaced Value15","9.1999999999999993","9.2",Replacer.ReplaceText,{"CC"}),
    #"Replaced Value17" = Table.ReplaceValue(#"Replaced Value16","9.3000000000000007","9.3",Replacer.ReplaceText,{"CC"}),
    #"Replaced Value18" = Table.ReplaceValue(#"Replaced Value17","10.199999999999999","10.2",Replacer.ReplaceText,{"CC"}),
    #"Added PC Number" = Table.AddColumn(#"Replaced Value18", "PC Number", each let
    SourceText = [Source.Name],
    PCText = Text.BetweenDelimiters(SourceText, "PC-", " "),
    PCNumber = "PC" & Text.End(PCText, 2)
  in
    PCNumber),
    ReplaceDescriptionNull = Table.ReplaceValue(#"Added PC Number",null,"n/a",Replacer.ReplaceValue,{"Description"}),
    ReplaceCCNull = Table.ReplaceValue(ReplaceDescriptionNull,null,"n/a",Replacer.ReplaceValue,{"CC"}),
    AddedCCtype = Table.AddColumn(ReplaceCCNull, "CC Type", each if Text.Contains([Description], "RISE & FALL") then "Rise & Fall" else if [CC] = "n/a" then "n/a" else if Text.Contains([CC], ".") then "Milestone" else if Text.StartsWith([CC], "V") then "Variation" else if Text.StartsWith([CC], "PV") then "Potential Variation" else "Cost Centre"),
    #"Filtered Rows6" = Table.SelectRows(AddedCCtype, each ([CC Type] <> "n/a") and ([Description] <> "n/a")),
    #"Changed Type6" = Table.TransformColumnTypes(#"Filtered Rows6",{{"% Complete", Percentage.Type}}),
    #"Merged Queries" = Table.NestedJoin(#"Changed Type6", {"PC Number"}, #"Rise and Fall", {"PC Number"}, "Rise and Fall", JoinKind.LeftOuter),
    #"Expanded Rise and Fall" = Table.ExpandTableColumn(#"Merged Queries", "Rise and Fall", {"Rise & Fall Amount"}, {"Rise and Fall.Rise & Fall Amount"}),
    #"Renamed Columns6" = Table.RenameColumns(#"Expanded Rise and Fall",{{"Rise and Fall.Rise & Fall Amount", "Rise and Fall Amount"}}),
    #"Added Custom3" = Table.AddColumn(#"Renamed Columns6", "Amount", each if [CC Type] = "Rise & Fall" then [Rise and Fall Amount] else [Claim Value]),
    #"Changed Type7" = Table.TransformColumnTypes(#"Added Custom3",{{"Amount", Currency.Type}}),
    #"Removed Columns3" = Table.RemoveColumns(#"Changed Type7",{"Rise and Fall Amount"}),
    #"Added CCsort" = Table.AddColumn(#"Removed Columns3", "CC Sort", each let
    CCValue = [CC],
    DescriptionValue = [Description],
    
    // Check if CC is "n/a"
    NAValue = if CCValue = "n/a" then "99999" else null,

    // Check if CC starts with "PV" or "V"
    AdoptValue = if Text.StartsWith(CCValue, "PV") or Text.StartsWith(CCValue, "V") then CCValue else null,

    // Separate numeric and alphabetic parts
    NumericPart = Text.Select(CCValue, {"0".."9", "."}),
    AlphabeticPart = Text.Select(CCValue, {"a".."z", "A".."Z"}),

    // Extract integer part and decimal part
    IntegerPart = if Text.Contains(NumericPart, ".") then Number.FromText(Text.BeforeDelimiter(NumericPart, ".")) else Number.FromText(NumericPart),
    DecimalPart = if Text.Contains(NumericPart, ".") then Text.AfterDelimiter(NumericPart, ".") else "0",
    
    // Adjust DecimalIncrement based on Description
    DecimalIncrement = if Text.StartsWith(DescriptionValue, "Additional") or Text.StartsWith(DescriptionValue, "Accommodation") 
                       then Number.FromText(DecimalPart) * 100 
                       else Number.FromText(DecimalPart) * 10,

    // Calculate the final value
    FinalValue = IntegerPart * 1000 + DecimalIncrement,

    // Add alphabetic value if present
    AlphabetValue = if Text.Length(AlphabeticPart) > 0 then Character.ToNumber(Text.Lower(AlphabeticPart)) - 96 else 0,
    FinalValueWithAlpha = FinalValue + AlphabetValue,

    // Ensure the output is always 5 digits
    Result = Text.PadStart(Text.From(FinalValueWithAlpha), 5, "0")
in
    if CCValue = null then null 
    else if NAValue <> null then NAValue 
    else if AdoptValue <> null then AdoptValue 
    else Result),
    #"Replaced Value19" = Table.ReplaceValue(#"Added CCsort","Pit Clearing","Pit Cleaning",Replacer.ReplaceText,{"Description"}),
    #"Added Custom2" = Table.AddColumn(#"Replaced Value19", "CC Description", each [CC] & " " & [Description]),
    #"Replaced Value20" = Table.ReplaceValue(#"Added Custom2","n/a ","",Replacer.ReplaceText,{"CC Description"}),
    #"Filtered Rows7" = Table.SelectRows(#"Replaced Value20", each true)
in
    #"Filtered Rows7"