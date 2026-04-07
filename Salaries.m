let
    Source = Tx,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Vendor] = "Staff Costs (Contractor)" or [Vendor] = "Staff Costs (Main Roads)")),

    #"Sorted Rows" = Table.Sort(#"Filtered Rows", {{"Date", Order.Descending}}),

    #"Added Full Name" =
        Table.AddColumn(
            #"Sorted Rows",
            "Name",
            each
                let
                    c0 = if [Comment] <> null then Text.Trim([Comment]) else "",
                    c1 = Text.Combine(List.Select(Text.SplitAny(c0, " "), each _ <> ""), " "),
                    parts = Text.Split(c1, " "),
                    lastName = if List.Count(parts) > 4 then parts{4} else "",
                    givenTokens = if List.Count(parts) > 5 then List.Skip(parts, 5) else {},
                    givenName = if List.Count(givenTokens) > 0 then Text.Combine(givenTokens, " ") else ""
                in
                    Text.Combine({givenName, lastName}, " "),
            type text
        ),

    // Existing replacements
    #"Replaced Value"  = Table.ReplaceValue(#"Added Full Name","Adrian Patrick MINOGUE","Adrian MINOGUE",Replacer.ReplaceText,{"Name"}),
    #"Replaced Value1" = Table.ReplaceValue(#"Replaced Value","Ruwani Isurika TENNAKOON","Izzy TENNAKOON",Replacer.ReplaceText,{"Name"}),
    #"Replaced Value2" = Table.ReplaceValue(#"Replaced Value1","Denby Southey ADAMS","Denby ADAMS",Replacer.ReplaceText,{"Name"}),
    #"Replaced Value3" = Table.ReplaceValue(#"Replaced Value2","Andrew David IVES","Andrew IVES",Replacer.ReplaceText,{"Name"}),
    #"Replaced Value4" = Table.ReplaceValue(#"Replaced Value3","Benjamin Joel BEAVIS","Benjamin BEAVIS",Replacer.ReplaceText,{"Name"}),
    #"Replaced Value5" = Table.ReplaceValue(#"Replaced Value4","Ruwani TENNAKOON","Izzy TENNAKOON",Replacer.ReplaceText,{"Name"}),

    // Null-safe normalization of "contains" cases
    #"Normalized Names" =
        Table.TransformColumns(
            #"Replaced Value5",
            {
                "Name",
                each 
                    if _ = null then null
                    else if Text.Contains(_, "Nhiari Lipscombe") then "Nhiari Lipscombe"
                    else if Text.Contains(_, "Kym Fothergill") then "Kym Fothergill"
                    else if Text.Contains(_, "Tony Carlino") then "Tony Carlino"
                    else _,
                type text
            }
        ),

    #"Changed Type" = Table.TransformColumnTypes(#"Normalized Names",{{"GL Year", Int64.Type}, {"Period", Int64.Type}}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Changed Type",{{"Exp Code", Int64.Type}}),

    // Dynamically remove all "Column*" fields plus known extras
    ColumnsToRemove = List.Select(Table.ColumnNames(#"Changed Type1"), each Text.StartsWith(_, "Column")),
    #"Removed Columns1" = Table.RemoveColumns(#"Changed Type1", ColumnsToRemove & {"Expenditure Item ID", "Document", "Invoice Number", "IsAccrual", "IsProgressClaim"})
in
    #"Removed Columns1"
