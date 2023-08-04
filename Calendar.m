let
   

   // Change start date to begining of year.
    // To automatically calculate this from the earliest date in a field, use
    // =List.Min(Table.Column(TableName, "Date Column"))
    #"StartDate"= List.Min(Table.Column(Budgets, "Date")),        
    
     // Change end date to end of year
    // To automatically calculate this from the latest date in a field, use
    // =List.Max(Table.Column(TableName, "Date Column"))
    #"EndDate" = List.Max(Table.Column(Forecasts, "Date")), 

    Source = PowerPlatform.Dataflows(null),
    Workspaces = Source{[Id="Workspaces"]}[Data],
    #"BI Platform Common" = Workspaces{[workspaceId="2743379b-ecc4-4445-87c4-5c672aac2764"]}[Data],
    #"Dimension Data" = #"BI Platform Common"{[dataflowId="e530cb41-3424-4f0a-83f2-f9e90277a4bb"]}[Data],
    Date_ = #"Dimension Data"{[entity="Date",version=""]}[Data],
    #"restrict Date Range" = Table.SelectRows(Date_, each [DateKey] >= #"StartDate" and [DateKey] <=  #"EndDate"),
    #"Removed Other Columns" = Table.SelectColumns(#"restrict Date Range",{"Date", "Year", "Month Number Of Year", "Month Short Name", "Day Name", "Day Of Week", "Calendar Quarter Number", "Calendar Quarter Name", "Day Relative", "Month Relative", "Year Relative", "Fiscal Month Number", "Fiscal Year", "Fiscal Quarter Number Text"}),
    #"Add YYYY/YY fin year" = Table.AddColumn(#"Removed Other Columns", "Financial Year", each Text.From(([Fiscal Year]-1)) & "/" & Text.From(Number.Mod([Fiscal Year],100)),type text),
    #"shorten Day Name" = Table.SplitColumn(#"Add YYYY/YY fin year", "Day Name", Splitter.SplitTextByPositions({0, 3}, false), {"Day Name", "Day Name.2"}),
    #"Removed Columns" = Table.RemoveColumns(#"shorten Day Name",{"Fiscal Year", "Day Name.2"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Columns",{"Date", "Year", "Month Number Of Year", "Month Short Name", "Day Name", "Day Of Week", "Calendar Quarter Number", "Calendar Quarter Name", "Day Relative", "Month Relative", "Year Relative", "Fiscal Month Number", "Financial Year", "Fiscal Quarter Number Text"}),
    colsForRename = {  {"Month Number Of Year","Month Number for Sort"}
                    ,{"Month Short Name","Month"}
                    ,{"Day Name","Day"}
                    ,{"Day Of Week","Day of Week for Sort"}
                    ,{"Calendar Quarter Number","Quarter"}
                    ,{"Calendar Quarter Name","YY-QQ"}
                    ,{"Day Relative","Days Since Today"}
                    ,{"Month Relative","Months Since Today"}
                    ,{"Year Relative","Years Since Today"}
                    ,{"Fiscal Month Number","Financial Month Number for Sort"}
                    ,{"Fiscal Quarter Number Text","Financial Quarter"}
                },
    renamed = Table.RenameColumns(#"Reordered Columns",colsForRename)
in
    renamed
