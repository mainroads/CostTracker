let
    Source = SharePoint.Files("https://mainroads.sharepoint.com/teams/MR-30000597-MEBD-PRJ-Commercial", [ApiVersion = 15]),
    #"Filtered Rows" = Table.SelectRows(Source, each Text.Contains([Extension], "csv")),
    #"Filtered Rows1" = Table.SelectRows(#"Filtered Rows", each [Name] = "TaskNumbers.csv"),
    #"TaskNumbers csv_https://mainroads sharepoint com/teams/MR-30000597-MEBD-PRJ-Commercial/Shared Documents/Commercial/Cost Tracker/Tx/" = #"Filtered Rows1"{[Name="TaskNumbers.csv",#"Folder Path"="https://mainroads.sharepoint.com/teams/MR-30000597-MEBD-PRJ-Commercial/Shared Documents/Commercial/Cost Tracker/Tx/"]}[Content],
    #"Imported CSV" = Csv.Document(#"TaskNumbers csv_https://mainroads sharepoint com/teams/MR-30000597-MEBD-PRJ-Commercial/Shared Documents/Commercial/Cost Tracker/Tx/",[Delimiter=",", Columns=2, Encoding=1252, QuoteStyle=QuoteStyle.None]),
    #"Changed Type" = Table.TransformColumnTypes(#"Imported CSV",{{"Column1", type text}, {"Column2", type text}}),
    #"Renamed Columns" = Table.RenameColumns(#"Changed Type",{{"Column1", "Task Number"}, {"Column2", "Task Name"}})
in
    #"Renamed Columns"