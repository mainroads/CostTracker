let
    Source = Excel.Workbook(Web.Contents("https://mainroads.sharepoint.com/teams/MR-30000597-MEBD-PRJ-Commercial/Shared%20Documents/Commercial/Cost%20Tracker/ProjectData.xlsx"), null, true),
    tblProjects_Table = Source{[Item="tblProjects",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(tblProjects_Table,{{"ProjectNumber", type text}, {"ProjectAbbr", type text}, {"ProjectDescription", type text}, {"DateforPC", type date}, {"AwardDate", type date}, {"ContractNumber", type text}}),
    BufferMeSideways = Table.Buffer(#"Changed Type")
in
    BufferMeSideways
