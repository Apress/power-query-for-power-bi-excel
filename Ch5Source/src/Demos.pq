// Errors
let
    Source = Excel.CurrentWorkbook(){[Name="Table1"]}[Content],
    ChangedType = Table.TransformColumnTypes(Source,{{"TextColumn", type text}}),
    InsertedCustom = Table.AddColumn(ChangedType, "Custom", each [NumberColumn] + (try Number.FromText([TextColumn]) otherwise 0))
in
    InsertedCustom

// CheckType
let
    Source = Excel.CurrentWorkbook(){[Name="Table1"]}[Content],
    InsertedCustom = Table.AddColumn(Source, "Custom", each [TextColumn] is number)
in
    InsertedCustom

// ConditionalLogic
let
    Source = Excel.CurrentWorkbook(){[Name="Table1"]}[Content],
    InsertedCustom = Table.AddColumn(Source, "Custom", each if [NumberColumn]>2 then "More than 2" else "2 or less")
in
    InsertedCustom

// Query1
let
    Source = ""
in
    Source