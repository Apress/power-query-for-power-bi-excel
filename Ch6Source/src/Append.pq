// Apples
let
    Source = Excel.CurrentWorkbook(){[Name="Apples"]}[Content]
in
    Source

// Oranges
let
    Source = Excel.CurrentWorkbook(){[Name="Oranges"]}[Content]
in
    Source

// Pears
let
    Source = Excel.CurrentWorkbook(){[Name="Pears"]}[Content]
in
    Source

// ApplesProfit
let
    Source = Excel.CurrentWorkbook(){[Name="ApplesProfit"]}[Content]
in
    Source

// Append1
let
    Source = Table.Combine({Apples,Oranges}),
    Append = Table.Combine({Source,Pears})
in
    Append

// TableCombine1
let
    Source = Table.Combine({Apples, Oranges, Pears})
in
    Source

// TableCombine2
let
    Source = Table.Combine({Apples, ApplesProfit})
in
    Source

// Merge1
let
    Source = Table.NestedJoin(Apples,{"Month", "Product"},ApplesProfit,{"Month", "Product"},"NewColumn",JoinKind.Inner),
    #"Expand NewColumn" = Table.ExpandTableColumn(Source, "NewColumn", {"Month", "Product", "Profit"}, {"NewColumn.Month", "NewColumn.Product", "NewColumn.Profit"})
in
    #"Expand NewColumn"

// Merge2
let
    Source = Table.NestedJoin(Apples,{"Month", "Product"},ApplesProfit,{"Month", "Product"},"NewColumn"),
    #"Expand NewColumn" = Table.ExpandTableColumn(Source, "NewColumn", {"Month", "Product", "Profit"}, {"NewColumn.Month", "NewColumn.Product", "NewColumn.Profit"})
in
    #"Expand NewColumn"

// TableJoin
let
    Source = Table.Join(
Apples,{"Month", "Product"},
ApplesProfit,{"Month","Product"}
,JoinKind.Inner)
in
    Source