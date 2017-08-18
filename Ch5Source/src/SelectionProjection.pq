// ListItems
let
    Source = {"A", "B", "C"}{4}?
in
    Source

// FruitSales
let
    //Load data from source table
    Source = Excel.CurrentWorkbook(){[Name="FruitSales"]}[Content],
    DuplicatesRemoved = Table.Distinct(Source, {"Fruit"}),
    Apples = DuplicatesRemoved{[Fruit="Apples"]}[Fruit]
in
    Apples