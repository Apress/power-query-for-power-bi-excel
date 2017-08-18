// ProductQuery
let
    Source = Excel.CurrentWorkbook(){[Name="Product"]}[Content]
in
    Source

// ProductQuery (2)
let
    Source = ProductQuery
in
    Source

// Sales
let
    Source = Excel.CurrentWorkbook(){[Name="Sales"]}[Content],
    FilteredRows = Table.SelectRows(Source, each ([Product] = ProductQuery[Product]{0}))
in
    FilteredRows