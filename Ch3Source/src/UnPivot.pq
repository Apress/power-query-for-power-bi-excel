// Unpivot
let
    Source = Excel.CurrentWorkbook(){[Name="Table1"]}[Content],
    Unpivot = Table.UnpivotOtherColumns(Source,{"Product"},"Attribute","Value")
in
    Unpivot

// Transpose

let
    Source = Excel.CurrentWorkbook(){[Name="Table1"]}[Content],
    TransposedTable = Table.Transpose(Source),
    FirstRowAsHeader = Table.PromoteHeaders(TransposedTable)
in
    FirstRowAsHeader