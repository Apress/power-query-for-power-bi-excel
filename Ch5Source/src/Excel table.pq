// Query1
/* This query loads data from the Content table on the worksheet
   and then filters and aggregates it */
let
    //Load data from Content table
    Source = Excel.CurrentWorkbook(){[Name="SalesTable"]}[Content],
    //Filter where Sales>5
    FilteredRows = Table.SelectRows(Source, each [Sales] > 5),
    //Sort in descending order by Sales
    SortedRows = Table.Sort( /* this is an embedded comment*/ FilteredRows,{{"Sales", Order.Descending}})
/* comment */
/* comment */ /* comment */

in
    SortedRows