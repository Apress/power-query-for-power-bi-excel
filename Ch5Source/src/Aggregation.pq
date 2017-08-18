// RowCount
let
    //Load data from source table
    Source = Excel.CurrentWorkbook(){[Name="SalesTable"]}[Content],
    //Count rows
    CountRowsInTable = Table.RowCount(Source)
in
    CountRowsInTable

// RowCountWithGroup
let
    //Load data from source table
    Source = Excel.CurrentWorkbook(){[Name="SalesTable"]}[Content],
    //Count rows and return a table
    GroupedRows = Table.Group(Source, {}, {{"Count of Rows", each Table.RowCount(_), type number}})
in
    GroupedRows

// CountAndSumWithGroup
let
    //Load data from source table
    Source = Excel.CurrentWorkbook(){[Name="SalesTable"]}[Content],
    //Group by Quarter, count rows and sum sales
    GroupedRows = Table.Group(
			Source, 
			{"Quarter"}, 
			{{"Count of Rows", each Table.RowCount(_), type number}, 
			{"Sum of Sales", each List.Sum([Sales]), type number}}
			)
in
    GroupedRows

// GroupKindGlobal
let
    Source = Excel.CurrentWorkbook(){[Name="DailySales"]}[Content],
    GroupedRows = Table.Group(Source, {"WeekdayOrWeekend"}, {{"Count of Rows", each Table.RowCount(_), type number}}, GroupKind.Global)
in
    GroupedRows

// GroupKindLocal
let
    Source = Excel.CurrentWorkbook(){[Name="DailySales"]}[Content],
    GroupedRows = Table.Group(Source, {"WeekdayOrWeekend"}, {{"Count of Rows", each Table.RowCount(_), type number}}, GroupKind.Local)
in
    GroupedRows

// TableSort
let
    Source = Excel.CurrentWorkbook(){[Name="DailySales"]}[Content],
    SortedRows = Table.Sort(Source,{{"WeekdayOrWeekend", Order.Descending}, {"Sales", Order.Ascending}})
in
    SortedRows