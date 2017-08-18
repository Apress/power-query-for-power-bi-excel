// DailySales
let
    Source = Excel.CurrentWorkbook(){[Name="DailySales"]}[Content],
    FilteredRows = Table.SelectRows(Source, each [Sales] > 5)
in
    FilteredRows

// UnPivot
let
    Source = Excel.CurrentWorkbook(){[Name="DailySales"]}[Content],
    Unpivot = Table.UnpivotOtherColumns(Source,{"Date", "WeekdayOrWeekend"},"Attribute","Value")
in
    Unpivot

// DailySales (2)
let
    Source = Excel.CurrentWorkbook(){[Name="DailySales"]}[Content],
    Custom1 = Table.Pivot(Source, {"Weekday", "Weekend"}, "WeekdayOrWeekend", "Sales")
in
    Custom1

// Pivot
let
    Source = Excel.CurrentWorkbook(){[Name="DailySales"]}[Content],
    PivottedTable = Table.Pivot(Source, List.Distinct(Table.Column(Source, "WeekdayOrWeekend")), "WeekdayOrWeekend", "Sales", List.Sum)
	
in
    PivottedTable