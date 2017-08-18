// DimDate
let
    Source = Sql.Database("localhost", "adventure works dw"),
    dbo_DimDate = Source{[Schema="dbo",Item="DimDate"]}[Data],
    FilteredRows = Table.SelectRows(dbo_DimDate, each ([EnglishDayNameOfWeek] = Weekday[Weekday]{0}))
in
    FilteredRows

// Weekday
let
    Source = Excel.CurrentWorkbook(){[Name="Weekday"]}[Content]
in
    Source

// DimDate2
let
    Source = Sql.Database("localhost", "adventure works dw", [Query="select * from DimDate where EnglishDayNameOfWeek='" & Weekday[Weekday]{0} & "'" ])
in
    Source

// DimDate3
let
    GetWeekDay = Weekday[Weekday]{0},
    GetSQL = "select * from DimDate where EnglishDayNameOfWeek='" & GetWeekDay  & "'",
    Source = Sql.Database("localhost", "adventure works dw",  [Query=GetSQL])
in
    Source

// FridayFromDimDate
let
    Source = Sql.Database("localhost", "adventure works dw"),
    dbo_DimDate = Source{[Schema="dbo",Item="DimDate"]}[Data],
    RemovedOtherColumns = Table.SelectColumns(dbo_DimDate,{"EnglishDayNameOfWeek"}),
    FilteredRows = Table.SelectRows(RemovedOtherColumns, each ([EnglishDayNameOfWeek] = "Friday")),
    DuplicatesRemoved = Table.Distinct(FilteredRows),
    Friday = DuplicatesRemoved{[EnglishDayNameOfWeek="Friday"]}[EnglishDayNameOfWeek]
in
    Friday

// DimDate4
let
    GetWeekDay = Excel.CurrentWorkbook(){[Name="Weekday"]}[Content],
    Source = Sql.Database("localhost", "adventure works dw",  [Query="select * from DimDate where EnglishDayNameOfWeek='" & GetWeekDay[Weekday]{0}  & "'"])
in
    Source

// DimDateFunction
let
    DimDateFunction = (WeekdayName) =>
let
    Source = Sql.Database("localhost", "adventure works dw"),
    dbo_DimDate = Source{[Schema="dbo",Item="DimDate"]}[Data],
    FilteredRows = Table.SelectRows(dbo_DimDate, each ([EnglishDayNameOfWeek] = WeekdayName))
in
    FilteredRows,
    InvokedDimDateFunction = DimDateFunction("Monday")
in
    InvokedDimDateFunction