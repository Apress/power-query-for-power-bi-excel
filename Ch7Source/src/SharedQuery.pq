// ProductCategory
let
    Source = Sql.Database("localhost", "adventure works dw"),
    dbo_DimProductCategory = Source{[Schema="dbo",Item="DimProductCategory"]}[Data],
    FilteredRows = Table.SelectRows(dbo_DimProductCategory, each ([ProductCategoryKey] <> 4))
in
    FilteredRows

// DaysPerYear
let
    Source = Sql.Database("localhost", "adventure works dw"),
    dbo_DimDate = Source{[Schema="dbo",Item="DimDate"]}[Data],
    GroupedRows = Table.Group(dbo_DimDate, {"CalendarYear"}, {{"Count", each Table.RowCount(_), type number}})
in
    GroupedRows