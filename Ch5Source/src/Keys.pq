// DimDateKeys
let
    Source = Sql.Database("localhost", "adventure works dw"),
    dbo_DimDate = Source{[Schema="dbo",Item="DimDate"]}[Data],
    Custom1 = Table.Keys(dbo_DimDate),
    TableFromList = Table.FromList(Custom1, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    #"Expand Column1" = Table.ExpandRecordColumn(TableFromList, "Column1", {"Columns", "Primary"}, {"Column1.Columns", "Column1.Primary"}),
    #"Expand Column1.Columns" = Table.ExpandListColumn(#"Expand Column1", "Column1.Columns")
in
    #"Expand Column1.Columns"

// Keys
let
    //Load data from source table
    Source = Excel.CurrentWorkbook(){[Name="DailySales"]}[Content],
    //Define a primary key on the Date column
    DefineDateKey = Table.AddKey(Source, {"Date"}, true)
in
    DefineDateKey

// DailySales
let
    Source = Excel.CurrentWorkbook(){[Name="DailySales"]}[Content],
    DuplicatesRemoved = Table.Distinct(Source, {"Date"}),
    Custom1 = Table.Keys(DuplicatesRemoved),
    TableFromList = Table.FromList(Custom1, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    Column1 = TableFromList{0}[Column1],
    Columns = Column1[Columns]
in
    Columns