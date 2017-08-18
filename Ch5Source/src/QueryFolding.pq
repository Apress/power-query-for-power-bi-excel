// DimProductCategory
let
    //connect to SQL Server
    Source = Sql.Database("localhost", "Adventure Works DW"),
    //connect to the DimProductCategory table
    dbo_DimProductCategory = Source{[Schema="dbo",Item="DimProductCategory"]}[Data],
    //filter the table where ProductCategoryKey is less than 4
    FilteredRows = Table.SelectRows(dbo_DimProductCategory, each [ProductCategoryKey] < 4),
    InsertedCustom = Table.AddColumn(FilteredRows, "Custom", each if [ProductCategoryKey]>3 then 1 else if [ProductCategoryKey]>2 then 2 else 3)
in
    InsertedCustom

// DimProductCategoryBuffered
let
    //connect to SQL Server
    Source = Sql.Database("localhost", "Adventure Works DW"),
    //connect to the DimProductCategory table
    dbo_DimProductCategory = Source{[Schema="dbo",Item="DimProductCategory"]}[Data],
    //buffer table
    BufferedTable = Table.Buffer(dbo_DimProductCategory),
    //filter the table where ProductCategoryKey is less than 4
    FilteredRows = Table.SelectRows(BufferedTable, each [ProductCategoryKey] < 4),
    //count the number of rows in the resulting table
    GroupedRows = Table.Group(FilteredRows, {}, {{"Count", each Table.RowCount(_), type number}})
in
    GroupedRows

// DimProductCategoryErrorFilter
let
    //connect to SQL Server
    Source = Sql.Database("localhost", "Adventure Works DW"),
    //connect to the DimProductCategory table
    dbo_DimProductCategory = Source{[Schema="dbo",Item="DimProductCategory"]}[Data],
    //insert custom column that returns an error in one row
    InsertedCustom = Table.AddColumn(dbo_DimProductCategory, "DivideByZero", each if [ProductCategoryKey]=1 then xyz else null),
    //filter table by ProductCategoryKey < 4
    FilteredRows = Table.SelectRows(InsertedCustom , each [ProductCategoryKey] < 4),
    //remove error row
    RemovedErrors = Table.RemoveRowsWithErrors(FilteredRows , {"DivideByZero"})

in
    RemovedErrors

// DimProductCategoryUnpivot
let
    //connect to SQL Server
    Source = Sql.Database("localhost", "Adventure Works DW"),
    //connect to the DimProductCategory table
    dbo_DimProductCategory = Source{[Schema="dbo",Item="DimProductCategory"]}[Data],
    Unpivot = Table.UnpivotOtherColumns(dbo_DimProductCategory,{"ProductCategoryKey"},"Attribute","Value")
in
    Unpivot

// DimProductCategoryPivot
let
    //connect to SQL Server
    Source = Sql.Database("localhost", "Adventure Works DW"),
    //connect to the DimProductCategory table
    dbo_DimProductCategory = Source{[Schema="dbo",Item="DimProductCategory"]}[Data],
    Custom1 = Table.Pivot(dbo_DimProductCategory, {"1", "2", "3", "4"}, "ProductCategoryKey", "EnglishProductCategoryName")
in
    Custom1

// DimProductCategoryFunctions
let
    //connect to SQL Server
    Source = Sql.Database("localhost", "Adventure Works DW"),
    //connect to the DimProductCategory table
    dbo_DimProductCategory = Source{[Schema="dbo",Item="DimProductCategory"]}[Data],
    AddValues = (x,y) => x + y,
    InsertedCustom = Table.AddColumn(dbo_DimProductCategory, "Custom", each AddValues([ProductCategoryKey],[ProductCategoryAlternateKey]))
in
    InsertedCustom