// PercentageShareOfTotalInput
let
    //Load data from the Excel worksheet
    Source = Excel.CurrentWorkbook(){[Name="PercentageShareOfTotalInput"]}[Content],
    //Return a table containing the grand total of sales
    GroupedRows = Table.Group(Source, {}, {{"GrandTotal", each List.Sum([Sales]), type number}}),
    //Returns the grand total as a value on its own
    GrandTotal = GroupedRows{0}[GrandTotal],
    //Fetch the original input table
    Custom1 = Source,
    //Insert a custom column to calculate the percentage share
    InsertedCustom = Table.AddColumn(Custom1, "Share", each [Sales]/GrandTotal)
in
    InsertedCustom

// PercentageGrowthPrevDayInput
let
    //Load data from Excel workbook
    Source = Excel.CurrentWorkbook(){[Name="PercentageGrowthPrevDayInput"]}[Content],
    //Set the type of the Date column to be date
    ChangedType = Table.TransformColumnTypes(Source,{{"Date", type date}}),
    //Calculate the date one day before the current date
    InsertedCustom = Table.AddColumn(ChangedType, "PreviousDate", each [Date] - #duration(1,0,0,0)),
    //Join the table to itself to look up the previous date value
    Merge = Table.NestedJoin(InsertedCustom,{"PreviousDate"},InsertedCustom,{"Date"},"NewColumn"),
    //Expand the Sales column from the joined table
    #"Expand NewColumn" = Table.ExpandTableColumn(Merge, "NewColumn", {"Sales"}, {"NewColumn.Sales"}),
    //Rename the new column to PreviousDateSales
    RenamedColumns = Table.RenameColumns(#"Expand NewColumn",{{"NewColumn.Sales", "PreviousDateSales"}}),
    //Calculate the percentage growth
    InsertedCustom1 = Table.AddColumn(RenamedColumns, "PercentageSalesGrowth", each ([Sales] - [PreviousDateSales]) / [PreviousDateSales]),
    //Remove the PreviousDate and PreviousDateSales columns
    RemovedColumns = Table.RemoveColumns(InsertedCustom1,{"PreviousDate", "PreviousDateSales"}),
    //Sort the table in ascending order by the Date column
    SortedRows = Table.Sort(RemovedColumns,{{"Date", Order.Ascending}})
in
    SortedRows

// TiedRanksInput
let
    //Load data from Excel worksheet
    Source = Excel.CurrentWorkbook(){[Name="TiedRanksInput"]}[Content],
    //Declare function to calculate the rank
    Rank = (SalesValue) => Table.RowCount(Table.SelectRows(Source, each [Sales]>SalesValue)) +1,
    //Go back to the original input table
    Custom1 = Source,
    //Add new custom column to show the rank values
    InsertedCustom = Table.AddColumn(Custom1, "Rank", each Rank([Sales]))
in
    InsertedCustom

// DistinctCustomersInput
let
    //Load data from the Excel worksheet
    Source = Excel.CurrentWorkbook(){[Name="DistinctCustomersInput"]}[Content],
    //Set the data type of the Date column to date
    ChangedType = Table.TransformColumnTypes(Source,{{"Date", type date}}),
    //Remove all duplicate combinations of Date and Customer
    DuplicatesRemoved = Table.Distinct(ChangedType, {"Date", "Customer"}),
    //Find the count of rows per date
    GroupedRows = Table.Group(DuplicatesRemoved, {"Date"}, {{"Distinct Customers", each Table.RowCount(_), type number}})
in
    GroupedRows