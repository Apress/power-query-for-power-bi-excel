// ListData
let
    //Load data from the Excel worksheet
    Source = Excel.CurrentWorkbook(){[Name="ListData"]}[Content],
    //Add index column
    InsertedIndex = Table.AddIndexColumn(Source,"Index"),
    //Calculate row type as a number
    InsertedCustom = Table.AddColumn(InsertedIndex, "RowType", each Number.Mod([Index],3)),
    //Convert the row type number to a text value
    InsertedCustom1 = Table.AddColumn(InsertedCustom, "RowTypeText", each if [RowType]=0 then "Name" else if [RowType]=1 then "Gender" else "Country"),
    //Identify each customer record
    InsertedCustom2 = Table.AddColumn(InsertedCustom1, "CustomerID", each Number.IntegerDivide([Index], 3)),
    //Remove columns not needed for output
    RemovedColumns = Table.RemoveColumns(InsertedCustom2,{"Index", "RowType"}),
    //Pivot the table
    Custom1 = Table.Pivot(RemovedColumns, {"Name", "Gender", "Country"}, "RowTypeText", "Data")
in
    Custom1

// ThisYearCustomers
let
    Source = Excel.CurrentWorkbook(){[Name="ThisYearCustomers"]}[Content]
in
    Source

// LastYearCustomers
let
    Source = Excel.CurrentWorkbook(){[Name="LastYearCustomers"]}[Content]
in
    Source

// Customers
let
    //Append the LastYearCustomers table to the ThisYearCustomers table
    Source = Table.Combine({ThisYearCustomers,LastYearCustomers}),
    //Merge the two columns into one
    MergedColumns = Table.CombineColumns(Source,{"Customers This Year", "Customers Last Year"},Combiner.CombineTextByDelimiter("", QuoteStyle.None),"Merged"),
    //Remove duplicate customers
    DuplicatesRemoved = Table.Distinct(MergedColumns),
    //Rename the only column in the table to Customer
    RenamedColumns = Table.RenameColumns(DuplicatesRemoved,{{"Merged", "Customer"}}),
    //Merge this query with the ThisYearCustomers table
    Merge = Table.NestedJoin(RenamedColumns,{"Customer"},ThisYearCustomers,{"Customers This Year"},"NewColumn"),
    //Aggregate the resulting column of tables by Count (Not Blank)
    #"Aggregate NewColumn" = Table.AggregateTableColumn(Merge, "NewColumn", {{"Customers This Year", List.NonNullCount, "Count (Not Blank) of NewColumn.Customers This Year"}}),
    //Rename the new column to ThisYear
    RenamedColumns1 = Table.RenameColumns(#"Aggregate NewColumn",{{"Count (Not Blank) of NewColumn.Customers This Year", "ThisYear"}}),
    //Merge this query with the LastYearCustomers table
    Merge1 = Table.NestedJoin(RenamedColumns1,{"Customer"},LastYearCustomers,{"Customers Last Year"},"NewColumn"),
    //Aggregate the resulting column of tables by Count (Not Blank)
    #"Aggregate NewColumn1" = Table.AggregateTableColumn(Merge1, "NewColumn", {{"Customers Last Year", List.NonNullCount, "Count (Not Blank) of NewColumn.Customers Last Year"}}),
    //Rename the new column to LastYear
    RenamedColumns2 = Table.RenameColumns(#"Aggregate NewColumn1",{{"Count (Not Blank) of NewColumn.Customers Last Year", "LastYear"}}),
    //Use the ThisYear and LastYear columns to classify each customer
    InsertedCustom = Table.AddColumn(RenamedColumns2, "Classification", each if [ThisYear]=1 and [LastYear]=1 then "Returning" else if [ThisYear]=1 and [LastYear]=0 then "New" else "Lost"),
    //Remove unwanted columns
    RemovedColumns = Table.RemoveColumns(InsertedCustom,{"ThisYear", "LastYear"})
in
    RemovedColumns

// Date
let
    //Create a list of 365 dates starting from January 1st 2014
    Source = List.Dates(#date(2014,1,1), 365, #duration(1,0,0,0) ),
    //Turn the list into a table
    TableFromList = Table.FromList(Source, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    //Rename the only column in the table to Date
    RenamedColumns = Table.RenameColumns(TableFromList,{{"Column1", "Date"}}),
    //Change the type of the column to Date
    ChangedType = Table.TransformColumnTypes(RenamedColumns,{{"Date", type date}}),
    //Duplicate the Date column
    DuplicatedColumn = Table.DuplicateColumn(ChangedType, "Date", "Copy of Date"),
    //Rename the duplicated column to Year
    RenamedColumns1 = Table.RenameColumns(DuplicatedColumn,{{"Copy of Date", "Year"}}),
    //Convert the dates in the Year column to years
    TransformedColumn = Table.TransformColumns(RenamedColumns1,{{"Year", Date.Year}}),
    //Add a custom column containing month names
    InsertedCustom = Table.AddColumn(TransformedColumn, "Month", each Date.ToText([Date], "MMMM")),
    //Add a custom column containing day names
    InsertedCustom1 = Table.AddColumn(InsertedCustom, "DayName", each Date.ToText([Date], "dddd")),
    //Add a custom column containing week numbers
    InsertedCustom2 = Table.AddColumn(InsertedCustom1, "Week", each Date.WeekOfYear([Date]))
in
    InsertedCustom2

// DateFunction
let
    DateFunction = (StartDate as date, EndDate as date) as table =>
let
    //Find the number of dates between the start date and end date
    NumberOfDays = Duration.Days(EndDate-StartDate)+1,
    //Create a list of dates starting from the start date
    Source = List.Dates(StartDate, NumberOfDays, #duration(1,0,0,0) ),
    //Turn the list into a table
    TableFromList = Table.FromList(Source, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    //Rename the only column in the table to Date
    RenamedColumns = Table.RenameColumns(TableFromList,{{"Column1", "Date"}}),
    //Change the type of the column to Date
    ChangedType = Table.TransformColumnTypes(RenamedColumns,{{"Date", type date}}),
    //Duplicate the Date column
    DuplicatedColumn = Table.DuplicateColumn(ChangedType, "Date", "Copy of Date"),
    //Rename the duplicated column to Year
    RenamedColumns1 = Table.RenameColumns(DuplicatedColumn,{{"Copy of Date", "Year"}}),
    //Convert the dates in the Year column to years
    TransformedColumn = Table.TransformColumns(RenamedColumns1,{{"Year", Date.Year}}),
    //Add a custom column containing month names
    InsertedCustom = Table.AddColumn(TransformedColumn, "Month", each Date.ToText([Date], "MMMM")),
    //Add a custom column containing day names
    InsertedCustom1 = Table.AddColumn(InsertedCustom, "DayName", each Date.ToText([Date], "dddd")),
    //Add a custom column containing week numbers
    InsertedCustom2 = Table.AddColumn(InsertedCustom1, "Week", each Date.WeekOfYear([Date]))
in
    InsertedCustom2
in
    DateFunction

// StockPrice
let
    //Load data from Excel
    Source = Excel.CurrentWorkbook(){[Name="StockPrice"]}[Content],
    //Set column data types
    ChangedType = Table.TransformColumnTypes(Source,{{"Time", type time}, {"Stock Price", type number}}),
    //Is Stock Price greater than or equal to 50?
    InsertedCustom = Table.AddColumn(ChangedType, "Above50", each [Stock Price]>=50),
    //Aggregate time ranges
    GroupedRows = Table.Group(InsertedCustom, {"Above50"}, {{"Start", each List.Min([Time]), type time}, {"EndTemp", each List.Max([Time]), type time}, {"Minutes", each Table.RowCount(_), type number}},GroupKind.Local),
    //Add one minute to the values in the EndTemp column
    InsertedCustom1 = Table.AddColumn(GroupedRows, "End", each [EndTemp] + #duration(0,0,1,0)),
    //Remove the EndTemp column
    RemovedColumns = Table.RemoveColumns(InsertedCustom1,{"EndTemp"}),
    //Move the End column in between Start and Minutes
    ReorderedColumns = Table.ReorderColumns(RemovedColumns,{"Above50", "Start", "End", "Minutes"}),
    //Filter to show only ranges where stock price is greater than or equal to 50
    FilteredRows = Table.SelectRows(ReorderedColumns, each ([Above50] = true)),
    //Sort by Minutes in descending order
    SortedRows = Table.Sort(FilteredRows,{{"Minutes", Order.Descending}}),
    //Keep first row of the table
    KeptFirstRows = Table.FirstN(SortedRows,1)
in
    KeptFirstRows