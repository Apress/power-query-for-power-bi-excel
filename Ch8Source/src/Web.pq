// GetWeather
let
    GetWeather = (Postcode1 as text) as table =>
let
    Source = Web.Page(Web.Contents("http://www.bbc.co.uk/weather/" & Postcode1)),
    Data0 = Source{0}[Data],
    ChangedType = Table.TransformColumnTypes(Data0,{{"Day", type text}, {"Weather", type text}, {"Maximum Day Temperature", type text}, {"Minimum Night Temperature", type text}, {"Wind Direction and Speed", type text}})
in
    ChangedType
in
    GetWeather

// Palaces
let
    //Load data from Excel worksheet
    Source = Excel.CurrentWorkbook(){[Name="Palaces"]}[Content],
    //Split post codes by space into two columns
    SplitColumnDelimiter = Table.SplitColumn(Source,"Postcode",Splitter.SplitTextByDelimiter(" "),{"Postcode.1", "Postcode.2"}),
    //Set the type of each column to text
    ChangedType = Table.TransformColumnTypes(SplitColumnDelimiter,{{"Postcode.1", type text}, {"Postcode.2", type text}}),
    //Remove the Postcode.2 column
    RemovedColumns = Table.RemoveColumns(ChangedType,{"Postcode.2"}),
    //Change the values in Postcode.1 to lowercase
    TransformedColumn = Table.TransformColumns(RemovedColumns,{{"Postcode.1", Text.Lower}}),
    //Call the GetWeather() function for each row
    InsertedCustom = Table.AddColumn(TransformedColumn, "Forecast", each GetWeather([Postcode.1])),
    //Expand the table values returned by the GetWeather() function
    #"Expand Forecast" = Table.ExpandTableColumn(InsertedCustom, "Forecast", {"Day", "Weather", "Maximum Day Temperature", "Minimum Night Temperature", "Wind Direction and Speed"}, {"Forecast.Day", "Forecast.Weather", "Forecast.Maximum Day Temperature", "Forecast.Minimum Night Temperature", "Forecast.Wind Direction and Speed"})
in
    #"Expand Forecast"

// RouteData
let
    Source = Excel.CurrentWorkbook(){[Name="RouteData"]}[Content],
    Custom1 = Web.Contents("http://dev.virtualearth.net/REST/V1/Routes/Driving", [Query=[#"wp.0"=Source{0}[Start], #"wp.1"=Source{0}[End], #"key"="Ajv0XBkMAN8GjO4wyg2yqHWf2j8dHXxmJW2seTcSf8TshlWF08a6EGsMK2f1Kllz"]]),
    ImportedJSON = Json.Document(Custom1),
    resourceSets = ImportedJSON[resourceSets],
    resourceSets1 = resourceSets{0},
    resources = resourceSets1[resources],
    resources1 = resources{0},
    travelDistance = resources1[travelDistance]
in
    travelDistance