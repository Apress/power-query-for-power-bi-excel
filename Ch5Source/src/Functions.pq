// SimpleFunction
let
    //Define a function that multiplies two numbers then adds one
    ExampleFunction = (x,y) => (x * y) +1,
    //Call the function and pass 3 and 4 as parameters
    QueryOutput = ExampleFunction(3,4)
in
    QueryOutput

// FunctionWithTypesAndOptional
let
    //Define a function that multiplies two numbers then adds one
    ExampleFunction = (x as number,y as number, optional z as number) as number => (x * y) + (if z=null then 1 else z),
    //Call the function and pass 3 and 4 as parameters
    QueryOutput = ExampleFunction(3,4)
in
    QueryOutput

// EachExample
let
    //Define a function that multiplies a number by 2
    ExampleFunction = each _ * 2,
    //Call the function and pass 3 as a parameter
    QueryOutput = ExampleFunction(3)
in
    QueryOutput

// TransformColumnsExample
let
    //Load data from source table
    Source = Excel.CurrentWorkbook(){[Name="FruitSales"]}[Content],
    //Multiply each value in the Sales column by 2
    SalesTimesTwo = Table.TransformColumns(Source, {"Sales", each _ * 2})
in
    SalesTimesTwo

// MultiplyThenAddOne
let
    MyFunction = (x,y) => (x * y) + 1
in
    MyFunction

// FruitSales
let
    //Load data from source table
    Source = Excel.CurrentWorkbook(){[Name="FruitSales"]}[Content],
    //Call function in a custom column
    InsertedCustom = Table.AddColumn(Source, "FunctionCallExample", each MultiplyThenAddOne([Sales],[Units]))
in
    InsertedCustom

// MultiplyThenAddOneV2
let
    //define function that will return the value of the inner let expression
    MyFunction = (x,y) => 
      let
          //multiply x and y
          Step1 = x * y,
          //add one to the result
          Step2 = Step1 + 1
      in
          Step2
in
    MyFunction

// RecursiveFunction
let
    //define a recursive function with one parameter
    DoubleToOneHundred = (x) =>
      //if the parameter is greater than one hundred, return the parameter
      if x > 100
       then x
       //otherwise call the function again with the original parameter
       //multiplied by two
       else @DoubleToOneHundred(x*2),
    //call the function with the value 4
    //the output is 128
    FunctionOutput = DoubleToOneHundred(4)
in
    FunctionOutput