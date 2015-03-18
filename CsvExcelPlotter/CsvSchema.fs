module CsvSchema

open Microsoft.Office.Interop.Excel
open FSharp.Data
open ProcessedCsvType

let rowName = "itemName"
let column1 = "item1"
let column2 = "item2"
let column3 = "item3"
let titles = [| column1; column2; column3 |]

let readCsvData (csvData: Runtime.CsvFile<CsvRow>) =
    [|for row in csvData.Rows -> [|row.[column1]; row.[column2]; row.[column3];|]|]
let readCsvRowName (csvData: Runtime.CsvFile<CsvRow>) =
    [|for row in csvData.Rows -> [|row.[rowName]; |]|]