module CsvSchema

open Microsoft.Office.Interop.Excel
open FSharp.Data
open ProcessedCsvType

let rowName = "itemName"
let column1 = "item1"
let column2 = "item2"
let column3 = "item3"
let titles = [| column1; column2; column3 |]