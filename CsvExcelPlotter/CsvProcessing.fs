﻿module CsvProcessing

open Microsoft.Office.Interop.Excel
open System
open FSharp.Data
open System.IO
open Excel.Util
open ProcessedCsvType

let rowName = "itemName"
let column1 = "item1"
let column2 = "item2"
let column3 = "item3"

let readCsvData (csvData: Runtime.CsvFile<CsvRow>) =
    [|for row in csvData.Rows do
        yield [|row.GetColumn(column1); row.GetColumn(column2); row.GetColumn(column3)|]|]
let readCsvRowName (csvData: Runtime.CsvFile<CsvRow>) =
    [|for row in csvData.Rows do
        yield [|row.GetColumn(rowName); |]|]

let processCsv csvData =
    match csvData with
    | Some(dataFile) -> let csv = readCsvData dataFile
                        let rowItem = readCsvRowName dataFile
                        let length = csv.Length
                        // Store data in arrays of strings or floats
                        let titles = [| column1; column2; column3 |]
                        let titleLength = titles.Length
                        let names = Array2D.init length 1 (fun i _ -> rowItem.[i].[0])
                        let data = Array2D.init length titleLength (fun i j -> csv.[i].[j])
                        let titleArea = "C2:" + string('C'+char(titleLength-1))+"2"
                        let namesArea = "B3:B"+string(3+length-1)
                        let dataArea = "C3:"+string('C'+char(titleLength-1))+string(3+length-1)
                        Some({Titles = titles; Names = names; CsvData = data; RowLength = length})

    | None -> None