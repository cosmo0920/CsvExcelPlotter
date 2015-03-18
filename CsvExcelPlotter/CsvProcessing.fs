module CsvProcessing

open Microsoft.Office.Interop.Excel
open FSharp.Data
open ProcessedCsvType
open CsvSchema

let processCsv csvData =
    match csvData with
    | Some(dataFile) -> let csv = readCsvData dataFile
                        let rowItem = readCsvRowName dataFile
                        let length = csv.Length
                        // Store data in arrays of strings or floats
                        let titleLength = titles.Length
                        let names = Array2D.init length 1 (fun i _ -> rowItem.[i].[0])
                        let data = Array2D.init length titleLength (fun i j -> csv.[i].[j])
                        let titleArea = "C2:" + string('C'+char(titleLength-1))+"2"
                        let namesArea = "B3:B"+string(3+length-1)
                        let dataArea = "C3:"+string('C'+char(titleLength-1))+string(3+length-1)
                        Some({Titles = titles; Names = names; CsvData = data; RowLength = length})

    | None -> None