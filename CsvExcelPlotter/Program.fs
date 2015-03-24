module Main
open Writer
open System.IO
open FSharp.Data
open CsvProcessing

[<EntryPoint>]
let main argv = 
    let csvFile = "datafile.csv"
    if not(File.Exists(csvFile)) then
        printfn "Target csv file %s doesn't exist." csvFile; failwith "quit"

    let dataFile = Some(CsvFile.Load(csvFile).Cache())
    let csv = processCsv dataFile
    // output to Excel
    write csv
    ignore newWorksheet

    0 // return an integer exit code