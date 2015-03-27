module Main
open Writer
open System.IO
open FSharp.Data
open CsvProcessing

[<EntryPoint>]
let main argv =
#if DEBUG
    let csvFile = __SOURCE_DIRECTORY__ + "\\datafile.csv"
#else
    let csvFile = "datafile.csv"
#endif
    if not(File.Exists(csvFile)) then
        printfn "Target csv file %s doesn't exist." csvFile; failwith "quit"

    let dataFile = Some(CsvFile.Load(csvFile).Cache())
    let csv = processCsv dataFile
    // output to Excel
    write csv

    0 // return an integer exit code