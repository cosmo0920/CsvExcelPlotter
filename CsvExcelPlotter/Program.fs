module Main
open Writer
open System.IO
open FSharp.Data
open CsvProcessing
open Microsoft.Office.Interop.Excel

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
    // Run Excel as a visible application
    let app = new ApplicationClass(Visible = true)
    let sheet1Name = "AddedSheet"
    let leftSpaceWidth = 2
    // Create new file and get the first worksheet
    let workbook = app.Workbooks.Add(XlWBATemplate.xlWBATWorksheet)
    // Note that worksheets are indexed from one instead of zero
    let worksheet = (workbook.Worksheets.[1] :?> Worksheet)

    // insert new work sheet
    let newWorksheet = (workbook.Worksheets.Add() :?> Worksheet)
    // output to Excel
    setExcelStyle newWorksheet sheet1Name leftSpaceWidth
    write newWorksheet csv

    0 // return an integer exit code