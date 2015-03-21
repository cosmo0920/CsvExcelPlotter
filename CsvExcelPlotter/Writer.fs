module Writer

open Microsoft.Office.Interop.Excel
open ProcessedCsvType
open CsvProcessing
open Excel.Util

// Run Excel as a visible application
let app = new ApplicationClass(Visible = true)
let sheet1Name = "Sheet1Test"
let leftSpaceWidth = 2
// Create new file and get the first worksheet
let workbook = app.Workbooks.Add(XlWBATemplate.xlWBATWorksheet) 
// Note that worksheets are indexed from one instead of zero
let worksheet = (workbook.Worksheets.[1] :?> Worksheet)

let writeToExcel (processedCsv: ProcessedCsv) =
    let length = processedCsv.RowLength
    let titleLength = processedCsv.Titles.Length
    let titleArea = "C2:" + string('C'+char(titleLength-1))+"2"
    let namesArea = "B3:B"+string(3+length-1)
    let dataArea = "C3:"+string('C'+char(titleLength-1))+string(3+length-1)
    worksheet.Range(titleArea).Value2 <- processedCsv.Titles
    worksheet.Range(namesArea).Value2 <- processedCsv.Names
    worksheet.Range(dataArea).Value2 <- processedCsv.CsvData
    worksheet.Name <- sheet1Name
    let titleRange = worksheet.Range(titleArea)
    let namesRange = worksheet.Range("B2:B"+string(3+length-1))
    let rowsRange = worksheet.Range(dataArea)

    SetPattern titleRange
    SimpleColorGreyFormat titleRange
    SimplePattern namesRange
    SetPattern rowsRange
    SimpleColorLightGrayFormat namesRange
    worksheet.Columns.Range("A:A").ColumnWidth <- leftSpaceWidth

let write (processCsv: ProcessedCsv option) = 
    match processCsv with
    | Some(csv) -> writeToExcel csv
    | None -> do failwith "quit" 