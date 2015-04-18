module Writer

open Microsoft.Office.Interop.Excel
open ProcessedCsvType
open CsvProcessing
open Excel.Util

let writeToExcel (sheet: Worksheet) (processedCsv: ProcessedCsv) =
    let length = processedCsv.RowLength
    let titleLength = processedCsv.Titles.Length
    let titleArea = "C2:" + string('C'+char(titleLength-1))+"2"
    let namesArea = "B3:B"+string(3+length-1)
    let dataArea = "C3:"+string('C'+char(titleLength-1))+string(3+length-1)
    sheet.Range(titleArea).Value2 <- processedCsv.Titles
    sheet.Range(namesArea).Value2 <- processedCsv.Names
    sheet.Range(dataArea).Value2 <- processedCsv.CsvData

    let titleRange = sheet.Range(titleArea)
    let namesRange = sheet.Range("B2:B"+string(3+length-1))
    let rowsRange = sheet.Range(dataArea)

    SetPattern titleRange
    SimpleColorGreyFormat titleRange
    SimplePattern namesRange
    SetPattern rowsRange
    SimpleColorLightGrayFormat namesRange

let setExcelStyle (worksheet: Worksheet)(sheetName: string)(leftSpaceWidth: int) =
    worksheet.Name <- sheetName
    worksheet.Columns.Range("A:A").ColumnWidth <- leftSpaceWidth

let write (worksheet: Worksheet)(processCsv: ProcessedCsv option) =
    match processCsv with
    | Some(csv) -> writeToExcel worksheet csv
    | None -> do failwith "quit" 