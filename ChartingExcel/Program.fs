// Learn more about F# at http://fsharp.net
// See the 'F# Tutorial' project for more help.
open Microsoft.Office.Interop.Excel
open System
open FSharp.Data
open System.IO

let csvFile = "datafile.csv"
if not(File.Exists(csvFile)) then
    printfn "Target csv file %s doesn't exist." csvFile; failwith "quit"

// Run Excel as a visible application
let app = new ApplicationClass(Visible = true)
let dataFile = Some(CsvFile.Load("datafile.csv").Cache())
let rowName = "itemName"
let column1 = "item1"
let column2 = "item2"
let column3 = "item3"
let sheet1Name = "Sheet1Test"
let leftSpaceWidth = 2
// Create new file and get the first worksheet
let workbook = app.Workbooks.Add(XlWBATemplate.xlWBATWorksheet) 
// Note that worksheets are indexed from one instead of zero
let worksheet = (workbook.Worksheets.[1] :?> Worksheet)

// 範囲の周りのセルに罫線を設定します。
let SetPattern (range:Range) =
    range.Borders.[XlBordersIndex.xlEdgeLeft].LineStyle <- XlLineStyle.xlContinuous
    range.Borders.[XlBordersIndex.xlEdgeLeft].Weight <- XlBorderWeight.xlThin
    range.Borders.[XlBordersIndex.xlEdgeTop].LineStyle <- XlLineStyle.xlContinuous
    range.Borders.[XlBordersIndex.xlEdgeTop].Weight <- XlBorderWeight.xlThin
    range.Borders.[XlBordersIndex.xlEdgeBottom].LineStyle <- XlLineStyle.xlContinuous
    range.Borders.[XlBordersIndex.xlEdgeBottom].Weight <- XlBorderWeight.xlThin
    range.Borders.[XlBordersIndex.xlEdgeRight].LineStyle <- XlLineStyle.xlContinuous
    range.Borders.[XlBordersIndex.xlEdgeRight].Weight <- XlBorderWeight.xlThin

// 上下左右を罫線で囲う
let SimplePattern (range:Range) =
    range.Borders.LineStyle <- XlLineStyle.xlContinuous

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
                        let rowsArea = "C3:E"+string(3+length-1)
                        worksheet.Range(titleArea).Value2 <- titles
                        worksheet.Range(namesArea).Value2 <- names
                        worksheet.Range(rowsArea).Value2 <- data
                        worksheet.Name <- sheet1Name
                        let titleRange = worksheet.Range(titleArea)
                        let range2 = worksheet.Range("B2:B"+string(3+length-1))
                        let rowsRange = worksheet.Range(rowsArea)

                        let SetBorders = SetPattern titleRange
                        let SetBorders2 = SimplePattern range2
                        let SetBorders3 = SetPattern rowsRange
                        worksheet.Columns.Range("A:A").ColumnWidth <- leftSpaceWidth
    | None -> do failwith "fail csv processing"

processCsv dataFile

[<EntryPoint>]
let main argv = 
    // output to Excel
    let excel = worksheet

    0 // return an integer exit code