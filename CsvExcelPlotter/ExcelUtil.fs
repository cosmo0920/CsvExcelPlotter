module Excel.Util
open Microsoft.Office.Interop.Excel

// 範囲の周りのセルに罫線(細い線)を設定します。範囲の内側のセルの境界には描画しません。
let SetPattern (range:Range) =
    range.Borders.[XlBordersIndex.xlEdgeLeft].LineStyle <- XlLineStyle.xlContinuous
    range.Borders.[XlBordersIndex.xlEdgeLeft].Weight <- XlBorderWeight.xlThin
    range.Borders.[XlBordersIndex.xlEdgeTop].LineStyle <- XlLineStyle.xlContinuous
    range.Borders.[XlBordersIndex.xlEdgeTop].Weight <- XlBorderWeight.xlThin
    range.Borders.[XlBordersIndex.xlEdgeBottom].LineStyle <- XlLineStyle.xlContinuous
    range.Borders.[XlBordersIndex.xlEdgeBottom].Weight <- XlBorderWeight.xlThin
    range.Borders.[XlBordersIndex.xlEdgeRight].LineStyle <- XlLineStyle.xlContinuous
    range.Borders.[XlBordersIndex.xlEdgeRight].Weight <- XlBorderWeight.xlThin

// 上下左右を罫線(細い線)で囲う
let SimplePattern (range:Range) =
    range.Borders.LineStyle <- XlLineStyle.xlContinuous

// セルを濃い灰色で塗る
let SimpleColorGreyFormat (range: Range) =
    range.Interior.Color <- XlRgbColor.rgbDimGrey;;

// セルを薄い灰色で塗る
let SimpleColorLightGrayFormat (range: Range) =
    range.Interior.Color <- XlRgbColor.rgbLightGrey;;