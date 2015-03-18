module ProcessedCsvType

open FSharp.Data

type ProcessedCsv = {Titles:string []; Names: string[,]; CsvData: string [,]; RowLength: int}
type TargetRowName = {Name: string}
type TargetColumnNames = {Names: string []}