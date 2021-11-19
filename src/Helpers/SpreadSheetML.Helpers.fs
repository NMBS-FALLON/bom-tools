module DESign.SpreadSheetML.Helpers

open System
open System.IO
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Spreadsheet
open System.Text.RegularExpressions
open System.Data.Common
open FSharp.Core
open DESign.Helpers

let GetSheetsByPartialName partialName (document : SpreadsheetDocument) =
  let allSheets = document.WorkbookPart.Workbook.Descendants<Sheet>()
  let sheets =
    allSheets |> Seq.filter (fun s -> s.Name.Value.Contains(partialName))
  let workSheetParts =
    sheets
    |> Seq.map
         (fun s ->
         ((document.WorkbookPart.GetPartById(s.Id.Value)) :?> WorksheetPart).Worksheet)
  workSheetParts


let TryGetCellAtColumn (columnName : string) (row : Row) =
  let possibleCell =
    row
    |> Seq.map (fun c -> c :?> Cell)
    |> Seq.filter
         (fun c ->
         c.CellReference.Value = columnName + (string row.RowIndex.Value))
    |> Seq.toList
  match (possibleCell |> Seq.isEmpty) with
  | true -> None
  | false -> Some(possibleCell |> Seq.head)

let (|EmptyCell|FormulaCell|SharedStringCell|BooleanCell|StringCell|Other|) (cell : Cell) =
  if cell = null then
    EmptyCell
  else if cell.CellFormula <> null then
    FormulaCell
  else if cell.DataType = null then
    Other
  else
    match cell.DataType.Value with
    | CellValues.String -> StringCell
    | CellValues.Boolean -> BooleanCell
    | CellValues.SharedString -> SharedStringCell
    | _ -> Other


let TryGetCellValueAtColumnAsString (columnName : string)
    (stringTable : SharedStringTable) (row : Row) =
  row
  |> TryGetCellAtColumn columnName
  |> Option.bind
       (fun c ->
       match c with
       | EmptyCell -> None
       | FormulaCell -> 
          let possibleCellValue = (c.ChildElements |> Seq.filter (fun i -> i.ToString().Contains("CellValue")))
          match possibleCellValue |> Seq.isEmpty with
          | true -> None
          | false -> 
            let value = (possibleCellValue |> Seq.head).InnerText
            match value with
            | Regex @"^ *$" _ -> None
            |  _ -> Some value
       | SharedStringCell ->
          let asXml =
             stringTable.SharedStringTablePart.SharedStringTable.Elements()
             |> Seq.tryItem (FSharp.Core.int.Parse c.InnerText)
          asXml |> Option.map (fun xml -> xml.InnerText)
       | BooleanCell ->
           match c.InnerText with
           | "0" -> Some "FALSE"
           | _ -> Some "TRUE"  
       | StringCell | Other -> 
          match c.InnerText with
          | Regex @"^ *$" _ -> None
          | s -> Some s   )

let inline TryGetCellValueAtColumnWithType< ^T when ^T : (static member Parse : string -> ^T)> (columnName : string) (stringTable : SharedStringTable) (row : Row) =
  let stringValue =
    row |> TryGetCellValueAtColumnAsString columnName stringTable

  match stringValue with
  | Some s ->
      try
        match s with
        | Regex @"^ *$" _ -> Ok None
        | _ -> Ok (Some ((^T : (static member Parse : string -> ^T) s)))
      with
        | _ -> Error (sprintf "Failed while attempting to parse %s as a %A at %s" s typeof< ^T> (columnName + (string row.RowIndex.Value)))
  | None -> Ok None



let GetStringTable (document : SpreadsheetDocument) =
  (document.WorkbookPart.GetPartsOfType<SharedStringTablePart>() |> Seq.head).SharedStringTable

let GetRows firstRowNum lastRowNum (worksheet : Worksheet) =
  (worksheet.Descendants<Row>()
   |> Seq.filter
        (fun r ->
        r.RowIndex.Value >= firstRowNum && r.RowIndex.Value <= lastRowNum))

let GetRow rowNum worksheet =
  worksheet |> GetRows rowNum rowNum |> Seq.head

let TryGetRow rowNum worksheet =
  let possibleRow = worksheet |> GetRows rowNum rowNum
  match possibleRow |> Seq.isEmpty with
  | true -> None
  | false -> Some (possibleRow |> Seq.head)

let GetRowsFromMultipleSheets firstRowNum lastRowNum
    (worksheets : Worksheet seq) =
  worksheets
  |> Seq.map (fun sheet -> sheet |> GetRows firstRowNum lastRowNum)
  |> Seq.concat