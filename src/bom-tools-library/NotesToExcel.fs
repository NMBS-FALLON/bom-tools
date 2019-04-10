module DESign.BomTools.NotesToExcel
 
open DESign.BomTools.Domain
open OfficeOpenXml
open System.IO
open DocumentFormat.OpenXml.Spreadsheet



let getParticularNotesByMark (job : Job) =
    let joistNotes =
        job.Joists
        |> Seq.collect
            (fun joist ->
                job
                |> Job.getJoistParticularNotes joist
                |> Seq.map
                    (fun note -> joist.Mark, note.Note, note.ID))

    let girderNotes =
        job.Girders
        |> Seq.collect
            (fun girder ->
                job |> Job.getGirderParticularNotes girder
                |> Seq.map
                    (fun note -> girder.Mark, note.Note, note.ID))

    let allNotes = [girderNotes;joistNotes] |> Seq.concat
    allNotes



let CreateBomInfoSheetFromJob (job : Job) =
    let package = new ExcelPackage()
    let notesByMarkSheet = package.Workbook.Worksheets.Add("Notes By Mark")
    notesByMarkSheet.Cells.[1,1].Value <- "Mark"
    notesByMarkSheet.Cells.[1,2].Value <- "Note"
    notesByMarkSheet.Cells.[1,3].Value <- "Note Id"
    let notesByMark = job |> getParticularNotesByMark
    notesByMark
    |> Seq.iteri
        (fun i (mark, note, noteId) ->
            notesByMarkSheet.Cells.[i + 2, 1].Value <- mark
            notesByMarkSheet.Cells.[i + 2, 2].Value <- note
            notesByMarkSheet.Cells.[i + 2, 3].Value <- noteId)

    let lastRow = notesByMarkSheet.Dimension.End.Row
    let tableRange = notesByMarkSheet.Cells.[1, 1, lastRow, 3]
    let tableName = "tblBomInfo"
    let table =notesByMarkSheet.Tables.Add(tableRange, tableName)
    table.TableStyle <- Table.TableStyles.Light8
    tableRange.AutoFitColumns()

    package
    




