module DESign.BomTools.NotesToExcel
 
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Spreadsheet
open DocumentFormat.OpenXml
open DESign.BomTools.Domain


let getParticularNotesByMark (job : Job) =
    let joistNotes =
        job.Joists
        |> Seq.map
            (fun joist ->
                let notes = job |> Job.getJoistParticularNotes joist
                notes
                |> Seq.map
                    (fun note -> joist.Mark, note.Note))
        |> Seq.concat

    let girderNotes =
        job.Girders
        |> Seq.map
            (fun girder ->
                let notes = job |> Job.getGirderParticularNotes girder
                notes
                |> Seq.map
                    (fun note -> girder.Mark, note.Note))
       |> Seq.concat

    let allNotes = [girderNotes;joistNotes] |> Seq.concat
    allNotes

let getParticularNotesByNote (job : Job) =
    let joistNotes =
        job.Joists
        |> Seq.map
            (fun joist ->
                let notes = job |> Job.getJoistParticularNotes joist
                notes
                |> Seq.map
                    (fun note -> (note, joist.Mark)))
        |> Seq.concat
    let girderNotes =
        job.Girders
        |> Seq.map
            (fun girder ->
                let notes = job |> Job.getGirderParticularNotes girder
                notes
                |> Seq.map
                    (fun note -> (note, girder.Mark)))
        |> Seq.concat
    let allNotes =
        [girderNotes;joistNotes]
        |> Seq.concat
        |> Seq.groupBy (fun (note, mark) -> note)
        |> Seq.map
            (fun (note, notesAndMarks) ->
                let marks =
                    notesAndMarks
                    |> Seq.map (fun(_, mark) -> mark)
                (note, marks))
        |> Seq.sortBy (fun ((note, _) : (Note*_)) -> note.ID.ToCharArray())
    allNotes

let createCellReference (header:string) (index:int) =
    StringValue(header + string(index))

let createTextCell text (header:string) (index:int) =
    let cell = new Cell()
    cell.CellReference <- createCellReference header index
    cell.DataType <- EnumValue(CellValues.InlineString)
    let inlineString = new InlineString()
    let t = new Text(Text = text)
    t |> inlineString.AppendChild |> ignore
    inlineString |> cell.AppendChild |> ignore
    cell :> OpenXmlElement

let particularNotesByMarkAsSheetData (job : Job) =
    let particularNotesByMark = job |> getParticularNotesByMark
    let sheetData = new SheetData()
    let headerRow = new Row(RowIndex = UInt32Value(uint32(1)))
    createTextCell "Mark" "A" 1 |> headerRow.Append
    createTextCell "Note" "B" 1 |> headerRow.Append
    (headerRow :> OpenXmlElement) |> sheetData.AppendChild |> ignore
    particularNotesByMark
    |> Seq.iteri
        (fun i (mark, note) ->
            let row = new Row(RowIndex = UInt32Value(uint32(i + 2)))
            let markCell = createTextCell mark "A" (i + 2)
            let noteCell = createTextCell note "B" (i + 2)
            markCell |> row.Append
            noteCell |> row.Append
            (row :> OpenXmlElement) |> sheetData.AppendChild |> ignore )
    sheetData

 
    
    
let createSpreadsheet (filepath:string) ( listOfSheetNamesAndSheetData: (string*SheetData) list) =
    using (SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook))
        (fun spreadsheetDocument ->
            let workbookPart = spreadsheetDocument.AddWorkbookPart(Workbook = new Workbook())
            let sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets())
            listOfSheetNamesAndSheetData
            |> List.iteri
                (fun i (sheetName, sheetData) ->                  
                    let worksheetPart = workbookPart.AddNewPart<WorksheetPart>()
                    worksheetPart.Worksheet <- new Worksheet(sheetData:> OpenXmlElement)
                    let sheet = new Sheet( Id = StringValue(spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart)),
                                           SheetId = UInt32Value(uint32 i),
                                           Name = StringValue(sheetName))
                    [sheet :> OpenXmlElement] |> sheets.Append ))

let createParticularNotesSpreadSheet (outputPath : string) (job : Job) =
    let particularNotesAsSheetData = job |> particularNotesByMarkAsSheetData
    createSpreadsheet
        outputPath
        ["Particular Notes By Mark", particularNotesAsSheetData
         "Particular Notes By Mark 2", particularNotesAsSheetData]
      
    

