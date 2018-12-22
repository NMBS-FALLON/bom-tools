module DESign.BomTools.Import

open System
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Spreadsheet
open FSharp.Core
open DESign.SpreadSheetML.Helpers
open DESign.BomTools.Dto
open DESign.BomTools.Domain

let handleWithFailwith result =
    match result with
    | Ok v -> v
    | Error s -> failwith s

let getLoadSheets (document : SpreadsheetDocument) = document |> GetSheetsByPartialName "L ("
let getGeneralNotesSheets (document : SpreadsheetDocument) =
    document |> GetSheetsByPartialName "P ("
let getParticularNotesSheets (document : SpreadsheetDocument) =
    document |> GetSheetsByPartialName "N ("
let getGirderSheets (document : SpreadsheetDocument) = document |> GetSheetsByPartialName "G ("
let getJoistSheets (document : SpreadsheetDocument) = document |> GetSheetsByPartialName("J (")

type Bom =
    { Document : SpreadsheetDocument
      StringTable : SharedStringTable
      GeneralNoteSheets : Worksheet seq
      ParticularNoteSheets : Worksheet seq
      LoadSheets : Worksheet seq
      JoistSheets : Worksheet seq
      GirderSheets : Worksheet seq }

    interface IDisposable with
        member this.Dispose() = (this.Document :> IDisposable).Dispose()

    member this.TryGetCellValueAtColumnAsString column row =
        TryGetCellValueAtColumnAsString column this.StringTable row
    member inline this.TryGetCellValueAtColumnWithType column row =
        TryGetCellValueAtColumnWithType column this.StringTable row |> handleWithFailwith

let GetBom(bomFileName : string) =
    let document = SpreadsheetDocument.Open(bomFileName, true)
    { Document = document
      StringTable = GetStringTable document
      GeneralNoteSheets = getGeneralNotesSheets document
      ParticularNoteSheets = getParticularNotesSheets document
      LoadSheets = getLoadSheets document
      JoistSheets = getJoistSheets document
      GirderSheets = getGirderSheets document }

let GetGeneralNotes(bom : Bom) =
    let stringTable = bom.StringTable
    let generalNoteSheets = bom.GeneralNoteSheets
    let firstRowNum = 8u
    let lastRowNum = 47u
    let rows = generalNoteSheets |> GetRowsFromMultipleSheets firstRowNum lastRowNum

    let generalNoteDtosWithoutProperIdsAndWithEmptyRows =
        rows
        |> Seq.map
               (fun row ->
               { NoteDto.ID = row |> TryGetCellValueAtColumnAsString "A" stringTable
                 NoteDto.Notes =
                     seq { yield row |> TryGetCellValueAtColumnAsString "B" stringTable } })

    let cleanNotes =
        let noteDtosWithCorrectIds =
            let mutable id = ""
            seq {
                for note in generalNoteDtosWithoutProperIdsAndWithEmptyRows do
                    if note.ID.IsNone then yield { note with ID = Some id }
                    else
                        id <-
                            match note.ID with
                            | Some id -> id
                            | None -> ""
                        yield note
            }

        let cleanNotes =
            noteDtosWithCorrectIds
            |> Seq.groupBy (fun n -> n.ID)
            |> Seq.map (fun (id, noteDtos) ->
                   let notes =
                       seq {
                           for note in noteDtos do
                               yield! note.Notes
                       }

                   let noteDto : NoteDto =
                       { ID = id
                         Notes = notes }

                   Note.Parse noteDto)

        cleanNotes

    cleanNotes

let GetParticularNotes(bom : Bom) =
    let stringTable = bom.StringTable
    let particularNoteSheets = bom.ParticularNoteSheets
    let firstRowNum = 13u
    let lastRowNum = 51u
    let rows = particularNoteSheets |> GetRowsFromMultipleSheets firstRowNum lastRowNum

    let particularNoteDtosWithoutProperIdsAndWithEmptyRows =
        rows
        |> Seq.map
               (fun row ->
               { NoteDto.ID = row |> TryGetCellValueAtColumnAsString "A" stringTable
                 NoteDto.Notes =
                     seq { yield row |> TryGetCellValueAtColumnAsString "B" stringTable } })

    let cleanNotes =
        let noteDtosWithCorrectIds =
            let mutable id = ""
            seq {
                for note in particularNoteDtosWithoutProperIdsAndWithEmptyRows do
                    if note.ID.IsNone then yield { note with ID = Some id }
                    else
                        id <-
                            match note.ID with
                            | Some id -> id
                            | None -> ""
                        yield note
            }

        let cleanNotes =
            noteDtosWithCorrectIds
            |> Seq.groupBy (fun n -> n.ID)
            |> Seq.map (fun (id, noteDtos) ->
                   let notes =
                       seq {
                           for note in noteDtos do
                               yield! note.Notes
                       }

                   let noteDto : NoteDto =
                       { ID = id
                         Notes = notes }

                   Note.Parse noteDto)

        cleanNotes

    cleanNotes

let GetLoads(bom : Bom) =
    let loadSheets = bom.LoadSheets
    let firstRowNum = 14u
    let lastRowNum = 55u
    let rows = loadSheets |> GetRowsFromMultipleSheets firstRowNum lastRowNum

    let loadDtos =
        let loadDtosWithoutProperIdsAndWithEmptyRows =
            rows
            |> Seq.map (fun row ->
                   { LoadDto.ID = row |> bom.TryGetCellValueAtColumnAsString "A"
                     LoadDto.Type = row |> bom.TryGetCellValueAtColumnAsString "B"
                     LoadDto.Category = row |> bom.TryGetCellValueAtColumnAsString "C"
                     LoadDto.Position = row |> bom.TryGetCellValueAtColumnAsString "D"
                     LoadDto.Load1Value = row |> bom.TryGetCellValueAtColumnWithType<float> "F"
                     LoadDto.Load1DistanceFt = row |> bom.TryGetCellValueAtColumnWithType<float> "G"
                     LoadDto.Load1DistanceIn = row |> bom.TryGetCellValueAtColumnWithType<float> "H"
                     LoadDto.Load2Value = row |> bom.TryGetCellValueAtColumnWithType<float> "I"
                     LoadDto.Load2DistanceFt = row |> bom.TryGetCellValueAtColumnWithType<float> "J"
                     LoadDto.Load2DistanceIn = row |> bom.TryGetCellValueAtColumnWithType<float> "K"
                     LoadDto.Ref = row |> bom.TryGetCellValueAtColumnAsString "L"
                     LoadDto.LoadCases = row |> bom.TryGetCellValueAtColumnAsString "M"
                     LoadDto.Remarks = row |> bom.TryGetCellValueAtColumnAsString "N" })

        let cleanLoadDtos =
            let loadsWithCorrectIds =
                let mutable id = ""
                seq {
                    for load in loadDtosWithoutProperIdsAndWithEmptyRows do
                        if load.ID.IsNone then yield { load with ID = Some id }
                        else
                            id <-
                                match load.ID with
                                | Some id ->
                                    id
                                | None -> ""
                            yield load
                }

            let rowsWithLoads = loadsWithCorrectIds |> Seq.filter (fun l -> l.Type.IsSome)
            rowsWithLoads

        cleanLoadDtos
    loadDtos |> Seq.map (fun l -> l |> Load.Parse)

let GetJoists(bom : Bom) =
    let stringTable = bom.StringTable
    let joistSheets = bom.JoistSheets

    let shortJoistSheets, longJoistSheets =
        joistSheets
        |> Seq.toList
        |> List.partition (fun sheet ->
               let valueAtA21 =
                   let row21 = sheet |> GetRow 21u
                   row21 |> TryGetCellValueAtColumnAsString "A" stringTable
               match valueAtA21 with
               | Some s when s.ToUpper().Contains("MARK") -> true
               | _ -> false)

    let rows =
        Seq.append (shortJoistSheets |> GetRowsFromMultipleSheets 23u 40u)
            (longJoistSheets |> GetRowsFromMultipleSheets 16u 45u)

    let joistDtos =
        rows
        |> Seq.map (fun row ->
               { JoistDto.Mark = row |> bom.TryGetCellValueAtColumnAsString "A"
                 JoistDto.Quantity = row |> bom.TryGetCellValueAtColumnWithType "B"
                 JoistDto.JoistSize = row |> bom.TryGetCellValueAtColumnAsString "C"
                 JoistDto.OverallLengthFt = row |> bom.TryGetCellValueAtColumnWithType<float> "D"
                 JoistDto.OverallLengthIn = row |> bom.TryGetCellValueAtColumnWithType<float> "E"
                 JoistDto.TcxlLengthFt = row |> bom.TryGetCellValueAtColumnWithType<float> "F"
                 JoistDto.TcxlLengthIn = row |> bom.TryGetCellValueAtColumnWithType<float> "G"
                 JoistDto.TcxlType = row |> bom.TryGetCellValueAtColumnAsString "H"
                 JoistDto.TcxrLengthFt = row |> bom.TryGetCellValueAtColumnWithType<float> "I"
                 JoistDto.TcxrLengthIn = row |> bom.TryGetCellValueAtColumnWithType<float> "J"
                 JoistDto.TcxrType = row |> bom.TryGetCellValueAtColumnAsString "K"
                 JoistDto.SeatDepthLeft = row |> bom.TryGetCellValueAtColumnWithType<float> "L"
                 JoistDto.SeatDepthRight = row |> bom.TryGetCellValueAtColumnWithType<float> "M"
                 JoistDto.BcxlLengthFt = row |> bom.TryGetCellValueAtColumnWithType<float> "N"
                 JoistDto.BcxlLengthIn = row |> bom.TryGetCellValueAtColumnWithType<float> "O"
                 JoistDto.BcxlType = row |> bom.TryGetCellValueAtColumnAsString "P"
                 JoistDto.BcxrLengthFt = row |> bom.TryGetCellValueAtColumnWithType<float> "Q"
                 JoistDto.BcxrLengthIn = row |> bom.TryGetCellValueAtColumnWithType<float> "R"
                 JoistDto.BcxrType = row |> bom.TryGetCellValueAtColumnAsString "S"
                 JoistDto.PunchedSeatsLeftFt = row |> bom.TryGetCellValueAtColumnWithType<float> "T"
                 JoistDto.PunchedSeatsLeftIn = row |> bom.TryGetCellValueAtColumnWithType<float> "U"
                 JoistDto.PunchedSeatsRightFt =
                     row |> bom.TryGetCellValueAtColumnWithType<float> "V"
                 JoistDto.PunchedSeatsRightIn =
                     row |> bom.TryGetCellValueAtColumnWithType<float> "W"
                 JoistDto.PunchedSeatsGa = row |> bom.TryGetCellValueAtColumnWithType<float> "X"
                 JoistDto.OverallSlope = row |> bom.TryGetCellValueAtColumnWithType<float> "GE"
                 JoistDto.Notes = row |> bom.TryGetCellValueAtColumnAsString "AA" })
    joistDtos
    |> Seq.filter
           (fun joistDto ->
           joistDto.Mark.IsSome || joistDto.Quantity.IsSome || joistDto.JoistSize.IsSome)
    |> Seq.map (fun joistDto -> Joist.Parse joistDto)

let GetGirders(bom : Bom) =
    let stringTable = bom.StringTable
    let girderSheets = bom.GirderSheets

    let shortGirderSheets, longGirderSheets =
        girderSheets
        |> Seq.toList
        |> List.partition (fun sheet ->
               let valueAtA26 =
                   let row26 = sheet |> GetRow 26u
                   row26 |> TryGetCellValueAtColumnAsString "A" stringTable
               match valueAtA26 with
               | Some s when s.ToUpper().Contains("MARK") -> true
               | _ -> false)

    let rows =
        Seq.append (shortGirderSheets |> GetRowsFromMultipleSheets 28u 45u)
            (longGirderSheets |> GetRowsFromMultipleSheets 14u 45u)

    let girderExcessInfoDtos =
        let girderSheets = bom.GirderSheets
        let firstRowNum = 14u
        let lastRowNum = 45u
        let rows = girderSheets |> GetRowsFromMultipleSheets firstRowNum lastRowNum

        let girderExcessInfoDtos =
            let girderDtosWithoutProperIdsAndWithEmptyRows =
                rows
                |> Seq.map (fun row ->
                       { GirderExcessInfoLine.Mark = row |> bom.TryGetCellValueAtColumnAsString "AB"
                         NearFarBoth = row |> bom.TryGetCellValueAtColumnAsString "AC"
                         AFt = row |> bom.TryGetCellValueAtColumnWithType<float> "AD"
                         AIn = row |> bom.TryGetCellValueAtColumnWithType<float> "AE"
                         PanelQuantity = row |> bom.TryGetCellValueAtColumnWithType<int> "AF"
                         PanelLengthFt = row |> bom.TryGetCellValueAtColumnWithType<float> "AG"
                         PanelLengthIn = row |> bom.TryGetCellValueAtColumnWithType<float> "AH"
                         BFt = row |> bom.TryGetCellValueAtColumnWithType<float> "AI"
                         BIn = row |> bom.TryGetCellValueAtColumnWithType<float> "AJ"
                         Load = row |> bom.TryGetCellValueAtColumnWithType<float> "AK"
                         AdditionalJoistLoads =
                             seq {
                                 yield { LocationFt =
                                             row |> bom.TryGetCellValueAtColumnAsString "AR"
                                         LocationIn =
                                             row |> bom.TryGetCellValueAtColumnAsString "AS"
                                         LoadValue =
                                             row |> bom.TryGetCellValueAtColumnWithType<float> "AT" }
                                 yield { LocationFt =
                                             row |> bom.TryGetCellValueAtColumnAsString "AV"
                                         LocationIn =
                                             row |> bom.TryGetCellValueAtColumnAsString "AW"
                                         LoadValue =
                                             row |> bom.TryGetCellValueAtColumnWithType<float> "AX" }
                                 yield { LocationFt =
                                             row |> bom.TryGetCellValueAtColumnAsString "AZ"
                                         LocationIn =
                                             row |> bom.TryGetCellValueAtColumnAsString "BA"
                                         LoadValue =
                                             row |> bom.TryGetCellValueAtColumnWithType<float> "BB" }
                                 yield { LocationFt =
                                             row |> bom.TryGetCellValueAtColumnAsString "BD"
                                         LocationIn =
                                             row |> bom.TryGetCellValueAtColumnAsString "BE"
                                         LoadValue =
                                             row |> bom.TryGetCellValueAtColumnWithType<float> "BF" }
                             } })

            let cleanLoadDtos =
                let loadsWithCorrectIds =
                    let mutable mark = ""
                    seq {
                        for girderLine in girderDtosWithoutProperIdsAndWithEmptyRows do
                            if girderLine.Mark.IsNone then
                                yield { girderLine with Mark = Some mark }
                            else
                                mark <-
                                    match girderLine.Mark with
                                    | Some mark ->
                                        mark
                                    | None ->
                                        ""
                                yield girderLine
                    }
                loadsWithCorrectIds

            cleanLoadDtos
        girderExcessInfoDtos

    let getPanelLocations (lines : GirderExcessInfoLine seq) =
        let lines = lines |> Seq.toList

        let rec getPanelLocations panelLocations lines =
            match lines with
            | [] -> panelLocations
            | head :: tail ->
                let firstPanel =
                    (head.AFt |> Option.defaultValue 0.0)
                    + (head.AIn |> Option.defaultValue 0.0) / 12.0
                let numPanels = head.PanelQuantity |> Option.defaultValue 0
                let panelSpacing =
                    (head.PanelLengthFt |> Option.defaultValue 0.0)
                    + (head.PanelLengthIn |> Option.defaultValue 0.0) / 12.0

                let newLines =
                    let mutable panelLocation = firstPanel
                    [ for i in 0..(numPanels - 1) do
                          panelLocation <- panelLocation + panelSpacing
                          yield panelLocation ]
                getPanelLocations (firstPanel :: (panelLocations |> List.append newLines)) tail
        getPanelLocations [] lines |> List.filter (fun panelLocation -> panelLocation <> 0.0)

    let getAdditionalJoistsOnGirders (lines : GirderExcessInfoLine seq) =
        lines
        |> Seq.map (fun l -> l.AdditionalJoistLoads)
        |> Seq.concat
        |> Seq.filter (fun l -> l.LoadValue.IsSome || l.LocationFt.IsSome || l.LocationIn.IsSome)

    let panelLocationDict =
        let groupedPanelLocations =
            girderExcessInfoDtos
            |> Seq.groupBy (fun l -> l.Mark)
            |> Seq.map (fun (mark, lines) -> mark, getPanelLocations lines)
            |> Map.ofSeq
        groupedPanelLocations

    let getPanelLocations mark = panelLocationDict |> Map.tryFind mark

    let girderDtos =
        rows
        |> Seq.map
               (fun row ->
               let mark = row |> bom.TryGetCellValueAtColumnAsString "A"
               { GirderDto.Mark = mark
                 GirderDto.Quantity = row |> bom.TryGetCellValueAtColumnWithType "B"
                 GirderDto.GirderSize = row |> bom.TryGetCellValueAtColumnAsString "C"
                 GirderDto.OverallLengthFt = row |> bom.TryGetCellValueAtColumnWithType<float> "D"
                 GirderDto.OverallLengthIn = row |> bom.TryGetCellValueAtColumnWithType<float> "E"
                 GirderDto.TcWidth = row |> bom.TryGetCellValueAtColumnWithType<float> "F"
                 GirderDto.TcxlLengthFt = row |> bom.TryGetCellValueAtColumnWithType<float> "G"
                 GirderDto.TcxlLengthIn = row |> bom.TryGetCellValueAtColumnWithType<float> "H"
                 GirderDto.TcxlType = row |> bom.TryGetCellValueAtColumnAsString "I"
                 GirderDto.TcxrLengthFt = row |> bom.TryGetCellValueAtColumnWithType<float> "J"
                 GirderDto.TcxrLengthIn = row |> bom.TryGetCellValueAtColumnWithType<float> "K"
                 GirderDto.TcxrType = row |> bom.TryGetCellValueAtColumnAsString "L"
                 GirderDto.SeatDepthLeft = row |> bom.TryGetCellValueAtColumnWithType<float> "M"
                 GirderDto.SeatDepthRight = row |> bom.TryGetCellValueAtColumnWithType<float> "N"
                 GirderDto.BcxlLengthFt = row |> bom.TryGetCellValueAtColumnWithType<float> "O"
                 GirderDto.BcxlLengthIn = row |> bom.TryGetCellValueAtColumnWithType<float> "P"
                 GirderDto.BcxrLengthFt = row |> bom.TryGetCellValueAtColumnWithType<float> "Q"
                 GirderDto.BcxrLengthIn = row |> bom.TryGetCellValueAtColumnWithType<float> "R"
                 GirderDto.PunchedSeatsLeftFt =
                     row |> bom.TryGetCellValueAtColumnWithType<float> "S"
                 GirderDto.PunchedSeatsLeftIn =
                     row |> bom.TryGetCellValueAtColumnWithType<float> "T"
                 GirderDto.PunchedSeatsRightFt =
                     row |> bom.TryGetCellValueAtColumnWithType<float> "U"
                 GirderDto.PunchedSeatsRightIn =
                     row |> bom.TryGetCellValueAtColumnWithType<float> "V"
                 GirderDto.PunchedSeatsGa = row |> bom.TryGetCellValueAtColumnWithType<float> "W"
                 GirderDto.OverallSlope = row |> bom.TryGetCellValueAtColumnWithType<float> "GE"
                 GirderDto.Notes = row |> bom.TryGetCellValueAtColumnAsString "Z"
                 GirderDto.NumKbRequired = row |> bom.TryGetCellValueAtColumnWithType<int> "AA"
                 GirderDto.ExcessInfo = girderExcessInfoDtos
                 GirderDto.PanelLocations = getPanelLocations mark
                 GirderDto.AdditionalJoistLoads =
                     girderExcessInfoDtos |> getAdditionalJoistsOnGirders })
    girderDtos
    |> Seq.filter
           (fun girderDto ->
           girderDto.Mark.IsSome || girderDto.Quantity.IsSome || girderDto.GirderSize.IsSome)
    |> Seq.map (fun girderDto -> Girder.Parse girderDto)


