#I __SOURCE_DIRECTORY__
#I @"../packages"
#r @"NETStandard.Library.NETFRamework/build/net461/lib/netstandard.dll"
#r @"DocumentFormat.OpenXml/lib/net46/DocumentFormat.OpenXml.dll"
#r @"System.IO.Packaging/lib/net46/System.IO.Packaging.dll"
#r @"WindowsBase"
#r @"../src/bom-tools-library/bin/debug/netstandard2.0/bom-tools.dll"


open DESign.BomTools
open DESign.BomTools.Import
open DESign.BomTools.Domain
open DESign.BomTools.Dto
open DESign.SpreadSheetML.Helpers
open DESign.BomTools.NotesToExcel


let bomFileName = @"C:\Users\darien.shannon\code\bom-tools\testBOMs\testbom1.xlsm"


let printGeneralNotes () =
    use bom = GetBom bomFileName
    let particularNotes = GetGeneralNotes bom
    particularNotes
    |> Seq.iter
        (fun note -> printfn "%s" note.Note )



let printParticularNotes () =
    use bom = GetBom bomFileName
    let particularNotes = GetParticularNotes bom
    particularNotes
    |> Seq.iter
        (fun note -> printfn "%s" note.Note )

let printLoads () =
    use bom = GetBom bomFileName
    let loads = GetLoads bom
    loads |> Seq.iter (fun load -> printfn "%A" load)

let printJoists () =
    use bom = GetBom bomFileName
    let joists = GetJoists bom
    joists |> Seq.iter (fun joist -> printfn "%A" joist)

let printGirders () =
    use bom = GetBom bomFileName
    let girders = GetGirders bom
    girders
    |> Seq.filter (fun girder -> girder.Mark = "G17")
    |> Seq.iter (fun girder -> printfn "%A" girder)

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

let particularNotesByMarkToCsv () =
    use bom = GetBom bomFileName
    let job = GetJob bom
    job |> createParticularNotesSpreadSheet @"C:\Users\darien.shannon\code\bom-tools\test\particularNotesByMark.xlsx"
                

printGeneralNotes ()
printParticularNotes ()
printLoads ()
printJoists ()
printGirders ()
printParticularNotes ()
particularNotesByMarkToCsv()

