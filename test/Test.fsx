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


let bomFileName = @"C:\Users\darien.shannon\code\DESign\bom-tools\testBOMs\testbom1.xlsm"


let printGeneralNotes () =
    use bom = GetBom bomFileName
    let particularNotes = GetGeneralNotes bom
    particularNotes
    |> Seq.iter
        (fun note ->
            note.Notes |> Seq.iter (fun note -> printfn "%s" note))



let printParticularNotes () =
    use bom = GetBom bomFileName
    let particularNotes = GetParticularNotes bom
    particularNotes
    |> Seq.iter
        (fun note ->
            note.Notes |> Seq.iter (fun note -> printfn "%s" note ))

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
    girders |> Seq.iter (fun girder -> printfn "%A" girder)

#time
printGeneralNotes ()
printParticularNotes ()
printLoads ()
printJoists ()
printGirders ()
printGirderExcessInfo ()

