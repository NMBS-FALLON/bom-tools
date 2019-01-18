// Learn more about F# at http://fsharp.org

open System
open DESign.BomTools
open Avalonia
open Avalonia.Controls

[<EntryPoint>]
let main argv =
    let window = new Window()
    let fileName =
        
        let makeFilters (fileDialogFilters : (string * string list) list) =
            let filters = new ResizeArray<FileDialogFilter>()
            for (name, extensions) in fileDialogFilters do
                let filter = new FileDialogFilter()
                filter.Name <- name
                filter.Extensions <- ResizeArray<string> extensions
                filters.Add(filter)
            filters

        let ofd = new OpenFileDialog()
        ofd.Filters <- makeFilters (["Excel Files", [".xlsx"; ".xlsm"]])
        ofd.InitialDirectory <- ""
        ofd.InitialFileName <- ""
        ofd.Title <- ""

        let fileName =
            match ofd.ShowAsync(window) |> Async.AwaitTask |> Async.RunSynchronously with
            | [|file|] -> Some file
            | _ -> None
        fileName

    //let fileName = @"C:\Users\darien.shannon\code\DESign\bom-tools\testBOMs\testbom1.xlsm"
    match fileName with
    | Some fileName -> 
        use bom = Import.GetBom fileName 
        let joists = Import.GetJoists bom
        let printJoists joists =
            joists |> Seq.iter (fun joist -> printfn "%A" joist)
        printJoists joists
    | None -> printfn "No File Selected!"
    Console.ReadLine() |> ignore
    0 // return an integer exit code
