// Learn more about F# at http://fsharp.org

open System
open DESign.BomTools

[<EntryPoint>]
let main argv =
    let test = Console.ReadLine()
    let fileName = @"C:\Users\darien.shannon\code\DESign\bom-tools\testBOMs\testbom1.xlsm"
    use bom = Import.GetBom fileName
    
    let joists = Import.GetJoists bom
    let printJoists joists =
        joists |> Seq.iter (fun joist -> printfn "%A" joist)
    printJoists joists
    0 // return an integer exit code
