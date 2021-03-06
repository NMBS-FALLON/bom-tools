open System.IO
#I __SOURCE_DIRECTORY__
#I @"../packages"
#r @"NETStandard.Library.NETFRamework/build/net461/lib/netstandard.dll"
#r @"DocumentFormat.OpenXml/lib/net46/DocumentFormat.OpenXml.dll"
#r @"System.IO.Packaging/lib/net46/System.IO.Packaging.dll"
#r @"MicrosoftOfficeCore\lib\net35\Office.dll"
#r @"Microsoft.Office.Interop.Excel\lib\net20\Microsoft.Office.Interop.Excel.dll"
#r @"EPPlus/lib/net40/EPPlus.dll"
#r @"../src/bin/Release/netstandard2.0/bom-tools.dll"
#r "nuget: System.IO.Packaging"
#r "nuget: WindowsBase"


open DESign.BomTools
open DESign.BomTools.Import
open DESign.BomTools.Domain
open DESign.BomTools.Dto
open DESign.SpreadSheetML.Helpers
open DESign.BomTools.NotesToExcel
open DESign.BomTools.AdjustLoads

let bomFileName = @"C:\Users\darien.shannon\code\bom-tools\testBOMs\testbom1.xlsm"
let bomOutPutPath = @"C:\Users\darien.shannon\code\bom-tools\testBOMs\BOM Notes.xlsx"

let getBomInfo () =
    use bom = GetBom bomFileName
    let job = GetJob bom
    use package = CreateBomInfoSheetFromJob job
    using (new FileStream(bomOutPutPath, FileMode.Create)) (fun fs -> package.SaveAs(fs))

let jobWithLc3Loads () =
    use bom = GetBom bomFileName
    let job = GetJob bom
    let newJob = GetSeperatedSeismicLoads job
    newJob.Loads
    |> Seq.toList

jobWithLc3Loads()


                
getBomInfo()