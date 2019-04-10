#I __SOURCE_DIRECTORY__
#I @"../packages"
#r @"NETStandard.Library.NETFRamework/build/net461/lib/netstandard.dll"
#r @"DocumentFormat.OpenXml/lib/net46/DocumentFormat.OpenXml.dll"
#r @"System.IO.Packaging/lib/net46/System.IO.Packaging.dll"
#r @"EPPlus/lib/net40/EPPlus.dll"
#r @"WindowsBase"
#r @"../src/bom-tools-library/bin/debug/netstandard2.0/bom-tools.dll"


open DESign.BomTools
open DESign.BomTools.Import
open DESign.BomTools.Domain
open DESign.BomTools.Dto
open DESign.SpreadSheetML.Helpers
open DESign.BomTools.NotesToExcel


let bomFileName = @"C:\Users\darien.shannon\code\bom-tools\testBOMs\testbom1.xlsm"

let getBomInfo () =
    use bom = GetBom bomFileName
    let job = GetJob bom
    createBomInfoSheet job @"C:\Users\darien.shannon\code\bom-tools\test\bomInfo.xlsx"
                
getBomInfo()