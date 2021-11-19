module DESign.BomTools.Vulcraft.Import.FromExcel

open System
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Spreadsheet
open FSharp.Core
open DESign.SpreadSheetML.Helpers
open DESign.BomTools.Vulcraft.Import.Dtos
open DESign.Helpers


let getGirderSheets (document : SpreadsheetDocument) = document |> GetSheetsByPartialName "GIRDER ("
let getJoistSheets (document : SpreadsheetDocument) = document |> GetSheetsByPartialName "JOIST ("
let getUpliftSheets (document : SpreadsheetDocument) = document |> GetSheetsByPartialName "UPLIFT ("

type Bom =
    { Document : SpreadsheetDocument
      StringTable : SharedStringTable
      GirderSheets : Worksheet seq
      JoistSheets : Worksheet seq
      UpliftSheets : Worksheet seq }

    interface IDisposable with
        member this.Dispose() = (this.Document :> IDisposable).Dispose()

    member this.TryGetCellValueAtColumnAsString column row =
        TryGetCellValueAtColumnAsString column this.StringTable row

let GetBom(bomFileName : string) =
    let document = SpreadsheetDocument.Open(bomFileName, true)
    { Document = document
      StringTable = GetStringTable document
      GirderSheets = getGirderSheets document
      JoistSheets = getJoistSheets document
      UpliftSheets = getUpliftSheets document}


let GetJoistsFromOneSheet (bom : Bom) (joistSheet : Worksheet) =
    let stringTable = bom.StringTable
    [ for markLine in 0u..4u do 
        
        let firstRowOfInfo = joistSheet |> GetRow (12u + markLine*2u)
        let mark = firstRowOfInfo |> bom.TryGetCellValueAtColumnAsString "A"
        let quantity = firstRowOfInfo |> bom.TryGetCellValueAtColumnAsString "I"
        let designationDepth = firstRowOfInfo |> bom.TryGetCellValueAtColumnAsString "O"
        let designationType = firstRowOfInfo |> bom.TryGetCellValueAtColumnAsString "X"
        let designationLoading = firstRowOfInfo |> bom.TryGetCellValueAtColumnAsString "AC"
        let overallLengthFt = firstRowOfInfo |> bom.TryGetCellValueAtColumnAsString "AU"
        let overallLengthIn = firstRowOfInfo |> bom.TryGetCellValueAtColumnAsString "BA"
        let tcxlFt = firstRowOfInfo |> bom.TryGetCellValueAtColumnAsString "BH"
        let tcxlIn = firstRowOfInfo |> bom.TryGetCellValueAtColumnAsString "BN"
        let tcxlType = firstRowOfInfo |> bom.TryGetCellValueAtColumnAsString "BU"
        let tcxrFt = firstRowOfInfo |> bom.TryGetCellValueAtColumnAsString "BX"
        let tcxrIn = firstRowOfInfo |> bom.TryGetCellValueAtColumnAsString "CD"
        let tcxrType = firstRowOfInfo |> bom.TryGetCellValueAtColumnAsString "CK"
        let bcxlFt = firstRowOfInfo |> bom.TryGetCellValueAtColumnAsString "DE"
        let bcxlIn = firstRowOfInfo |> bom.TryGetCellValueAtColumnAsString "DJ"
        let bcxlType = firstRowOfInfo |> bom.TryGetCellValueAtColumnAsString "DB"
        let bcxrFt = firstRowOfInfo |> bom.TryGetCellValueAtColumnAsString "DT"
        let bcxrIn = firstRowOfInfo |> bom.TryGetCellValueAtColumnAsString "DY"
        let bcxrType = firstRowOfInfo |> bom.TryGetCellValueAtColumnAsString "DQ"
        let bearingDepthLe = firstRowOfInfo |> bom.TryGetCellValueAtColumnAsString "EG"
        let bearingDepthRe = firstRowOfInfo |> bom.TryGetCellValueAtColumnAsString "EN"
        let baySlope = firstRowOfInfo |> bom.TryGetCellValueAtColumnAsString "EU"
        let baySlopeHighEnd = firstRowOfInfo |> bom.TryGetCellValueAtColumnAsString "FE"
        let leSeatSlope = firstRowOfInfo |> bom.TryGetCellValueAtColumnAsString "FH"
        let leSeatSlopeHighLow = firstRowOfInfo |> bom.TryGetCellValueAtColumnAsString "FR"
        let reSeatSlope = firstRowOfInfo |> bom.TryGetCellValueAtColumnAsString "FU"
        let reSeatSlopeHighLow = firstRowOfInfo |> bom.TryGetCellValueAtColumnAsString "GE"
        
        let secondRowOfInfo = joistSheet |> GetRow (25u + markLine*2u)
        let leBearingSlotLocFt = secondRowOfInfo |> bom.TryGetCellValueAtColumnAsString "AQ"
        let leBearingSlotLocIn = secondRowOfInfo |> bom.TryGetCellValueAtColumnAsString "AV"
        let leBearingSlotSize = secondRowOfInfo |> bom.TryGetCellValueAtColumnAsString "BE"
        let leBearingSlotLength = secondRowOfInfo |> bom.TryGetCellValueAtColumnAsString "BL"
        let leBearingSlotGage = secondRowOfInfo |> bom.TryGetCellValueAtColumnAsString "BR"
        let reBearingSlotLocFt = secondRowOfInfo |> bom.TryGetCellValueAtColumnAsString "CB"
        let reBearingSlotLocIn = secondRowOfInfo |> bom.TryGetCellValueAtColumnAsString "CG"
        let reBearingSlotSize = secondRowOfInfo |> bom.TryGetCellValueAtColumnAsString "CP"
        let reBearingSlotLength = secondRowOfInfo |> bom.TryGetCellValueAtColumnAsString "CW"
        let reBearingSlotGage = secondRowOfInfo |> bom.TryGetCellValueAtColumnAsString "DC"
        
        let thirdRowOfInfo = joistSheet |> GetRow (38u + markLine*2u)
        let netUplift = thirdRowOfInfo |> bom.TryGetCellValueAtColumnAsString "I"
        let adLoad = thirdRowOfInfo |> bom.TryGetCellValueAtColumnAsString "BN"
        
        let fourthRowOfInfo = joistSheet |> GetRow (51u + markLine*2u)
        let remarks = fourthRowOfInfo |> bom.TryGetCellValueAtColumnAsString "I"
        let generalRemarks =
            let generalRemarks =
                [for r in [ for i in 0u..4u -> 51u + i*2u] do
                  yield
                    bom.TryGetCellValueAtColumnAsString "EK" (joistSheet |> GetRow r)
                    |> Option.defaultValue ""]
                |> String.concat " "
            match generalRemarks with
            | Regex @"^ *$" _ -> None
            |  _ -> Some generalRemarks
            

        let tcConcentratedLoads =
            [for addRow in [0u..1u] do
                let row = joistSheet |> GetRow (134u + markLine*2u + addRow)
                for (loadCol, distFtCol, distInCol) in
                    [("I", "AB", "AH");
                    ("AS", "BL", "BR");
                    ("CC", "CV", "DB" );
                    ("DM", "EF", "EL");
                    ("EW", "FP", "FV") ] do

                    yield
                        {| Load = row |> bom.TryGetCellValueAtColumnAsString loadCol
                           DistanceFt = row |> bom.TryGetCellValueAtColumnAsString distFtCol
                           DistanceIn = row |> bom.TryGetCellValueAtColumnAsString distInCol |} ]

        let bcConcentratedLoads =
            [for addRow in [0u..1u] do
                let row = joistSheet |> GetRow (147u + markLine*2u + addRow)
                for (loadCol, distFtCol, distInCol) in
                    [("I", "AB", "AH");
                    ("AS", "BL", "BR");
                    ("CC", "CV", "DB" );
                    ("DM", "EF", "EL");
                    ("EW", "FP", "FV") ] do

                    yield
                        {| Load = row |> bom.TryGetCellValueAtColumnAsString loadCol
                           DistanceFt = row |> bom.TryGetCellValueAtColumnAsString distFtCol
                           DistanceIn = row |> bom.TryGetCellValueAtColumnAsString distInCol |} ]

        
        let fifthRowOfInfo = joistSheet |> GetRow (101u + markLine*2u)
        let bcUniformLoad = fifthRowOfInfo |> bom.TryGetCellValueAtColumnAsString "AG"
        let tcBendCheckLoad = fifthRowOfInfo |> bom.TryGetCellValueAtColumnAsString "BC"
        let bcBendCheckLoad = fifthRowOfInfo |> bom.TryGetCellValueAtColumnAsString "BM"
        let mainDrifts =
            [for (startLocFtCol, startLocInCol, startLoadCol, endLoadCol, lengthFtCol, lengthInCol) in
                [("CA", "CG", "CN", "CY", "DJ", "DP");
                ("EA", "EG", "EN", "EY", "FJ", "FP")] do

                yield
                    {|
                        StartLocFt = fifthRowOfInfo |> bom.TryGetCellValueAtColumnAsString startLocFtCol
                        StartLocIn = fifthRowOfInfo |> bom.TryGetCellValueAtColumnAsString startLocInCol
                        StartLoad = fifthRowOfInfo |> bom.TryGetCellValueAtColumnAsString startLoadCol
                        EndLoadCol = fifthRowOfInfo |> bom.TryGetCellValueAtColumnAsString endLoadCol
                        LengthFt = fifthRowOfInfo |> bom.TryGetCellValueAtColumnAsString lengthFtCol
                        LengthIn = fifthRowOfInfo |> bom.TryGetCellValueAtColumnAsString lengthInCol
                    |} ]

        
        {
            Mark = mark
            Quantity = quantity
            DesignationInfo = {| Depth = designationDepth ; DesignationType = designationType ; Loading = designationLoading |}
            OverallLength = {| Feet = overallLengthFt ; Inch = overallLengthIn |}
            Tcxl = {| Length = {| Feet = tcxlFt ; Inch = tcxlIn |} ; TcxType = tcxlType|}
            Tcxr = {| Length = {| Feet = tcxrFt ; Inch = tcxrIn |} ; TcxType = tcxrType|}
            Bcxl = {| Length = {| Feet = bcxlFt ; Inch = bcxlIn |} ; BcxType = bcxlType|}
            Bcxr = {| Length = {| Feet = bcxrFt ; Inch = bcxrIn |} ; BcxType = bcxrType|}
            BearingDepthLe = bearingDepthLe
            BearingDepthRe = bearingDepthRe
            BaySlope = baySlope
            BaySlopeHighEnd = baySlopeHighEnd
            LeSeatSlope = leSeatSlope
            LeSeatSlopeHighLow = leSeatSlopeHighLow
            ReSeatSlope = reSeatSlope
            ReSeatSlopeHighLow = reSeatSlopeHighLow
            LeBearingSlotInfo = 
              {|
                  Location = {| Feet = leBearingSlotLocFt ; Inch = leBearingSlotLocIn |}
                  Size = leBearingSlotSize
                  Length = leBearingSlotLength
                  Gage = leBearingSlotGage
              |}
            ReBearingSlotInfo =
              {|
                  Location = {|Feet = reBearingSlotLocFt ; Inch = reBearingSlotLocIn |}
                  Size = reBearingSlotSize
                  Length = reBearingSlotLength
                  Gage = reBearingSlotGage
              |}
            NetUplift = netUplift
            AdLoad = adLoad
            Remarks = remarks
            GeneralRemarks = generalRemarks
            TcConcentratedLoads = tcConcentratedLoads
            BcConcentratedLoads = bcConcentratedLoads
            BcUniformLoad = bcUniformLoad
            TcBendCheckLoad = tcBendCheckLoad
            BcBendCheckLoad = bcBendCheckLoad
            MainDrifts = mainDrifts
        } ]


    

let GetJoists(bom : Bom) =
    let stringTable = bom.StringTable
    let joistSheets = bom.JoistSheets
    joistSheets
        |> Seq.map
            (fun jSheet -> GetJoistsFromOneSheet bom jSheet)
        |> Seq.concat


let GetJoistInfo (fileName : string) =
    let bom = GetBom fileName
    GetJoists bom