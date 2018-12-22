namespace DESign.BomTools.Domain

open DESign.BomTools.Dto
open System
open FSharp.Core

type AdditionalLoad =
    { Location : float
      Load : float }

type Girder =
    { Mark : string
      Quantity : int
      GirderSize : string
      TcWidth : float option
      OverallLength : float
      TcxlLength : float
      TcxlType : string option
      TcxrLength : float
      TcxrType : string option
      SeatDepthLeft : float
      SeatDepthRight : float
      BcxlLength : float option
      BcxrLength : float option
      PunchedSeatsLeft : float option
      PunchedSeatsRight : float option
      PunchedSeatsGa : float option
      OverallSlope : float
      SpecialNotes : string seq
      LoadNotes : string seq
      NumKbRequired : int option
      PanelLocations : float seq
      AdditionalJoists : AdditionalLoad seq }
    static member Parse(girderDto : GirderDto) =
        let mark =
            match girderDto.Mark with
            | Some mark -> mark
            | None -> failwith "There is a mark without a label"

        let quantity =
            match girderDto.Quantity with
            | Some qty -> qty
            | None -> failwith (sprintf "Mark %s does not have a quantity" mark)

        let girderSize =
            match girderDto.GirderSize with
            | Some girderSize -> girderSize
            | None ->
                failwith (sprintf "Mark %s does not have a girder size" mark)

        let overallLength =
            match girderDto.OverallLengthFt, girderDto.OverallLengthIn with
            | Some ft, Some inch -> ft + inch / 12.0
            | Some ft, None -> ft
            | None, Some inch -> inch
            | None, None ->
                failwith (sprintf "Mark %s does not have an overal length" mark)

        let tcxlLength =
            match girderDto.TcxlLengthFt, girderDto.TcxlLengthIn with
            | Some ft, Some inch -> ft + inch / 12.0
            | Some ft, None -> ft
            | None, Some inch -> inch
            | None, None -> 0.0

        let tcxlType = girderDto.TcxlType

        let tcxrLength =
            match girderDto.TcxrLengthFt, girderDto.TcxrLengthIn with
            | Some ft, Some inch -> ft + inch / 12.0
            | Some ft, None -> ft
            | None, Some inch -> inch
            | None, None -> 0.0

        let tcxrType = girderDto.TcxrType

        let seatDepthLeft =
            match girderDto.SeatDepthLeft with
            | Some sd -> sd
            | None ->
                match girderSize with
                | KCS | K -> 2.5
                | LH -> 5.0
                | G -> 7.5

        let seatDepthRight =
            match girderDto.SeatDepthRight with
            | Some sd -> sd
            | none ->
                match girderSize with
                | KCS | K -> 2.5
                | LH -> 5.0
                | G -> 7.5

        let bcxlLength =
            match girderDto.BcxlLengthFt, girderDto.BcxlLengthIn with
            | Some ft, Some inch -> Some(ft + inch / 12.0)
            | Some ft, None -> Some ft
            | None, Some inch -> Some inch
            | None, None -> None

        let bcxrLength =
            match girderDto.BcxrLengthFt, girderDto.BcxrLengthIn with
            | Some ft, Some inch -> Some(ft + inch / 12.0)
            | Some ft, None -> Some ft
            | None, Some inch -> Some inch
            | None, None -> None

        let punchedSeatsLeft =
            match girderDto.PunchedSeatsLeftFt, girderDto.PunchedSeatsLeftIn with
            | Some ft, Some inch -> Some(ft + inch / 12.0)
            | Some ft, None -> Some ft
            | None, Some inch -> Some inch
            | None, None -> None

        let punchedSeatsRight =
            match girderDto.PunchedSeatsRightFt, girderDto.PunchedSeatsRightIn with
            | Some ft, Some inch -> Some(ft + inch / 12.0)
            | Some ft, None -> Some ft
            | None, Some inch -> Some inch
            | None, None -> None

        let punchedSeatsGa = girderDto.PunchedSeatsGa

        let overallSlope =
            match girderDto.OverallSlope with
            | Some slope -> slope
            | None -> 0.0

        let specialNotes = // this can definitly be improved with regex
            match girderDto.Notes with
            | Some notes ->
                if notes.Contains("[") then
                    notes.Split([| "["; "]" |], StringSplitOptions.RemoveEmptyEntries).[0]
                        .Split([| "," |], StringSplitOptions.RemoveEmptyEntries)
                    |> Seq.map (fun s -> s.Replace(" ", ""))
                else Seq.empty
            | None -> Seq.empty

        let loadNotes = // this can definitly be improved with regex
            match girderDto.Notes with
            | Some notes ->
                if notes.Contains("[") && notes.Contains("(") then
                    notes.Split([| "("; ")" |], StringSplitOptions.RemoveEmptyEntries).[1]
                        .Split([| "," |], StringSplitOptions.RemoveEmptyEntries)
                    |> Seq.map (fun s -> s.Replace(" ", ""))
                else if notes.Contains("(") then
                    notes.Split([| "("; ")" |], StringSplitOptions.RemoveEmptyEntries).[0]
                        .Split([| "," |], StringSplitOptions.RemoveEmptyEntries)
                    |> Seq.map (fun s -> s.Replace(" ", ""))
                else Seq.empty
            | None -> Seq.empty

        let tcWidth = girderDto.TcWidth
        let numKbRequired = girderDto.NumKbRequired

        let panelLocations =
            match girderDto.PanelLocations with
            | Some panels -> panels
            | None -> failwith (sprintf "No Panel Locations on girder %s" mark)

        let (|PanelPointLoad|LocatedLoad|) (additionalJoist : AdditionalJoistOnGirderDto) =
            match additionalJoist.LocationFt with
            | Some loc when loc.ToUpper().Contains("P") -> PanelPointLoad
            | _ -> LocatedLoad

        let additionalJoists =
            girderDto.AdditionalJoistLoads
            |> Seq.filter (fun a -> a.LoadValue.IsSome)
            |> Seq.map (fun a ->
                   match a with
                   | PanelPointLoad ->
                       let panelNum =
                           match a.LocationIn with
                           | Some panelNum ->
                               FSharp.Core.int.Parse(panelNum.Replace("#", ""))
                           | None ->
                               failwith
                                   (sprintf
                                        "Mark %s has an 'additional load' that is missing the panel number"
                                        mark)

                       let panelLocation =
                           match panelLocations |> Seq.tryItem panelNum with
                           | Some loc -> loc
                           | None ->
                               failwith
                                   "Mark %s has an 'additional load' that is located at a non-existant panel point"

                       let loadValue =
                           match a.LoadValue with
                           | Some v -> v
                           | None ->
                               failwith
                                   (sprintf
                                        "Mark %s has an 'additional load' that does not have a load value"
                                        mark)

                       { Location = panelLocation
                         Load = loadValue }
                   | LocatedLoad ->
                       let locationFt =
                           match a.LocationFt with
                           | Some ft when ft = "" -> 0.0
                           | Some ft ->
                               match FSharp.Core.float.TryParse ft with
                               | true, v -> v
                               | false, _ ->
                                   failwith
                                       (sprintf
                                            "Mark %s has an 'additional load' that does not have a proper Location Ft"
                                            mark)
                           | None -> 0.0

                       let locationIn =
                           match a.LocationIn with
                           | Some inch when inch = "" -> 0.0
                           | Some inch ->
                               match FSharp.Core.float.TryParse inch with
                               | true, v -> v / 12.0
                               | false, _ ->
                                   failwith
                                       (sprintf
                                            "Mark %s has an 'additional load' that does not have a proper Location In"
                                            mark)
                           | None -> 0.0

                       let loadValue =
                           match a.LoadValue with
                           | Some v -> v
                           | None ->
                               failwith
                                   (sprintf
                                        "Mark %s has an 'additional load' that does not have a load value"
                                        mark)

                       { Location = locationFt + locationIn
                         Load = loadValue })

        let girder =
            { Mark = mark
              Quantity = quantity
              GirderSize = girderSize
              OverallLength = overallLength
              TcWidth = tcWidth
              TcxlLength = tcxlLength
              TcxlType = tcxlType
              TcxrLength = tcxrLength
              TcxrType = tcxrType
              SeatDepthLeft = seatDepthLeft
              SeatDepthRight = seatDepthRight
              BcxlLength = bcxlLength
              BcxrLength = bcxrLength
              PunchedSeatsLeft = punchedSeatsLeft
              PunchedSeatsRight = punchedSeatsRight
              PunchedSeatsGa = punchedSeatsGa
              OverallSlope = overallSlope
              SpecialNotes = specialNotes
              LoadNotes = loadNotes
              NumKbRequired = numKbRequired
              PanelLocations = panelLocations
              AdditionalJoists = additionalJoists }

        girder
