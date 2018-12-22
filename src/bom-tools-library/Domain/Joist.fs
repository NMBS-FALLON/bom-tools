namespace DESign.BomTools.Domain

open DESign.BomTools.Dto
open System
open FSharp.Core

[<AutoOpen>]
module internal Helpers =
    let (|KCS|K|LH|G|) (joistSize : string) =
        match joistSize with
        | size when size.ToUpper().Contains("KCS") -> KCS
        | size when size.ToUpper().Contains("K") -> K
        | size when size.ToUpper().Contains("LH") -> LH
        | size when size.ToUpper().Contains("G") -> G
        | _ -> failwith (sprintf "Unknown joist type: %s" joistSize)

type Joist =
    { Mark : string
      Quantity : int
      JoistSize : string
      OverallLength : float
      TcxlLength : float
      TcxlType : string option
      TcxrLength : float
      TcxrType : string option
      SeatDepthLeft : float
      SeatDepthRight : float
      BcxlLength : float option
      BcxlType : string option
      BcxrLength : float option
      BcxrType : string option
      PunchedSeatsLeft : float option
      PunchedSeatsRight : float option
      PunchedSeatsGa : float option
      OverallSlope : float
      SpecialNotes : string seq
      LoadNotes : string seq }
    static member Parse(joistDto : JoistDto) =
        let mark =
            match joistDto.Mark with
            | Some mark -> mark
            | None -> failwith "There is a mark without a label"

        let quantity =
            match joistDto.Quantity with
            | Some qty -> qty
            | None -> failwith (sprintf "Mark %s does not have a quantity" mark)

        let joistSize =
            match joistDto.JoistSize with
            | Some joistSize -> joistSize
            | None ->
                failwith (sprintf "Mark %s does not have a joist size" mark)

        let overallLength =
            match joistDto.OverallLengthFt, joistDto.OverallLengthIn with
            | Some ft, Some inch -> ft + inch / 12.0
            | Some ft, None -> ft
            | None, Some inch -> inch
            | None, None ->
                failwith (sprintf "Mark %s does not have an overal length" mark)

        let tcxlLength =
            match joistDto.TcxlLengthFt, joistDto.TcxlLengthIn with
            | Some ft, Some inch -> ft + inch / 12.0
            | Some ft, None -> ft
            | None, Some inch -> inch
            | None, None -> 0.0

        let tcxlType = joistDto.TcxlType

        let tcxrLength =
            match joistDto.TcxrLengthFt, joistDto.TcxrLengthIn with
            | Some ft, Some inch -> ft + inch / 12.0
            | Some ft, None -> ft
            | None, Some inch -> inch
            | None, None -> 0.0

        let tcxrType = joistDto.TcxrType

        let seatDepthLeft =
            match joistDto.SeatDepthLeft with
            | Some sd -> sd
            | None ->
                match joistSize with
                | KCS | K -> 2.5
                | LH -> 5.0
                | G -> 7.5

        let seatDepthRight =
            match joistDto.SeatDepthRight with
            | Some sd -> sd
            | none ->
                match joistSize with
                | KCS | K -> 2.5
                | LH -> 5.0
                | G -> 7.5

        let bcxlLength =
            match joistDto.BcxlLengthFt, joistDto.BcxlLengthIn with
            | Some ft, Some inch -> Some(ft + inch / 12.0)
            | Some ft, None -> Some ft
            | None, Some inch -> Some inch
            | None, None -> None

        let bcxlType = joistDto.BcxlType

        let bcxrLength =
            match joistDto.BcxrLengthFt, joistDto.BcxrLengthIn with
            | Some ft, Some inch -> Some(ft + inch / 12.0)
            | Some ft, None -> Some ft
            | None, Some inch -> Some inch
            | None, None -> None

        let bcxrType = joistDto.BcxrType

        let punchedSeatsLeft =
            match joistDto.PunchedSeatsLeftFt, joistDto.PunchedSeatsLeftIn with
            | Some ft, Some inch -> Some(ft + inch / 12.0)
            | Some ft, None -> Some ft
            | None, Some inch -> Some inch
            | None, None -> None

        let punchedSeatsRight =
            match joistDto.PunchedSeatsRightFt, joistDto.PunchedSeatsRightIn with
            | Some ft, Some inch -> Some(ft + inch / 12.0)
            | Some ft, None -> Some ft
            | None, Some inch -> Some inch
            | None, None -> None

        let punchedSeatsGa = joistDto.PunchedSeatsGa

        let overallSlope =
            match joistDto.OverallSlope with
            | Some slope -> slope
            | None -> 0.0

        let specialNotes = // this can definitly be improved with regex
            match joistDto.Notes with
            | Some notes ->
                if notes.Contains("[") then
                    notes.Split([| "["; "]" |], StringSplitOptions.RemoveEmptyEntries).[0]
                        .Split([| "," |], StringSplitOptions.RemoveEmptyEntries)
                    |> Seq.map (fun s -> s.Replace(" ", ""))
                else Seq.empty
            | None -> Seq.empty

        let loadNotes = // this can definitly be improved with regex
            match joistDto.Notes with
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

        { Mark = mark
          Quantity = quantity
          JoistSize = joistSize
          OverallLength = overallLength
          TcxlLength = tcxlLength
          TcxlType = tcxlType
          TcxrLength = tcxrLength
          TcxrType = tcxrType
          SeatDepthLeft = seatDepthLeft
          SeatDepthRight = seatDepthRight
          BcxlLength = bcxlLength
          BcxlType = bcxlType
          BcxrLength = bcxrLength
          BcxrType = bcxrType
          PunchedSeatsLeft = punchedSeatsLeft
          PunchedSeatsRight = punchedSeatsRight
          PunchedSeatsGa = punchedSeatsGa
          OverallSlope = overallSlope
          SpecialNotes = specialNotes
          LoadNotes = loadNotes }
