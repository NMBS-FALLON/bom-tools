namespace DESign.BomTools.Domain

open DESign.BomTools.Dto
open System
open FSharp.Core
open DESign.Helpers


type JoistSeries =
    | JS
    | KCS
    | K
    | LH
    | DLH

    override this.ToString () =
        match this with
        | JS -> "JS"
        | KCS -> "KCS"
        | K -> "K"
        | LH -> "LH"
        | DLH -> "DLH"

    static member Parse s =
        match s with
        | "JS" -> Ok JS
        | "KCS" -> Ok KCS
        | "K" -> Ok K
        | "LH" -> Ok LH
        | "DLH" -> Ok DLH
        | _ -> sprintf "Unknown joist series, %s" s |> Error

type JoistDesignationInfo =
    {
        Depth : float
        Series : JoistSeries
        SeriesNumber : string option
        TotalAndLive : (float * float) option
    }



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

    member this.DesignationInfo = 
        match this.JoistSize with
        | Regex @"(\d+\.?\d*)(KCS|K|LH|DLH)(\d+\.?\d*)\/(\d+\.?\d*)" [depth; series; tl; ll] ->
            let depth = FSharp.Core.float.Parse depth
            let series = JoistSeries.Parse series
            let tl = FSharp.Core.float.Parse tl
            let ll = FSharp.Core.float.Parse ll
            match series with
            | Ok series -> Ok {Depth = depth; Series = series; SeriesNumber = None; TotalAndLive = Some (tl, ll)}
            | Error e -> sprintf "Mark %s: %s." this.Mark e |> Error
        | Regex @"(\d+\.?\d*)(KCS|K|LH|DLH)(\d{1,2}\b)" [depth; series; seriesNumber] ->
            let depth = FSharp.Core.float.Parse depth
            let series = JoistSeries.Parse series
            match series with
            | Ok series -> Ok {Depth = depth; Series = series; SeriesNumber = Some seriesNumber; TotalAndLive = None}
            | Error e -> sprintf "Mark %s: %s." this.Mark e |> Error
        | _ -> sprintf "Mark %s: Designation is not in the correct format" this.Mark |> Error


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
                | DESign.Helpers.KCS | DESign.Helpers.K -> 2.5
                | DESign.Helpers.LH -> 5.0
                | DESign.Helpers.G -> 7.5

        let seatDepthRight =
            match joistDto.SeatDepthRight with
            | Some sd -> sd
            | none ->
                match joistSize with
                | DESign.Helpers.KCS | DESign.Helpers.K -> 2.5
                | DESign.Helpers.LH -> 5.0
                | DESign.Helpers.G -> 7.5

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

    member this.IsTlOverLl =
        match this.JoistSize with
        | Regex @"\d+\.?\d*[K|LH|DLH](\d+)/(\d+)" [tl; ll] -> true
        | _ -> false 

    member this.UniformDead =
        match this.JoistSize with
        | Regex @"\d+\.?\d*[K|LH|DLH](\d+)/(\d+)" [tl; ll] ->
            let tl = FSharp.Core.double.Parse tl
            let ll = FSharp.Core.double.Parse ll
            Some (tl - ll)
        | _ -> None

    member this.UniformLive =
        match this.JoistSize with
        | Regex @"\d+\.?\d*[K|LH|DLH](\d+)/(\d+)" [_; ll] ->
            let ll = FSharp.Core.double.Parse ll
            Some (ll)
        | _ -> None



