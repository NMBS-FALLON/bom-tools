namespace DESign.BomTools.Domain

open DESign.BomTools.Dto
open System
open FSharp.Core

type Load =
    { ID : string
      Type : string
      Category : string
      Position : string
      Load1Value : float
      Load1DistanceFt : float option
      Load1DistanceIn : float option
      Load2Value : float option
      Load2DistanceFt : float option
      Load2DistanceIn : float option
      Ref : string option
      LoadCases : int seq
      Remarks : string option }
    static member Parse(loadDto : LoadDto) =
        let id =
            match loadDto.ID with
            | Some id -> id
            | None -> failwith "There is a load that does not have a 'LOAD #'."

        let _type =
            match loadDto.Type with
            | Some _type -> _type
            | None ->
                failwith (sprintf "Load # %s has a load without a 'TYPE'." id)

        let category =
            match loadDto.Category with
            | Some cat -> cat
            | None ->
                failwith (sprintf "Load # %s has a load without a 'CAT'." id)

        let position =
            match loadDto.Position with
            | Some pos -> pos
            | None ->
                failwith (sprintf "Load # %s has a load without a 'POS'." id)

        let load1Value =
            match loadDto.Load1Value with
            | Some load -> load
            | None ->
                failwith
                    (sprintf "Load # %s has a load without a 'LOAD 1 VALUE'." id)

        let loadCases =
            match loadDto.LoadCases with
            | None -> Seq.empty
            | Some lcs ->
                let test =
                    lcs.Replace(" ", "")
                       .Split([| "," |], StringSplitOptions.RemoveEmptyEntries)
                    |> Seq.ofArray
                test |> Seq.map FSharp.Core.int.Parse

        { ID = id
          Type = _type
          Category = category
          Position = position
          Load1Value = load1Value
          Load1DistanceFt = loadDto.Load1DistanceFt
          Load1DistanceIn = loadDto.Load1DistanceFt
          Load2Value = loadDto.Load2Value
          Load2DistanceFt = loadDto.Load2DistanceFt
          Load2DistanceIn = loadDto.Load2DistanceIn
          Ref = loadDto.Ref
          LoadCases = loadCases
          Remarks = loadDto.Remarks }
