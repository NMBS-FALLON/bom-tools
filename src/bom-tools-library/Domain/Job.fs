namespace DESign.BomTools.Domain

open DESign.BomTools.Dto
open DESign.BomTools.Domain
open System
open FSharp.Core
open System.Text.RegularExpressions
open DocumentFormat.OpenXml.ExtendedProperties



type Job =
  { Name : string
    GeneralNotes : Note seq
    ParticularNotes : Note seq
    Loads : Load seq
    Girders : Girder seq
    Joists : Joist seq}

module Job =

  let (|Regex|_|) pattern input =
        let m = Regex.Match(input, pattern)
        if m.Success then Some(List.tail [ for g in m.Groups -> g.Value ])
        else None

  
  let tryGetGroups (pattern : string) (note : Note) =
    let fullNote =
      match note.Notes |> Seq.isEmpty with
      | true -> None
      | false -> Some (note.Notes |> Seq.reduce (fun first second -> first + " " + second))
    fullNote
    |> Option.bind
      (fun note -> 
        match note with
        | Regex pattern groups -> Some groups
        | _ -> None )

  let tryGetNotes pattern (notes : Note seq) =
    let possibleNotes =
      notes |> Seq.choose (fun note -> note |> tryGetGroups pattern)
    if possibleNotes |> Seq.isEmpty then
      None
    else
      Some possibleNotes

  let tryGetSds (job : Job) =
    let sdsPat = @"SDS *= *(\d+\.?\d*)"
    let possibleSdsNotes = job.GeneralNotes |> tryGetNotes sdsPat
  
    possibleSdsNotes
    |> Option.bind
      (fun notes -> 
        let sdsString = notes |> Seq.head |> List.item 0 
        Some (FSharp.Core.float.Parse sdsString) )

  type LiveLoadNote =
    | Percent of float
    | Kip of float

  let tryGetTypicalLiveLoad (job : Job) =
    let liveLoadPat = @"[LS] *= *(\d+\.?\d*) *([Kk%])"
    let possibleLiveLoadNote =
      job.GeneralNotes
      |> Seq.choose
        (fun note -> note |> tryGetGroups liveLoadPat)
    if possibleLiveLoadNote |> Seq.isEmpty then
      None
    else
      let liveLoadValueString = possibleLiveLoadNote |> Seq.head |> List.item 0
      let liveLoadValue = FSharp.Core.float.Parse liveLoadValueString
      let liveLoadRepresentationString = possibleLiveLoadNote |> Seq.head |> List.item 1
      let liveLoadNote =
        match liveLoadRepresentationString with
        | "K" | "k" -> Kip liveLoadValue
        | "%" -> Percent liveLoadValue
        | _ -> failwith "This shouldnt happen"
      Some liveLoadNote


    
