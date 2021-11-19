namespace DESign.BomTools.Domain

open DESign.BomTools.Dto
open DESign.BomTools.Domain
open System
open FSharp.Core
open System.Text.RegularExpressions
open DocumentFormat.OpenXml.ExtendedProperties
open DESign.Helpers



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
        match note.Note with
        | Regex pattern groups -> Some groups
        | _ -> None 

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

  let inline tryGetLiveLoadNote (joist: ^JoistOrGirder ) (job : Job) =
    let specialNotes = (^JoistOrGirder: (member SpecialNotes : seq<string>) (joist))
    let possilbeParticularNote =
      job.ParticularNotes
      |> Seq.filter (fun n -> specialNotes |> Seq.contains (n.ID))
      |> Seq.map
        (fun n ->
          match n.Note with
          | Regex @"[LS] *= *(\d+\.?\d*) *(:?K|%)" [llNote; llType] ->
            let ll = FSharp.Core.float.Parse llNote
            match llType with
            | "K" | "k" -> Kip ll |> Some
            | "%" -> Percent ll |> Some
            | _ -> None
          | _ -> None )
      |> Seq.choose id
      |> Seq.tryItem 0
    let possibleTypicalNote =
      job |> tryGetTypicalLiveLoad
    match (possilbeParticularNote, possibleTypicalNote) with
    | (Some particularNote, _ ) -> Some particularNote
    | ( None, Some typicalNote) -> Some typicalNote
    | _ -> None
    
    


  let getJoistLoads (joist: Joist) (job: Job) =
    job.Loads
    |> Seq.filter (fun load -> joist.LoadNotes |> Seq.contains load.ID)

  let getGirderLoads (girder: Girder) (job: Job) =
    job.Loads
    |> Seq.filter (fun load -> girder.LoadNotes |> Seq.contains load.ID)

  let getJoistParticularNotes (joist: Joist) (job: Job) =
    job.ParticularNotes
    |> Seq.filter (fun note -> joist.SpecialNotes |> Seq.contains note.ID)

  let getGirderParticularNotes (girder: Girder) (job: Job) =
    job.ParticularNotes
    |> Seq.filter (fun note -> girder.SpecialNotes |> Seq.contains note.ID)




    
