namespace DESign.BomTools.Domain

open DESign.BomTools.Dto
open System
open FSharp.Core

type Note =
    { ID : string
      Notes : string seq }
    static member Parse(noteDto : NoteDto) =
        let id =
            match noteDto.ID with
            | Some s -> s
            | None -> failwith "There is a note that does not have a 'NAME'." // I dont think this is even possible...

        let notes =
            noteDto.Notes
            |> Seq.filter (fun noteDto -> noteDto.IsSome && noteDto.Value <> "")
            |> Seq.map (fun noteDto -> noteDto.Value)

        { ID = id
          Notes = notes }
