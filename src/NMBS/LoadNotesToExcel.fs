module DESign.BomTools.LoadNotesToExcel

open DESign.BomTools.Domain
open OfficeOpenXml
open System.IO
open DocumentFormat.OpenXml.Spreadsheet



let getLoadNotesByMark (job : Job) =
   let joistLoadNotes =
       job.Joists
       |> Seq.collect
           (fun joist ->
               job
               |> Job.getJoistLoads joist
               |> Seq.map
                   (fun load -> joist.Mark, load))

   let girderLoadNotes =
       job.Girders
       |> Seq.collect
           (fun girder ->
               job |> Job.getGirderLoads girder
               |> Seq.map
                   (fun load -> girder.Mark, load))

   let additionalGirderLoads =
        job.Girders
        |> Seq.collect
            (fun girder ->
                let additionalLoadOnGirder =
                    DESign.Helpers.handleResultWithFailure girder.AdditionalJoists
                    |> Seq.map
                        (fun additionalLoad ->
                            let locationFt = System.Math.Truncate additionalLoad.Location
                            let locationIn = (additionalLoad.Location - locationFt) * 12.0
                            Load.Create("", "C", "TL", "TC", additionalLoad.Load * 1000.0, Some locationFt, Some locationIn, None, None, None, Some "L-BL", [], Some "JOIST TRANSFER LOAD TO GIRDER"))
                additionalLoadOnGirder
                |> Seq.map (fun addLoad -> girder.Mark, addLoad))

   let allLoads = [girderLoadNotes; additionalGirderLoads; joistLoadNotes] |> Seq.concat
   allLoads

let getLoadNotesByNote (job : Job) =
   
   let joistLoadsByNote =
       job.Loads
       |> Seq.map
           (fun load ->
               let joistsWithThisLoad =
                   job.Joists
                   |> Seq.filter (fun joist -> joist.LoadNotes |> Seq.contains load.ID)
                   |> Seq.map (fun joist -> joist.Mark)
               load, joistsWithThisLoad)


   let girderLoadsByNote =
        job.Loads
        |> Seq.map
            (fun load ->
                let girdersWithThisLoad =
                    job.Girders
                    |> Seq.filter (fun girder -> girder.LoadNotes |> Seq.contains load.ID)
                    |> Seq.map (fun girder -> girder.Mark)
                load, girdersWithThisLoad)



   let allLoads = [joistLoadsByNote; girderLoadsByNote] |> Seq.concat
   allLoads


let CreateBomInfoSheetFromJob (job : Job) =
   let package = new ExcelPackage()
   let loadNotesByMarkSheet = package.Workbook.Worksheets.Add("Load Notes By Mark")
   loadNotesByMarkSheet.Cells.[1,1].Value <- "Mark"
   loadNotesByMarkSheet.Cells.[1,2].Value <- "ID"
   loadNotesByMarkSheet.Cells.[1,3].Value <- "Type"
   loadNotesByMarkSheet.Cells.[1,4].Value <- "Category"
   loadNotesByMarkSheet.Cells.[1,5].Value <- "Position"
   loadNotesByMarkSheet.Cells.[1,6].Value <- "Value 1"
   loadNotesByMarkSheet.Cells.[1,7].Value <- "Val 1 Distance Ft"
   loadNotesByMarkSheet.Cells.[1,8].Value <- "Val 1 Distance In"
   loadNotesByMarkSheet.Cells.[1,9].Value <- "Value 2"
   loadNotesByMarkSheet.Cells.[1,10].Value <- "Val 2 Distance Ft"
   loadNotesByMarkSheet.Cells.[1,11].Value <- "Val 2 Distance In"
   loadNotesByMarkSheet.Cells.[1,12].Value <- "Reference"
   loadNotesByMarkSheet.Cells.[1,13].Value <- "Load Case(s)"
   loadNotesByMarkSheet.Cells.[1,14].Value <- "Remarks"
   let notesByMark = job |> getLoadNotesByMark
   notesByMark
   |> Seq.iteri
       (fun i (mark, load) ->
           loadNotesByMarkSheet.Cells.[i + 2, 1].Value <- mark
           loadNotesByMarkSheet.Cells.[i + 2, 2].Value <- load.ID
           loadNotesByMarkSheet.Cells.[i + 2, 3].Value <- load.Type
           loadNotesByMarkSheet.Cells.[i + 2, 4].Value <- load.Category
           loadNotesByMarkSheet.Cells.[i + 2, 5].Value <- load.Position
           loadNotesByMarkSheet.Cells.[i + 2, 6].Value <- load.Load1Value 
           loadNotesByMarkSheet.Cells.[i + 2, 7].Value <- match load.Load1DistanceFt with Some d -> box d | None -> null
           loadNotesByMarkSheet.Cells.[i + 2, 8].Value <- match load.Load1DistanceIn with Some d -> box d | None -> null
           loadNotesByMarkSheet.Cells.[i + 2, 9].Value <- match load.Load2Value with Some v -> box v | None -> null
           loadNotesByMarkSheet.Cells.[i + 2, 10].Value <- match load.Load2DistanceFt with Some d -> box d | None -> null
           loadNotesByMarkSheet.Cells.[i + 2, 11].Value <- match load.Load2DistanceIn with Some d -> box d | None -> null
           loadNotesByMarkSheet.Cells.[i + 2, 12].Value <- load.Ref |> Option.defaultValue ""
           loadNotesByMarkSheet.Cells.[i + 2, 13].Value <- load.LoadCases |> Seq.map (fun v -> v.ToString()) |> DESign.Helpers.ConvertToCommaSeparatedString
           loadNotesByMarkSheet.Cells.[i + 2, 14].Value <- load.Remarks |> Option.defaultValue "" )


   let lastRowForLoadNotesByMarkSheet = loadNotesByMarkSheet.Dimension.End.Row
   let tableRangeForLoadNotesByMarkSheet = loadNotesByMarkSheet.Cells.[1, 1, lastRowForLoadNotesByMarkSheet, 14]
   let loadNotesByMarkTableName = "tblLoadNotesByMark"
   let loadNotesByMarkTable =loadNotesByMarkSheet.Tables.Add(tableRangeForLoadNotesByMarkSheet, loadNotesByMarkTableName)
   loadNotesByMarkTable.TableStyle <- Table.TableStyles.Light8
   tableRangeForLoadNotesByMarkSheet.AutoFitColumns()

   (*
   // Load Notes By Note
   let loadNotesByNoteSheet = package.Workbook.Worksheets.Add("Load Notes By Note")
   loadNotesByNoteSheet.Cells.[1,1].Value <- "ID"
   loadNotesByNoteSheet.Cells.[1,2].Value <- "Type"
   loadNotesByNoteSheet.Cells.[1,3].Value <- "Category"
   loadNotesByNoteSheet.Cells.[1,4].Value <- "Position"
   loadNotesByNoteSheet.Cells.[1,5].Value <- "Value 1"
   loadNotesByNoteSheet.Cells.[1,6].Value <- "Val 1 Distance Ft"
   loadNotesByNoteSheet.Cells.[1,7].Value <- "Val 1 Distance In"
   loadNotesByNoteSheet.Cells.[1,8].Value <- "Value 2"
   loadNotesByNoteSheet.Cells.[1,9].Value <- "Val 2 Distance Ft"
   loadNotesByNoteSheet.Cells.[1,10].Value <- "Val 2 Distance In"
   loadNotesByNoteSheet.Cells.[1,11].Value <- "Reference"
   loadNotesByNoteSheet.Cells.[1,12].Value <- "Load Case(s)"
   loadNotesByNoteSheet.Cells.[1,13].Value <- "Remarks"
   loadNotesByNoteSheet.Cells.[1,14].Value <- "Marks With This Load"

   let notesByNote = job |> getLoadNotesByNote
   notesByNote
   |> Seq.iteri
       (fun i (load, marks) ->
           loadNotesByNoteSheet.Cells.[i + 2, 1].Value <- load.ID
           loadNotesByNoteSheet.Cells.[i + 2, 2].Value <- load.Type
           loadNotesByNoteSheet.Cells.[i + 2, 3].Value <- load.Category
           loadNotesByNoteSheet.Cells.[i + 2, 4].Value <- load.Position
           loadNotesByNoteSheet.Cells.[i + 2, 5].Value <- load.Load1Value 
           loadNotesByNoteSheet.Cells.[i + 2, 6].Value <- match load.Load1DistanceFt with Some d -> box d | None -> null
           loadNotesByNoteSheet.Cells.[i + 2, 7].Value <- match load.Load1DistanceIn with Some d -> box d | None -> null
           loadNotesByNoteSheet.Cells.[i + 2, 8].Value <- match load.Load2Value with Some v -> box v | None -> null
           loadNotesByNoteSheet.Cells.[i + 2, 9].Value <- match load.Load2DistanceFt with Some d -> box d | None -> null
           loadNotesByNoteSheet.Cells.[i + 2, 10].Value <- match load.Load2DistanceIn with Some d -> box d | None -> null
           loadNotesByNoteSheet.Cells.[i + 2, 11].Value <- load.Ref |> Option.defaultValue ""
           loadNotesByNoteSheet.Cells.[i + 2, 12].Value <- load.LoadCases |> Seq.map (fun v -> v.ToString()) |> DESign.Helpers.ConvertToCommaSeparatedString
           loadNotesByNoteSheet.Cells.[i + 2, 13].Value <- load.Remarks |> Option.defaultValue ""
           loadNotesByNoteSheet.Cells.[i + 2, 14].Value <- marks |> DESign.Helpers.ConvertToCommaSeparatedString )


   let lastRowForLoadNotesByNoteSheet = loadNotesByNoteSheet.Dimension.End.Row
   let tableRangeForLoadNotesByNoteSheet = loadNotesByNoteSheet.Cells.[1, 1, lastRowForLoadNotesByNoteSheet, 14]
   let LoadNotesByNoteTableName = "tblLoadNotesByNote"
   let loadNotesByNoteTable =loadNotesByNoteSheet.Tables.Add(tableRangeForLoadNotesByNoteSheet, LoadNotesByNoteTableName)
   loadNotesByNoteTable.TableStyle <- Table.TableStyles.Light8
   tableRangeForLoadNotesByNoteSheet.AutoFitColumns()
   *)
   package
   

let CreateBomInfoSheetFromJob2 (job : Job) =
   let package = new ExcelPackage()
   let loadNotesByMarkSheet = package.Workbook.Worksheets.Add("Load Notes By Mark")
   loadNotesByMarkSheet.Cells.[1,1].Value <- "Mark"
   loadNotesByMarkSheet.Cells.[1,2].Value <- "ID"
   loadNotesByMarkSheet.Cells.[1,3].Value <- "Type"
   loadNotesByMarkSheet.Cells.[1,4].Value <- "Category"
   loadNotesByMarkSheet.Cells.[1,5].Value <- "Position"
   loadNotesByMarkSheet.Cells.[1,6].Value <- "Value 1"
   loadNotesByMarkSheet.Cells.[1,7].Value <- "Val 1 Distance Ft"
   loadNotesByMarkSheet.Cells.[1,8].Value <- "Val 1 Distance In"
   loadNotesByMarkSheet.Cells.[1,9].Value <- "Value 2"
   loadNotesByMarkSheet.Cells.[1,10].Value <- "Val 2 Distance Ft"
   loadNotesByMarkSheet.Cells.[1,11].Value <- "Val 2 Distance In"
   loadNotesByMarkSheet.Cells.[1,12].Value <- "Reference"
   loadNotesByMarkSheet.Cells.[1,13].Value <- "Load Case(s)"
   loadNotesByMarkSheet.Cells.[1,14].Value <- "Remarks"
   let notesByMark = job.Loads
   notesByMark
   |> Seq.iteri
       (fun i load ->
           loadNotesByMarkSheet.Cells.[i + 2, 1].Value <- ""
           loadNotesByMarkSheet.Cells.[i + 2, 2].Value <- load.ID
           loadNotesByMarkSheet.Cells.[i + 2, 3].Value <- load.Type
           loadNotesByMarkSheet.Cells.[i + 2, 4].Value <- load.Category
           loadNotesByMarkSheet.Cells.[i + 2, 5].Value <- load.Position
           loadNotesByMarkSheet.Cells.[i + 2, 6].Value <- load.Load1Value 
           loadNotesByMarkSheet.Cells.[i + 2, 7].Value <- match load.Load1DistanceFt with Some d -> box d | None -> null
           loadNotesByMarkSheet.Cells.[i + 2, 8].Value <- match load.Load1DistanceIn with Some d -> box d | None -> null
           loadNotesByMarkSheet.Cells.[i + 2, 9].Value <- match load.Load2Value with Some v -> box v | None -> null
           loadNotesByMarkSheet.Cells.[i + 2, 10].Value <- match load.Load2DistanceFt with Some d -> box d | None -> null
           loadNotesByMarkSheet.Cells.[i + 2, 11].Value <- match load.Load2DistanceIn with Some d -> box d | None -> null
           loadNotesByMarkSheet.Cells.[i + 2, 12].Value <- load.Ref |> Option.defaultValue ""
           loadNotesByMarkSheet.Cells.[i + 2, 13].Value <- load.LoadCases |> Seq.map (fun v -> v.ToString()) |> DESign.Helpers.ConvertToCommaSeparatedString
           loadNotesByMarkSheet.Cells.[i + 2, 14].Value <- load.Remarks |> Option.defaultValue "" )


   let lastRowForLoadNotesByMarkSheet = loadNotesByMarkSheet.Dimension.End.Row
   let tableRangeForLoadNotesByMarkSheet = loadNotesByMarkSheet.Cells.[1, 1, lastRowForLoadNotesByMarkSheet, 14]
   let loadNotesByMarkTableName = "tblLoadNotesByMark"
   let loadNotesByMarkTable =loadNotesByMarkSheet.Tables.Add(tableRangeForLoadNotesByMarkSheet, loadNotesByMarkTableName)
   loadNotesByMarkTable.TableStyle <- Table.TableStyles.Light8
   tableRangeForLoadNotesByMarkSheet.AutoFitColumns()

   (*
   // Load Notes By Note
   let loadNotesByNoteSheet = package.Workbook.Worksheets.Add("Load Notes By Note")
   loadNotesByNoteSheet.Cells.[1,1].Value <- "ID"
   loadNotesByNoteSheet.Cells.[1,2].Value <- "Type"
   loadNotesByNoteSheet.Cells.[1,3].Value <- "Category"
   loadNotesByNoteSheet.Cells.[1,4].Value <- "Position"
   loadNotesByNoteSheet.Cells.[1,5].Value <- "Value 1"
   loadNotesByNoteSheet.Cells.[1,6].Value <- "Val 1 Distance Ft"
   loadNotesByNoteSheet.Cells.[1,7].Value <- "Val 1 Distance In"
   loadNotesByNoteSheet.Cells.[1,8].Value <- "Value 2"
   loadNotesByNoteSheet.Cells.[1,9].Value <- "Val 2 Distance Ft"
   loadNotesByNoteSheet.Cells.[1,10].Value <- "Val 2 Distance In"
   loadNotesByNoteSheet.Cells.[1,11].Value <- "Reference"
   loadNotesByNoteSheet.Cells.[1,12].Value <- "Load Case(s)"
   loadNotesByNoteSheet.Cells.[1,13].Value <- "Remarks"
   loadNotesByNoteSheet.Cells.[1,14].Value <- "Marks With This Load"

   let notesByNote = job |> getLoadNotesByNote
   notesByNote
   |> Seq.iteri
       (fun i (load, marks) ->
           loadNotesByNoteSheet.Cells.[i + 2, 1].Value <- load.ID
           loadNotesByNoteSheet.Cells.[i + 2, 2].Value <- load.Type
           loadNotesByNoteSheet.Cells.[i + 2, 3].Value <- load.Category
           loadNotesByNoteSheet.Cells.[i + 2, 4].Value <- load.Position
           loadNotesByNoteSheet.Cells.[i + 2, 5].Value <- load.Load1Value 
           loadNotesByNoteSheet.Cells.[i + 2, 6].Value <- match load.Load1DistanceFt with Some d -> box d | None -> null
           loadNotesByNoteSheet.Cells.[i + 2, 7].Value <- match load.Load1DistanceIn with Some d -> box d | None -> null
           loadNotesByNoteSheet.Cells.[i + 2, 8].Value <- match load.Load2Value with Some v -> box v | None -> null
           loadNotesByNoteSheet.Cells.[i + 2, 9].Value <- match load.Load2DistanceFt with Some d -> box d | None -> null
           loadNotesByNoteSheet.Cells.[i + 2, 10].Value <- match load.Load2DistanceIn with Some d -> box d | None -> null
           loadNotesByNoteSheet.Cells.[i + 2, 11].Value <- load.Ref |> Option.defaultValue ""
           loadNotesByNoteSheet.Cells.[i + 2, 12].Value <- load.LoadCases |> Seq.map (fun v -> v.ToString()) |> DESign.Helpers.ConvertToCommaSeparatedString
           loadNotesByNoteSheet.Cells.[i + 2, 13].Value <- load.Remarks |> Option.defaultValue ""
           loadNotesByNoteSheet.Cells.[i + 2, 14].Value <- marks |> DESign.Helpers.ConvertToCommaSeparatedString )


   let lastRowForLoadNotesByNoteSheet = loadNotesByNoteSheet.Dimension.End.Row
   let tableRangeForLoadNotesByNoteSheet = loadNotesByNoteSheet.Cells.[1, 1, lastRowForLoadNotesByNoteSheet, 14]
   let LoadNotesByNoteTableName = "tblLoadNotesByNote"
   let loadNotesByNoteTable =loadNotesByNoteSheet.Tables.Add(tableRangeForLoadNotesByNoteSheet, LoadNotesByNoteTableName)
   loadNotesByNoteTable.TableStyle <- Table.TableStyles.Light8
   tableRangeForLoadNotesByNoteSheet.AutoFitColumns()
   *)
   package


