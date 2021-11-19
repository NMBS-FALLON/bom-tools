module DESign.BomTools.AdjustLoads

open DESign.BomTools.Domain
open DESign.Helpers



let GetSeperatedSeismicLoads (job: Job) =
    let sds =
        match Job.tryGetSds job with
        | Some sds -> sds
        | None -> failwith "No SDS found in notes."

    let requiresSeperationJoist (joist : Joist) =
        match joist.DesignationInfo with
        | Ok designationInfo ->
            match designationInfo.TotalAndLive with
            | Some _ ->
                let loads = Job.getJoistLoads joist job
                loads
                |> Seq.exists (fun l -> l.Category = "SM" && (l.LoadCases |> Seq.contains 1 || l.LoadCases = Seq.empty))
            | None -> false
        | _ -> false
                

    let requiresSeperationGirder (girder : Girder) =
        match girder.DesignationInfo with
        | Ok designationInfo ->
            let loads = Job.getGirderLoads girder job
            loads
            |> Seq.exists (fun l -> l.Category = "SM" && (l.LoadCases |> Seq.contains 1 || l.LoadCases = Seq.empty))
        | _ -> false


    

    let seperateSeismicOnJoist (joist : Joist) =
        let originalJoistLoads = Job.getJoistLoads joist job
        let lc1LoadsToLc3 =
            originalJoistLoads
            |> Seq.map
                (fun load ->
                    if  (load.LoadCases = Seq.empty || load.LoadCases |> Seq.contains 1)
                        && not (load.LoadCases |> Seq.contains 3) then
                        
                        if load.Category = "SM" then
                            let newLoadCases = 
                                load.LoadCases
                                |> Seq.filter (fun i -> i <> 1)
                                |> Seq.append (seq [3])
                            {load with LoadCases = newLoadCases}

                        else if load.Category <> "WL" && load.Category <> "IP" then
                                let newLoadCases =
                                    load.LoadCases
                                    |> Seq.append (seq [3])
                                {load with ID = joist.Mark; LoadCases = newLoadCases} 
                        else
                            load
                    else
                        load)
        let uniformSeismicAndDeadLoad =
            match joist.DesignationInfo with
            | Ok designationInfo ->
                match designationInfo.TotalAndLive with
                | Some (tl, ll) ->
                    let dl = tl - ll
                    let loadValue = 0.14 * sds * dl
                    let uniformDead =
                        Load.Create(joist.Mark, "U", "CL", "TC", dl, None, None, None, None, None, None, (seq [3]), None )
                    let uniformSeismic =
                        Load.Create(joist.Mark, "U", "SM", "TC", 0.14 * sds * dl, None, None, None, None, None, None, (seq [3]), None)
                    seq [uniformDead; uniformSeismic]
                | None ->
                    Seq.empty
            | Error e -> Seq.empty
        let originalLoadsWithOutSeismicInLc1 =
            originalJoistLoads
            |> Seq.filter (fun l -> l.Category <> "SM")
            |> Seq.map (fun l -> {l with ID = joist.Mark})

        Seq.concat [originalLoadsWithOutSeismicInLc1; lc1LoadsToLc3; uniformSeismicAndDeadLoad]

    let seperateSeismicOnGirder (girder : Girder) =
        let llNote =
            let liveLoadNote = job |> Job.tryGetLiveLoadNote girder
            match liveLoadNote with
            | Some liveLoadNote -> liveLoadNote
            | None -> failwith (sprintf "Mark %s: No Live Load Percent found" girder.Mark)

        let ll =
            match girder.DesignationInfo with
            | Ok designationInfo -> 
                match llNote with
                | Job.LiveLoadNote.Percent percentLL -> designationInfo.TotalLoad * (percentLL / 100.0)
                | Job.LiveLoadNote.Kip kipLL -> kipLL
            | _ -> 0.0


        let originalGirderLoads = Job.getGirderLoads girder job
        let lc1LoadsToLc3 =
            originalGirderLoads
            |> Seq.map
                (fun load ->
                    if load.Category = "SM"
                        && (load.LoadCases |> Seq.contains 1)
                        && not (load.LoadCases |> Seq.contains 3) then
                            let newLoadCases = 
                                load.LoadCases
                                |> Seq.filter (fun i -> i <> 1)
                                |> Seq.append (seq [3])
                            {load with LoadCases = newLoadCases}
                    else if load.Category <> "WL"
                        && load.Category <> "IP"
                        && (load.LoadCases |> Seq.contains 1)
                        && not (load.LoadCases |> Seq.contains 3) then
                            let newLoadCases =
                                load.LoadCases
                                |> Seq.append (seq [3])
                            {load with ID = girder.Mark; LoadCases = newLoadCases} 
                    else
                        load)

        let panelPointDeadLoads =
            match girder.DesignationInfo with
            | Ok designationInfo ->
                let tl = designationInfo.TotalLoad
                let dl = tl - ll
                seq [ for loc in girder.PanelLocations do
                        yield Load.Create(girder.Mark, "C", "CL", "TC", dl * 1000.0, Some loc, None, None, None, None, None, (seq [3]), None)]
            | _ -> Seq.empty

        let uniformSeismic =
            match girder.DesignationInfo with
            | Ok designationInfo ->
                let tl = designationInfo.TotalLoad
                let dl = tl - ll
                let panelSpacings =
                    let panelLocations = girder.PanelLocations;
                    let mutable i = 0 ;
                    seq [ while i < Seq.length girder.PanelLocations - 1 do
                            yield (panelLocations |> Seq.item (i + 1)) - (panelLocations |> Seq.item i) 
                            i <- i + 1 ]
                let minSpacing =
                    Seq.min panelSpacings
                let uniformDead = (dl * 1000.0) / minSpacing
                let uniformSeismic =
                    Load.Create(girder.Mark, "U", "SM", "TC", 0.14 * sds * uniformDead, None, None, None, None, None, None, (seq [3]), None)
                seq [uniformSeismic]
            | Error e -> Seq.empty

        let originalGirderLoadsWithoutSeismicInLC1 =
            originalGirderLoads
            |> Seq.filter (fun l -> l.Category <> "SM")
            |> Seq.map (fun l -> {l with ID = girder.Mark})

        Seq.concat [originalGirderLoadsWithoutSeismicInLC1; lc1LoadsToLc3 ;uniformSeismic; panelPointDeadLoads]

    
    
    let joistLoads =
        job.Joists
        |> Seq.filter requiresSeperationJoist
        |> Seq.collect seperateSeismicOnJoist

    let girderLoads =
        job.Girders
        |> Seq.filter requiresSeperationGirder
        |> Seq.collect seperateSeismicOnGirder
    
    { job with Loads = Seq.concat [job.Loads; joistLoads; girderLoads]}
        



    


    

 