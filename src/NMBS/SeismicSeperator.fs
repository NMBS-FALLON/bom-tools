module DESign.BomTools.SeismicSeperator

open System
open System.IO
open Microsoft.Office.Interop.Excel
open System.Runtime.InteropServices
open DESign.ArrayExtensions
open System.Text.RegularExpressions
open DESign.Helpers


let inline handleErrorWithStringMessage msg exp =
    try exp with | _ -> (failwith msg)

type SeperatorInfo =
    {
        SeperateSeismic : bool
        CheckInwardPressureOnGirders : bool
    }

module Load = 

    type T =
        {
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
        LoadCases : int list
        }

        member this.LoadCaseString =
            match this.LoadCases with
            | [] -> ""
            | _ -> 
                this.LoadCases
                |> List.map string
                |> List.reduce (fun s1 s2 -> s1 + "," + s2)
        
    let create(loadType, category, position, load1Value, load1DistanceFt, load1DistanceIn, load2Value, load2DistanceFt, load2DistanceIn, ref, loadcases) =
            {Type = loadType; Category = category; Position = position; Load1Value = load1Value;
                Load1DistanceFt = load1DistanceFt; Load1DistanceIn = load1DistanceIn; Load2Value = load2Value;
                Load2DistanceFt = load2DistanceFt; Load2DistanceIn = load2DistanceIn; Ref = ref; LoadCases = loadcases}

open Load

type LoadNote =
    {
    LoadNumber : string
    Load : Load.T
    }


type Panel =
    {
    Number : int
    LengthFt : float
    LengthIn : float
    }

type GirderGeometry =
    {
    Mark : String
    A_Ft : float
    A_In : float
    B_Ft : float
    B_In : float
    Panels : Panel list
    }

    member this.NumPanels =
        1 + List.length (this.Panels)

let getGirderGeometry (a2D : obj[,]) =
    let startIndex = Array2D.base1 a2D 
    let endIndex = (a2D |> Array2D.length1) - (if startIndex = 0 then 1 else 0)
    let startColIndex = Array2D.base2 a2D
    let mutable row = startIndex
    
    
    [while row <= endIndex do
        let numPanels =
            let mutable panelCounter = 0
            let markCell = a2D.[row, startColIndex]
            let panelNoCell = a2D.[row, startColIndex + 4]

            if (markCell <> null && markCell <> box "") then
                panelCounter <- 1
                let mutable continueCounting = true
                while continueCounting = true do
                    row <- row + 1
                    let markCell = a2D.[row, startColIndex]
                    let panelNoCell = a2D.[row, startColIndex + 4]
                    if (markCell = null || markCell = box "") &&
                       (panelNoCell <> null && panelNoCell <> box "") then
                        panelCounter <- panelCounter + 1
                    else
                        continueCounting <- false
                        row <- row - 1
            panelCounter
       
        let test =
           [for j = 1 to 1 do
               yield sprintf "%i" j]

        let panels =
            [for j = 1 to numPanels do
                let mark = string (a2D.[row- numPanels + 1, startColIndex])
                let n = Convert.ToInt32 (a2D.[row- numPanels + j, startColIndex + 4]) |> handleErrorWithStringMessage ( sprintf "Error with panels at mark %s" mark)
                let feet = Convert.ToDouble (a2D.[row- numPanels + j, startColIndex + 5]) |> handleErrorWithStringMessage ( sprintf "Error with panels at mark %s" mark)
                let inch = Convert.ToDouble (a2D.[row- numPanels + j, startColIndex + 6]) |> handleErrorWithStringMessage ( sprintf "Error with panels at mark %s" mark)
                for k = 1 to Convert.ToInt32 (a2D.[row- numPanels + j, startColIndex + 4]) |> handleErrorWithStringMessage ( sprintf "Error with panels at mark %s" mark) do
                    yield
                        {
                        Number = 1
                        LengthFt = Convert.ToDouble (a2D.[row- numPanels + j, startColIndex + 5]) |> handleErrorWithStringMessage ( sprintf "Error with panels at mark %s" mark)
                        LengthIn = Convert.ToDouble (a2D.[row- numPanels + j, startColIndex + 6]) |> handleErrorWithStringMessage ( sprintf "Error with panels at mark %s" mark)
                        }]
        if numPanels <> 0 then
            let mark = string (a2D.[row- numPanels + 1, startColIndex])
            let aFt = Convert.ToDouble (a2D.[row- numPanels + 1, startColIndex + 2]) |> handleErrorWithStringMessage ( sprintf "Error with panels at mark %s" mark)
            let aIn = Convert.ToDouble (a2D.[row- numPanels + 1, startColIndex + 3]) |> handleErrorWithStringMessage ( sprintf "Error with panels at mark %s" mark)
            let bFt = Convert.ToDouble (a2D.[row, startColIndex + 7]) |> handleErrorWithStringMessage ( sprintf "Error with panels at mark %s" mark)
            let bIn = Convert.ToDouble (a2D.[row, startColIndex + 8]) |> handleErrorWithStringMessage ( sprintf "Error with panels at mark %s" mark)
            yield
                 {
                 Mark = string (a2D.[row- numPanels + 1, startColIndex])
                 A_Ft = Convert.ToDouble (a2D.[row- numPanels + 1, startColIndex + 2]) |> handleErrorWithStringMessage ( sprintf "Error with panels at mark %s" mark)
                 A_In = Convert.ToDouble (a2D.[row- numPanels + 1, startColIndex + 3]) |> handleErrorWithStringMessage ( sprintf "Error with panels at mark %s" mark)
                 B_Ft = Convert.ToDouble (a2D.[row, startColIndex + 7]) |> handleErrorWithStringMessage ( sprintf "Error with panels at mark %s" mark)
                 B_In = Convert.ToDouble (a2D.[row, startColIndex + 8]) |> handleErrorWithStringMessage ( sprintf "Error with panels at mark %s" mark)
                 Panels = panels
                 }
        row <- row+ 1] 

let getPanelDim (panel : int) (girderGeom : GirderGeometry) =
    let mutable ft = girderGeom.A_Ft
    let mutable inch = girderGeom.A_In
    let mutable i = 0
    while i < panel - 1 do
        ft <- ft + girderGeom.Panels.[i].LengthFt
        inch <- inch + girderGeom.Panels.[i].LengthIn
        i <- i + 1
    ft <- ft + (inch / 12.0) - ((inch / 12.0) % 1.0)
    inch <- ((inch / 12.0) % 1.0) * 12.0
    (ft, inch)   

    
let getLoadNotes (note : string) =
    if note.Contains("(") then
        let loadNoteStart = note.IndexOf("(")
        let loadNotes = note.Substring(loadNoteStart)
        let loadNotes = loadNotes.Split([|"("; ","; ")"|], StringSplitOptions.RemoveEmptyEntries)
        let loadNotes = loadNotes |> List.ofArray
        loadNotes |> List.map (fun (s : string) -> s.Trim())
    else
        []

let getSpecialNotes (note : string) =
    if note.Contains("[") then
        let specialNotesStart = note.IndexOf("[")
        let specialNotesEnd = note.IndexOf("]")
        let specialNotes = note.Substring(specialNotesStart, specialNotesEnd + 1)
        let specialNotes = specialNotes.Split([|"["; ","; "]"|], StringSplitOptions.RemoveEmptyEntries)
        let specialNotes = specialNotes |> List.ofArray
        specialNotes |> List.map (fun (s: string) -> s.Trim())
    else
        []

type Note =
    {
    Number : string
    Text : string
    }

type Joist =
    {
    Mark : string
    JoistSize : string
    OverallLength : float
    LE_Depth : float
    RE_Depth : float
    Overall_TC_Pitch : float
    NotesString : string option
    SlopeSpecialNotes : Note List
    }

    member this.LoadNoteList =
        match (this.NotesString) with
        | Some notes -> getLoadNotes notes
        | None -> []

    member this.SpecialNoteList =
        match (this.NotesString) with
        | Some notes -> Some (getSpecialNotes notes)
        | None -> None

    member this.EndDepths =
        let slopeSpecialNote =
            match this.SpecialNoteList with
            | Some list -> 
                [for slopeSpecialNote in this.SlopeSpecialNotes do
                     if List.contains slopeSpecialNote.Number list then
                         yield slopeSpecialNote.Text]
            | None -> []

        let endDepths =
            match slopeSpecialNote with
            | [] -> 0.0, 0.0
            | _ ->
                match slopeSpecialNote.[0] with
                | Regex @"SP *: *(\d+\.?\d*)/(\d+\.?\d*)" [leDepth; reDepth] -> (float leDepth, float reDepth)
                | _ -> 0.0, 0.0 

        endDepths



    member this.UDL =
        let size = this.JoistSize
        if size.Contains("/") then
            let sizeAsArray = size.Split( [|"LH"; "K"; "/"|], StringSplitOptions.RemoveEmptyEntries)
            let TL = float sizeAsArray.[1]
            let LL = float sizeAsArray.[2]
            let DL = TL - LL
            Some(Load.create("U", "CL", "TC", DL,
                         None, None, None, None, None, None, [3]))
        else
            None       

    member this.Sds sds =
        match this.UDL with
        | Some udl -> 
            let sds = 0.14 * sds * udl.Load1Value
            Some (Load.create ("U", "SM", "TC", sds,
                          None, None, None, None, None, None, [3]))
        | None -> None



    member this.LC3Loads (loadNotes :LoadNote list) sds (seperatorInfo : SeperatorInfo) =

        let thisMarksLoads =
            loadNotes                   
            |> List.filter
                (fun note ->
                    this.LoadNoteList |> List.contains note.LoadNumber)
            |> List.map (fun note -> note.Load)

        let hasSeismic =
            thisMarksLoads
            |> Seq.exists (fun l -> l.Category = "SM" && (l.LoadCases = [] || l.LoadCases |> Seq.contains 1))
        
        let lc3Loads =
            if hasSeismic && seperatorInfo.SeperateSeismic then
                match this.UDL, (this.Sds sds) with
                | Some udl, Some sds ->
                    loadNotes
                    |> List.filter (fun note -> this.LoadNoteList |> List.contains note.LoadNumber)
                    |> List.map (fun note -> note.Load)
                    |> List.filter (fun load -> load.Category <> "WL" && load.Category <> "SM" && (load.LoadCases = [] || (load.LoadCases |> List.contains 1)))
                    |> List.map (fun load -> {load with LoadCases = [3]})
                    |> List.append [udl; sds]
                | _ -> []
            else
                []
        
        lc3Loads

type AdditionalLoadType =
    | SingleLoad of float
    | TotalOverLiveLoad of float * float

type _AdditionalJoist =
    {
    LocationFt : obj
    LocationIn : obj
    Load : AdditionalLoadType
    }

    member this.ToLoads() =
        match this.Load with
        | SingleLoad load ->
            [{
            Type = "CB"
            Category = "CL"
            Position = "TC"
            Load1Value = load * 1000.0
            Load1DistanceFt =
                match this.LocationFt with
                | v when box "P" = v -> -1.0 |> Some
                | _ -> parseObjToFloatOptionWithFailure this.LocationFt
            Load1DistanceIn =
                match this.LocationIn with
                | :? string as v ->
                    if v.Contains("#") then
                        v.Replace("#", "") |> double |> Some
                    else
                        parseObjToFloatOptionWithFailure this.LocationIn
                | _ -> parseObjToFloatOptionWithFailure this.LocationIn
            Load2Value = None
            Load2DistanceFt = None
            Load2DistanceIn = None
            Ref = Some "L-BL"
            LoadCases = []
            }]
        | TotalOverLiveLoad (totalLoad, liveLoad) ->
            let deadLoad = System.Math.Round((totalLoad - liveLoad) * 1000.0)
            let liveLoad = System.Math.Round(liveLoad * 1000.0)
            [
                {
                    Type = "C"
                    Category = "CL"
                    Position = "TC"
                    Load1Value = deadLoad
                    Load1DistanceFt =
                        match this.LocationFt with
                        | v when box "P" = v -> -1.0 |> Some
                        | _ -> parseObjToFloatOptionWithFailure this.LocationFt
                    Load1DistanceIn =
                        match this.LocationIn with
                        | :? string as v ->
                            if v.Contains("#") then
                                v.Replace("#", "") |> double |> Some
                            else
                                parseObjToFloatOptionWithFailure this.LocationIn
                        | _ -> parseObjToFloatOptionWithFailure this.LocationIn
                    Load2Value = None
                    Load2DistanceFt = None
                    Load2DistanceIn = None
                    Ref = Some "L-BL"
                    LoadCases = []
                };
                {
                    Type = "C"
                    Category = "LL"
                    Position = "TC"
                    Load1Value = liveLoad
                    Load1DistanceFt =
                        match this.LocationFt with
                        | v when box "P" = v -> -1.0 |> Some
                        | _ -> parseObjToFloatOptionWithFailure this.LocationFt
                    Load1DistanceIn =
                        match this.LocationIn with
                        | :? string as v ->
                            if v.Contains("#") then
                                v.Replace("#", "") |> double |> Some
                            else
                                parseObjToFloatOptionWithFailure this.LocationIn
                        | _ -> parseObjToFloatOptionWithFailure this.LocationIn
                    Load2Value = None
                    Load2DistanceFt = None
                    Load2DistanceIn = None
                    Ref = Some "L-BL"
                    LoadCases = []
                }
            ]



type AdditionalJoist =
    {
    Mark : string
    AdditionalJoists : Load.T list
    }

type Girder =
    {
    Mark : string
    GirderSize : string
    OverallLengthFt : float
    OverallLengthIn : float
    TcxlLengthFt : float
    TcxlLengthIn : float
    TcxrLengthFt : float
    TcxrLengthIn : float
    NotesString : string option
    AdditionalJoists : Load.T list
    GirderGeometry : GirderGeometry
    //LiveLoadUNO : string
    //LiveLoadSpecialNotes : Note List
    }

    member this.LoadNoteList =
        match (this.NotesString) with
        | Some notes -> Some (getLoadNotes notes)
        | None -> None

    member this.SpecialNoteList =
        match (this.NotesString) with
        | Some notes -> Some (getSpecialNotes notes)
        | None -> None

    member this.BaseLength =
        (this.OverallLengthFt + this.OverallLengthIn/12.0) -
         (this.TcxlLengthFt + this.TcxlLengthIn/12.0) -
          (this.TcxrLengthFt + this.TcxlLengthIn/12.0)

    member this.PDL (liveLoadUNO : string) (liveLoadSpecialNotes : Note List)=
        let isInTotalOverLive = this.GirderSize.Contains("/")

        if isInTotalOverLive then
            let size = this.GirderSize
            let sizeAsArray = size.Split( [|"G"; "BG"; "VG"; "/"; "N"; "K"|], StringSplitOptions.RemoveEmptyEntries)
            let TL = float sizeAsArray.[2]
            let LL = float sizeAsArray.[3]
            let DL = System.Math.Round(1000.0 * (TL - LL))
            DL
        else
            let size = this.GirderSize
            let sizeAsArray = size.Split( [|"G"; "BG"; "VG"; "N"; "K"|], StringSplitOptions.RemoveEmptyEntries)
            let TL = float sizeAsArray.[2]

            let N = int sizeAsArray.[1]

            (*
            let minSpace =
                let geometry = this.GirderGeometry
                let aSpace = geometry.A_Ft + geometry.A_In / 12.0
                let bSpace = geometry.B_Ft + geometry.B_In / 12.0
                let minPanelSpace = List.min (geometry.Panels |> List.map (fun geom -> geom.LengthFt + geom.LengthIn / 12.0))
                List.min [aSpace; bSpace; minPanelSpace] *)

            let liveLoadSpecialNote =
                match this.SpecialNoteList with
                | Some list -> 
                    [for liveLoadSpecialNote in liveLoadSpecialNotes do
                         if List.contains liveLoadSpecialNote.Number list then
                             yield liveLoadSpecialNote.Text]
                | None -> []

            let liveLoad =
                match liveLoadSpecialNote with
                | [] -> liveLoadUNO
                | _ -> liveLoadSpecialNote.[0] 
                     

            let LL =
                match liveLoad with
                | Regex @" *[LS] *= *(\d+\.?\d*) *[Kk] *" [value] -> float value
                | Regex @" *[LS] *= *(\d+\.?\d*) *% *" [percent] ->
                    let fraction = float percent/100.0
                    TL*fraction
                | _ -> 0.0

            let DL = 1000.0 * (TL - LL)
            DL

    member this.PanelLocations =
        let geom = this.GirderGeometry
        [for i = 1 to geom.NumPanels do yield getPanelDim i geom]

    member this.SdsLoads sds liveLoadUNO liveLoadSpecialNotes =
        let dl = this.PDL liveLoadUNO liveLoadSpecialNotes
        let geom = this.GirderGeometry
        [for i = 1 to geom.NumPanels do
            let distanceFt, distanceIn = getPanelDim i geom
            yield
                Load.create("C", "SM", "TC", dl * 0.14 * sds, Some distanceFt, Some distanceIn,
                            None, None, None, Some "L-BL", [3])]

    member this.DeadLoads liveLoadUNO liveLoadSpecialNotes =
        let dl = this.PDL liveLoadUNO liveLoadSpecialNotes
        let geom = this.GirderGeometry
        [for i = 1 to geom.NumPanels do
            let distanceFt, distanceIn = getPanelDim i geom
            yield
                Load.create("C", "CL", "TC", dl, Some distanceFt, Some distanceIn,
                             None, None, None, Some "L-BL", [3])]
        



    member this.Lc3and4Loads (loadNotes :LoadNote list) (sds : float) (liveLoadUNO : string) (liveLoadSpecialNotes : Note List) (inwardPressureUNO : string) (inwardPressureSpecialNotes : Note List) (seperatorInfo : SeperatorInfo)=
        
        let additionalJoistLoads =
            this.AdditionalJoists
            |> List.map (fun load -> {load with LoadCases = [3]})

        let thisMarksLoads =
            loadNotes                   
            |> List.filter (fun note ->
                match this.LoadNoteList with
                | Some loadNoteList -> loadNoteList |> List.contains note.LoadNumber
                | None -> false)
            |> List.map (fun note -> note.Load)

        let hasSeismic =
            thisMarksLoads
            |> Seq.exists (fun l -> l.Category = "SM" && (l.LoadCases = [] || l.LoadCases |> Seq.contains 1))

        let hasUniformInwardPressure =
            thisMarksLoads
            |> Seq.exists (fun l -> l.Type = "U" && l.Category = "IP")
                
        let sds_From_UDLs_And_UTLs =
            thisMarksLoads
            |> List.filter (fun load -> load.Type = "U" && (load.Category = "DL" || load.Category = "TL") && (load.LoadCases = [] || (load.LoadCases |> List.contains 1)))
            |> List.map (fun load ->
                let categoryFactor = if load.Category = "DL" then 1.0 else 0.5
                let load1Value = 0.14 * sds * load.Load1Value * categoryFactor
                let load2Value =
                    match (load.Load2Value) with
                    | Some v -> Some (0.14 * sds * v * categoryFactor)
                    | None -> None
                {load with Category = "SM"; Load1Value = load1Value; Load2Value = load2Value; LoadCases = [3]})
                
        let directLoadsToLC3 =
            thisMarksLoads
            |> List.filter (fun load -> load.Category <> "WL" && load.Category <> "IP" && load.Category <> "SM" && (load.LoadCases = [] || (load.LoadCases |> List.contains 1)))
            |> List.map (fun load -> {load with LoadCases = [3]})                 
                
       
        let size = this.GirderSize
        let sizeAsArray = size.Split( [|"G"; "BG"; "VG"; "N"; "K"|], StringSplitOptions.RemoveEmptyEntries)
        let isInTotalOverLive = this.GirderSize.Contains("/")
        let TL = 
            if isInTotalOverLive then
                let size = this.GirderSize
                let sizeAsArray = size.Split( [|"G"; "BG"; "VG"; "/"; "N"; "K"|], StringSplitOptions.RemoveEmptyEntries)
                let TL = float sizeAsArray.[2]
                TL
            else
                float sizeAsArray.[2]
                
        let N = float sizeAsArray.[1]
        (*
        let minSpace =
            let geometry = this.GirderGeometry
            let aSpace = geometry.A_Ft + geometry.A_In / 12.0
            let bSpace = geometry.B_Ft + geometry.B_In / 12.0
            let minPanelSpace = List.min (geometry.Panels |> List.map (fun geom -> geom.LengthFt + geom.LengthIn / 12.0))
            List.min [aSpace; bSpace; minPanelSpace] *)


        let liveLoadSpecialNote =
            match this.SpecialNoteList with
            | Some list -> 
                [for liveLoadSpecialNote in liveLoadSpecialNotes do
                     if List.contains liveLoadSpecialNote.Number list then
                         yield liveLoadSpecialNote.Text]
            | None -> []

        let liveLoad =
            match liveLoadSpecialNote with
            | [] -> liveLoadUNO
            | _ -> liveLoadSpecialNote.[0] 
                 
          
        let LL =
            if isInTotalOverLive then
                let size = this.GirderSize
                let sizeAsArray = size.Split( [|"G"; "BG"; "VG"; "/"; "N"; "K"|], StringSplitOptions.RemoveEmptyEntries)
                let LL = float sizeAsArray.[3]
                LL
            else if (hasUniformInwardPressure && seperatorInfo.CheckInwardPressureOnGirders) || (seperatorInfo.SeperateSeismic && hasSeismic) then
                match liveLoad with
                | Regex @" *[LS] *= *(\d+\.?\d*) *[Kk] *" [value] -> float value
                | Regex @" *[LS] *= *(\d+\.?\d*) *% *" [percent] ->
                    let fraction = float percent/100.0
                    TL*fraction
                | _ -> failwith (sprintf "Mark %s: This mark has IP and/or seismic loading but is missing an LL%% note." this.Mark )
            else
                0.0

        let inwardPressureSpecialNote =
            match this.SpecialNoteList with
            | Some list -> 
                [for inwardPressureSpecialNote in inwardPressureSpecialNotes do
                     if List.contains inwardPressureSpecialNote.Number list then
                         yield inwardPressureSpecialNote.Text]
            | None -> []

        let inwardPressure =
            match inwardPressureSpecialNote with
            | [] -> inwardPressureUNO
            | _ -> inwardPressureSpecialNote.[0] 
                 

        let IP =
            if (seperatorInfo.CheckInwardPressureOnGirders && hasUniformInwardPressure) then
                match inwardPressure with
                | Regex @" *IP *= *(\d+\.?\d*) *[Kk] *" [value] -> float value
                | Regex @" *IP *= *(\d+\.?\d*) *% *" [percent] ->
                    let fraction = float percent/100.0
                    TL*fraction
                | _ -> failwith (sprintf "Mark %s: This mark has IP loading but is missing an IP%% note." this.Mark)

            else
                0.0

        let llPercent = LL / TL
        let ipPercent = IP / TL
        let dlPercent = (TL - LL) / TL
        let TL1 = 1.0
        let TL2 = dlPercent + 0.75*(llPercent + ipPercent)
        let TL3 = 0.85 + ipPercent



        let liveLoads =
            [ for panelFt, panelIn in this.PanelLocations do
                yield Load.create("C", "LL", "TC", LL * 1000.0, (Some panelFt), (Some panelIn), None, None, None, Some "L-BL", [4])]

        
        
        let requiresLc3Loads = hasSeismic && seperatorInfo.SeperateSeismic
        let requiresLc4Loads = (hasUniformInwardPressure && seperatorInfo.CheckInwardPressureOnGirders) && not (TL1 > TL2 || TL3 > TL2)

        let lc3Loads =
            if requiresLc3Loads then
                []
                |> List.append sds_From_UDLs_And_UTLs
                |> List.append (this.SdsLoads sds liveLoadUNO liveLoadSpecialNotes)
                |> List.append (this.DeadLoads liveLoadUNO liveLoadSpecialNotes)
                |> List.append additionalJoistLoads
                |> List.append directLoadsToLC3
                |> List.map
                    (fun load ->
                        let load1Value = Math.Ceiling load.Load1Value

                        let load2Value =
                            match load.Load2Value with
                            | Some v -> Some (Math.Ceiling v)
                            | None -> None

                        {load with Load1Value = load1Value; Load2Value = load2Value})
            else
                []

        let lc1ToLC4 =
            thisMarksLoads
            |> List.filter (fun load -> load.Category <> "SM" && (load.LoadCases = [] || (load.LoadCases |> List.contains 1)))
            |> List.filter (fun load -> load.Load1Value > 0.0)
            |> List.map (fun load -> {load with LoadCases = [4]})   

        let lc4Loads =
            if requiresLc4Loads then
                []
                |> List.append liveLoads
                |> List.append (this.DeadLoads liveLoadUNO liveLoadSpecialNotes)
                |> List.append additionalJoistLoads
                |> List.append lc1ToLC4
                |> List.map
                    (fun load ->
                        let load1Value = Math.Ceiling load.Load1Value
                    
                        let load2Value = load.Load2Value |> Option.map Math.Ceiling

                        {load with Load1Value = load1Value; Load2Value = load2Value; LoadCases = [4]})
            else
                []
           
            

        List.concat [lc3Loads ; lc4Loads]
            


type BOM = 
    {
    GeneralNotes : Note list
    SpecialNotes : Note list
    Joists : Joist list
    Girders : Girder list
    }   



module CleanBomInfo =

    let nullableToOption<'T> value =
        match (box value) with
        | null  -> None
        | value when value = (box "") -> None
        | _ -> Some ((box value) :?> 'T)

    let toDouble (s : obj) =
            match (box s) with
            | v when v = (box "") -> 0.0
            | _ -> Convert.ToDouble(s)




    module CleanLoads =

        let getLoadCases (loadCaseString: string) =
            let loadNotes = loadCaseString.Split([|","|], StringSplitOptions.RemoveEmptyEntries)
            let loadNotes = loadNotes
                            |> List.ofArray
                            |> List.map (fun string -> System.Int32.Parse(string.Trim()))
            loadNotes


        let getLoadFromArraySlice (a : obj []) =
            {
            Type = string a.[1]
            Category = string  a.[2]
            Position = string a.[3]
            Load1Value = handleResultWithFailure (floatFromObj a.[5])
            Load1DistanceFt =  parseObjToFloatOptionWithFailure a.[6]
            Load1DistanceIn = parseObjToFloatOptionWithFailure a.[7]
            Load2Value = parseObjToFloatOptionWithFailure a.[8]
            Load2DistanceFt = parseObjToFloatOptionWithFailure a.[9]
            Load2DistanceIn = parseObjToFloatOptionWithFailure a.[10]
            Ref = stringOptionFromObj  a.[11]
            LoadCases = getLoadCases (string a.[12])
            }

        let getLoadNotesFromArray (a2D : obj[,]) =
            let mutable startRowIndex = Array2D.base1 a2D 
            let endIndex = (a2D |> Array2D.length1) - (if startRowIndex = 0 then 1 else 0)
            let startColIndex = Array2D.base2 a2D
            let loadNotes = 
                let mutable loadNumber = ""
                [for currentIndex = startRowIndex to endIndex do
                    if a2D.[currentIndex, startColIndex + 1] <> null && a2D.[currentIndex, startColIndex + 1] <> (box "") then
                        if a2D.[currentIndex, startColIndex] <> null && a2D.[currentIndex, startColIndex] <> (box "") then
                            loadNumber <- (string a2D.[currentIndex, startColIndex]).Trim()
                        let load =
                            try
                                getLoadFromArraySlice a2D.[currentIndex, *]
                            with
                                | Failure(msg) -> failwith (sprintf "Issue with load line %i; %s" currentIndex msg)
                                
                        yield {LoadNumber = loadNumber; Load = getLoadFromArraySlice a2D.[currentIndex, *]}]
            loadNotes

    module CleanNotes =

        let getNotesFromArray (a2D : obj [,]) =
            let mutable startRowIndex = Array2D.base1 a2D 
            let endIndex = (a2D |> Array2D.length1) - (if startRowIndex = 0 then 1 else 0)
            let startColIndex = Array2D.base2 a2D
            let notes : Note list =
                let mutable currentIndex = startRowIndex
                [while currentIndex <= endIndex do
                    let mutable noteNumber = ""
                    let mutable note = ""
                    let mutable additionalLines = 0
                    if a2D.[currentIndex, startColIndex] <> null && a2D.[currentIndex, startColIndex] <> (box "") then
                        noteNumber <- string a2D.[currentIndex, startColIndex]
                        note <- string a2D.[currentIndex, startColIndex + 1]
                        while (currentIndex + additionalLines + 1 < endIndex && (a2D.[currentIndex + additionalLines + 1, startColIndex] = null || a2D.[currentIndex + additionalLines + 1,startColIndex] = (box ""))) do
                            additionalLines <- additionalLines + 1
                            note <- String.concat " " [note; string a2D.[currentIndex + additionalLines, startColIndex + 1]]

                        yield
                            {
                            Number = noteNumber
                            Text = note
                            }
                    currentIndex <- currentIndex + 1 + additionalLines]
            notes

    module CleanJoists =

        let getJoistsFromArray (a2D : obj [,]) =
            let mutable startRowIndex = Array2D.base1 a2D 
            let endIndex = (a2D |> Array2D.length1) - (if startRowIndex = 0 then 1 else 0)
            let startColIndex = Array2D.base2 a2D
            let joists : Joist list =
                [for currentIndex = startRowIndex to endIndex do
                    if a2D.[currentIndex, startColIndex] <> null && a2D.[currentIndex, startColIndex] <> (box "") then
                        yield
                            {
                            Mark = string a2D.[currentIndex, startColIndex]
                            JoistSize = string a2D.[currentIndex, startColIndex + 2]
                            NotesString = nullableToOption<string> a2D.[currentIndex, startColIndex + 26]
                            OverallLength = 
                                let feet = toDouble (a2D.[currentIndex, startColIndex + 3])
                                let inches = toDouble (a2D.[currentIndex, startColIndex + 4])
                                feet + inches/12.0
                            LE_Depth = 0.0
                            RE_Depth = 0.0
                            Overall_TC_Pitch = toDouble (a2D.[currentIndex, startColIndex + 186])
                            SlopeSpecialNotes = []
                            }]
            joists

        let addSlopeNotesToJoists (joists: Joist list, slopeSpecialNotes : Note list) =
            [for joist in joists do
                yield {joist with SlopeSpecialNotes = slopeSpecialNotes}]



    module CleanGirders =

        let getAdditionalJoistsFromArraySlice (a : obj [])  =
            let mutable col = 16
            [while col <= 28 do
                if (a.[col + 2] <> null && a.[col + 2] <> (box "")) then
                    if (a.[col] <> null && a.[col] <> (box "")) || (a.[col + 1] <> null && a.[col + 1] <> (box "")) then
                        let additionalJoist =
                            let loadString = string (a.[col + 2])
                            let load =
                                if loadString.Contains("/") then
                                    let loadStringArray = loadString.Split([|'/'|], StringSplitOptions.RemoveEmptyEntries)
                                    let totalLoad = float loadStringArray.[0]
                                    let liveLoad = float loadStringArray.[1]
                                    TotalOverLiveLoad(totalLoad, liveLoad)
                                else
                                    SingleLoad(System.Convert.ToDouble(loadString))

                            {
                            LocationFt = string a.[col]
                            LocationIn = string a.[col + 1]
                            Load = load
                            }
                        yield additionalJoist.ToLoads()
                col <- col + 4] |> List.concat

        let getAdditionalJoistsFromArray (a2D : obj [,]) =
            let mutable startRowIndex = Array2D.base1 a2D
            let colIndex = Array2D.base2 a2D
            let endIndex = (a2D |> Array2D.length1) - (if startRowIndex = 0 then 1 else 0)
            let mutable mark = ""
            let additionalJoists : AdditionalJoist list =
                [for currentIndex = startRowIndex to endIndex do
                    mark <-
                        if (isNull a2D.[currentIndex, colIndex] || (string a2D.[currentIndex, colIndex]) = "") then
                            mark
                        else
                            (string a2D.[currentIndex, colIndex])                       
                    yield
                        {
                        Mark = mark
                        AdditionalJoists = getAdditionalJoistsFromArraySlice a2D.[currentIndex, *] |> handleErrorWithStringMessage ( sprintf "Error with additional joist information at mark %s" mark)
                        } ]
                |> List.filter (fun row -> row.Mark <> "" && not (Seq.isEmpty row.AdditionalJoists))
            additionalJoists

        let getGirders (sheet1 : obj [,], sheet2 : obj [,]) =
            let mutable startIndex = Array2D.base1 sheet1
            let colIndex = Array2D.base2 sheet1
            let endIndex = (sheet1 |> Array2D.length1) - (if startIndex = 0 then 1 else 0)
            let allGirderGeometry = getGirderGeometry sheet2
            let girders : Girder list =
                [for currentIndex = startIndex to endIndex do
                    if sheet1.[currentIndex, colIndex] <> null && sheet1.[currentIndex, colIndex] <> (box "") then
                        let mark = string sheet1.[currentIndex, colIndex]
                        let geometry = allGirderGeometry |> List.find (fun geom -> geom.Mark = mark)
                        yield
                            {
                            Mark = mark
                            GirderSize = string sheet1.[currentIndex, colIndex + 2]
                            OverallLengthFt = toDouble(sheet1.[currentIndex, colIndex + 3])
                            OverallLengthIn = toDouble(sheet1.[currentIndex, colIndex + 4])
                            TcxlLengthFt = toDouble(sheet1.[currentIndex, colIndex + 6])
                            TcxlLengthIn = toDouble(sheet1.[currentIndex, colIndex + 7])
                            TcxrLengthFt = toDouble(sheet1.[currentIndex, colIndex + 9])
                            TcxrLengthIn = toDouble(string sheet1.[currentIndex, colIndex + 10])
                            NotesString =  nullableToOption<string> sheet1.[currentIndex, colIndex + 25]
                            AdditionalJoists = []
                            GirderGeometry = geometry
                            }]
            girders



        let addAdditionalJoistLoadsToGirders (girders: Girder list, additionalJoists : AdditionalJoist list) =
            [for girder in girders do
                let additionalJoistsOnGirder = additionalJoists |> List.filter (fun addJoist -> addJoist.Mark = girder.Mark)
                let additionalLoads =
                    [for addJoist in additionalJoistsOnGirder do
                        for load in addJoist.AdditionalJoists do
                            let mutable locationFt = load.Load1DistanceFt
                            let mutable locationIn = load.Load1DistanceIn
        
                            if (locationFt) = Some -1.0 then
                                let panel = System.Int32.Parse((string locationIn.Value))
                                let ft, inch = getPanelDim panel girder.GirderGeometry
                                locationFt <- Some ft
                                locationIn <- Some inch                   
                            yield {load with Load1DistanceFt = locationFt; Load1DistanceIn = locationIn}] 
                let additionalJoists = girder.AdditionalJoists |> List.append additionalLoads
                yield {girder with AdditionalJoists = additionalJoists}]

                
let saveWorkbook (title : string) (suffix : string) (workbook : Workbook) =
        let title = title.ToUpper().Replace(".XLSM", sprintf " (%s).XLSM" suffix)
        let title = title.ToUpper().Replace(".XLSX", sprintf " (%s).XLSM" suffix)
        workbook.SaveAs(title)

let removeProtection (reportPath: string) =
    let tempExcelApp = new Microsoft.Office.Interop.Excel.ApplicationClass(Visible = true)
    tempExcelApp.DisplayAlerts = false |> ignore
    tempExcelApp.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable |> ignore
    let workbooks = tempExcelApp.Workbooks
    let mutable workbook = workbooks.Add()

    //let bom = tempExcelApp.Workbooks.Open(bomPath)
    try 
        tempExcelApp.DisplayAlerts <- false
        let tempReportPath = System.IO.Path.GetTempFileName()      
        File.Delete(tempReportPath)
        File.Copy(reportPath, tempReportPath)
        workbook <- workbooks.Open(tempReportPath)
        tempExcelApp.EnableEvents <- false

        tempExcelApp.EnableEvents <- true
    
        workbook |> saveWorkbook reportPath "NO PROTECTION"

    finally
        workbook.Close(false)
        Marshal.ReleaseComObject(workbook) |> ignore
        System.GC.Collect() |> ignore
        workbooks.Close()
        Marshal.ReleaseComObject(workbooks) |> ignore
        tempExcelApp.Quit()
        Marshal.ReleaseComObject(tempExcelApp) |> ignore
        System.GC.Collect() |> ignore    

let getAllInfo (reportPath:string) getInfoFunction modifyWorkbookFunctions =
    let tempExcelApp = new Microsoft.Office.Interop.Excel.ApplicationClass(Visible = true)
    tempExcelApp.DisplayAlerts = false |> ignore
    tempExcelApp.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable |> ignore
    let workbooks = tempExcelApp.Workbooks
    let mutable workbook = workbooks.Add()
    
    //let bom = tempExcelApp.Workbooks.Open(bomPath)
    try 
        tempExcelApp.DisplayAlerts <- false
        let tempReportPath = System.IO.Path.GetTempFileName()      
        File.Delete(tempReportPath)
        File.Copy(reportPath, tempReportPath)
        workbook <- workbooks.Open(tempReportPath)
        let info = getInfoFunction workbook
        tempExcelApp.EnableEvents <- false
        for modifyWorkbookFunction in modifyWorkbookFunctions do
            modifyWorkbookFunction workbook info

        tempExcelApp.EnableEvents <- true
        
        workbook |> saveWorkbook reportPath "IMPORT"

        info
    finally
        workbook.Close(false)
        Marshal.ReleaseComObject(workbook) |> ignore
        System.GC.Collect() |> ignore
        workbooks.Close()
        Marshal.ReleaseComObject(workbooks) |> ignore
        tempExcelApp.Quit()
        Marshal.ReleaseComObject(tempExcelApp) |> ignore
        System.GC.Collect() |> ignore           

let getInfo (bom: Workbook) =

    let workSheetNames = [for sheet in bom.Worksheets -> (sheet :?> Worksheet).Name] 

    let loads =
        let loadSheetNames = workSheetNames |> List.filter (fun name -> name.Contains("L ("))
        if (List.isEmpty loadSheetNames) then
            []
        else
            let arrayList =
                seq [for sheet in bom.Worksheets do
                        let sheet = (sheet :?> Worksheet)
                        if sheet.Name.Contains("L (") then
                            let loads = sheet.Range("A14","M55").Value2 :?> obj [,]
                            yield loads]   
            let loadsAsArray = Array2D.joinMany (Array2D.joinByRows) arrayList
            let loads = CleanBomInfo.CleanLoads.getLoadNotesFromArray loadsAsArray
            loads

    let generalNotes =
        let generalNotesNames = workSheetNames |> List.filter (fun name -> name.Contains("P ("))
        if (List.isEmpty generalNotesNames) then
            []
        else
            let arrayList =
                seq [for sheet in bom.Worksheets do
                        let sheet = (sheet :?> Worksheet)
                        if sheet.Name.Contains("P (") then
                            yield sheet.Range("A8", "H47").Value2 :?> obj [,]]
            let notesAsArray = Array2D.joinMany (Array2D.joinByRows) arrayList
            let notes = CleanBomInfo.CleanNotes.getNotesFromArray notesAsArray
            notes

    let isLiveLoadNote (note: string) =
            Regex.IsMatch(note, "[LS] *= *(\d+\.?\d*) *([Kk%]) *")

    

    let liveLoadUNO =
        let liveLoadNotes = generalNotes |> List.filter (fun note -> isLiveLoadNote note.Text)
        match liveLoadNotes with
        | [] -> ""
        | _ -> liveLoadNotes.[0].Text

    let isInwardPressureNote (note : string) =
        Regex.IsMatch(note, "IP *= *(\d+\.?\d*) *([Kk%]) *")

    let inwardPressureUNO =
        let inwardPressureNotes = generalNotes |> List.filter (fun note -> isInwardPressureNote note.Text)
        match inwardPressureNotes with
        | [] -> ""
        | _ -> inwardPressureNotes.[0].Text

    let isSlopeLoadNote (note: string) =
        Regex.IsMatch(note, "SP *: *(\d+\.?\d*)/(\d+\.?\d*)")



    let SDS =
        let isSDSNote (note : string) =
            Regex.IsMatch(note, "[Ss][Dd][Ss] *= *(\d+\.?\d*) *")
        let sdsNotes = generalNotes |> List.filter (fun note -> isSDSNote note.Text)
        let sds =
            match sdsNotes with
            | [] -> failwith "No SDS value found! Please specify an SDS value using the syntax \"SDS = {Numerical Value}"
            | _ -> 
                let sdsNote = sdsNotes.[0]
                match sdsNote.Text with
                | Regex @"[Ss][Dd][Ss] *= *(\d+\.?\d*) *" [sds] -> float sds
                | _ -> failwith "No SDS value found! Please specify an SDS value using the syntax \"SDS = {Numerical Value}"
        sds
                
    

    let specialNotes =
        let specialNotesNames = workSheetNames |> List.filter (fun name -> name.Contains("N ("))
        if (List.isEmpty specialNotesNames) then
            []
        else
            let arrayList =
                seq [for sheet in bom.Worksheets do
                        let sheet = (sheet :?> Worksheet)
                        if sheet.Name.Contains("N (") then
                            yield sheet.Range("A13", "J51").Value2 :?> obj [,]]
            let notesAsArray = Array2D.joinMany (Array2D.joinByRows) arrayList
            let notes = CleanBomInfo.CleanNotes.getNotesFromArray notesAsArray
            notes

    let slopeSpecialNotes =
        (specialNotes |> List.filter (fun note -> isSlopeLoadNote note.Text))

    let liveLoadSpecialNotes =
        (specialNotes |> List.filter (fun note -> isLiveLoadNote note.Text))

    let inwardPressureSpecialNotes =
        (specialNotes |> List.filter (fun note -> isInwardPressureNote note.Text))

    

    let joists =
        let joistSheetNames = workSheetNames |> List.filter (fun name -> name.Contains("J ("))
        if (List.isEmpty joistSheetNames) then
            []
        else
            let arrayList =
                seq [for sheet in bom.Worksheets do
                        let sheet = (sheet :?> Worksheet)
                        if sheet.Name.Contains("J (") then
                            if (sheet.Range("A21").Value2 :?> string) = "MARK" then
                                yield sheet.Range("A23","GF40").Value2 :?> obj [,]
                            else
                                yield sheet.Range("A16", "GF45").Value2 :?> obj [,]]

            let joistsAsArray = Array2D.joinMany (Array2D.joinByRows) arrayList
            let joists = CleanBomInfo.CleanJoists.getJoistsFromArray joistsAsArray
            let joistsWithSlopeNotes = CleanBomInfo.CleanJoists.addSlopeNotesToJoists (joists, slopeSpecialNotes)
            joists

    let girders =
        let girderSheetNames = workSheetNames |> List.filter (fun name -> name.Contains("G ("))
        if (List.isEmpty girderSheetNames) then
            []
        else
            let arrayList1 =
                seq [for sheet in bom.Worksheets do
                        let sheet = (sheet :?> Worksheet)
                        if sheet.Name.Contains("G (") then
                            if (sheet.Range("A26").Value2 :?> string) = "MARK" then
                                yield sheet.Range("A28","AA45").Value2 :?> obj [,]
                            else
                                yield sheet.Range("A14", "AA45").Value2 :?> obj [,]]
            let arrayList2 =
                seq [for sheet in bom.Worksheets do
                        let sheet = (sheet :?> Worksheet)
                        if sheet.Name.Contains("G (") then
                            yield sheet.Range("AB14","BG45").Value2 :?> obj [,]]
            let girdersAsArray = Array2D.joinMany (Array2D.joinByRows) arrayList1
            let additionalJoistsAsArray = Array2D.joinMany (Array2D.joinByRows) arrayList2
            let girders = CleanBomInfo.CleanGirders.getGirders (girdersAsArray, additionalJoistsAsArray)
            
            let additionalJoists =
                CleanBomInfo.CleanGirders.getAdditionalJoistsFromArray additionalJoistsAsArray

            let girders = (CleanBomInfo.CleanGirders.addAdditionalJoistLoadsToGirders (girders, additionalJoists))
            girders

    (joists, girders, loads, SDS, liveLoadUNO, liveLoadSpecialNotes, inwardPressureUNO, inwardPressureSpecialNotes)


module Modifiers =
    let seperateSeismic (seperatorInfo : SeperatorInfo) (bom : Workbook) (bomInfo : Joist list * Girder list * LoadNote list * float * string * Note list * string * Note list): Unit =

        bom.Unprotect()
        //for sheet in bom.Worksheets do
        //    let sheet = (sheet :?> Worksheet)
        //    sheet.Unprotect("AAABBBBBABA-")
    
        let workSheetNames = [for sheet in bom.Worksheets -> (sheet :?> Worksheet).Name]


        let switchSmToLc3 (a2D : obj [,]) =
            let startRow = Array2D.base1 a2D
            let endRow = (Array2D.length1 a2D) - (if startRow = 0 then 1 else 0)
            let startCol = Array2D.base2 a2D
            for currentIndex = startRow to endRow do
                let lc = (string a2D.[currentIndex, startCol + 12]).Trim()
                if a2D.[currentIndex, startCol + 2] = (box "SM") && lc = "" then
                    a2D.[currentIndex, startCol + 12] <- box "3"
                if a2D.[currentIndex, startCol + 2] = (box "SM") && (lc.Contains("1")) then
                    a2D.[currentIndex, startCol + 12] <- box (lc.Replace("1", "3"))

    

        let changeSmLoadsToLC3() =
            let loadSheetNames = workSheetNames |> List.filter (fun name -> name.Contains("L ("))
            if (List.isEmpty loadSheetNames) then
                ()
            else
                for sheet in bom.Worksheets do
                    let sheet = (sheet :?> Worksheet)
                    if sheet.Name.Contains("L (") then
                        let loads = sheet.Range("A14","M55").Value2 :?> obj [,]
                        switchSmToLc3 loads
                        sheet.Range("A14", "M55").Value2 <- loads 

        let addLoadNote (mark : string) (note : string) =
            if (mark.Length > 0 && note.Length > 0) then
                let loadNote = "S" + mark
                if (String.exists (fun c -> c = '(') note) then
                    let insertLocation = note.IndexOf(")")
                    let newNote = note.Substring(0, insertLocation) + ", " + loadNote + ")"
                    newNote
                else
                    let newNote = note + " (" + loadNote + ")"
                    newNote
            else
               ""
(*
        let removeLL_FromGirder (mark : string) (designation: string) =
            if (mark.Length > 0 && designation.Length > 0) then
                let designationArray = designation.Split([|'/'; 'K'|], StringSplitOptions.RemoveEmptyEntries)
                let newDesignation =
                    if Array.length designationArray = 3 then
                        Some (designationArray.[0] + "K" + designationArray.[2])
                    else
                        None
                match newDesignation with
                | Some _ -> ()
                | None -> System.Windows.Forms.MessageBox.Show(sprintf "Mark %s is not in TL/LL format; please fix" mark) |> ignore; ()
                match newDesignation with
                | Some s -> s
                | _ -> designation
            else
                ""
*)

        let addLC3LoadsToLoadNotes() =
            let joists, girders, loads, SDS, liveLoadUNO, liveLoadSpecialNotes, inwardPressureUNO, inwardPressureSpecialNotes = bomInfo
            let joistsWithLC3Loads = joists |> List.filter (fun joist -> List.isEmpty (joist.LC3Loads loads SDS seperatorInfo) = false)
            let joistSheetNames = workSheetNames |> List.filter (fun name -> name.Contains ("J ("))
            if (List.isEmpty joistSheetNames) then ()
            else
                for sheet in bom.Worksheets do
                    let sheet = (sheet :?> Worksheet)
                    if sheet.Name.Contains("J (") then
                        let array =
                            if (sheet.Range("A21").Value2 :?> string) = "MARK" then
                                sheet.Range("A23","AA40").Value2 :?> obj [,]
                            else
                                sheet.Range("A16", "AA45").Value2 :?> obj [,]
                        let startRowIndex = Array2D.base1 array
                        let endRowIndex = (array |> Array2D.length1) - (if startRowIndex = 0 then 1 else 0) 
                        let colIndex = Array2D.base2 array
                    
                        for i = startRowIndex to endRowIndex do
                            let joistMarksWithLC3Loads =
                                joistsWithLC3Loads |> List.map (fun joist -> joist.Mark)
                            let mark = string array.[i, colIndex]
                            if (joistMarksWithLC3Loads |> List.contains mark) then
                                array.[i, colIndex + 26] <- box (addLoadNote mark (string array.[i, colIndex + 26]))
                        if (sheet.Range("A21").Value2 :?> string) = "MARK" then
                            sheet.Range("AA23","AA40").Value2 <- array.[*,array.GetLength(1)..] ////////////////////////////////////////////////////////////////////////////
                        else
                            sheet.Range("AA16", "AA45").Value2 <- array.[*,array.GetLength(1)..] ///////////////////////////////////////////////////////////////////////////

            let girdersWithLC3Loads = girders |> List.filter (fun girder -> List.isEmpty (girder.Lc3and4Loads loads SDS liveLoadUNO liveLoadSpecialNotes inwardPressureUNO inwardPressureSpecialNotes seperatorInfo) = false)
            let girderWorksheetNames = workSheetNames |> List.filter (fun name -> name.Contains ("G ("))
            if (List.isEmpty girderWorksheetNames) then ()
            else
                for sheet in bom.Worksheets do
                    let sheet = (sheet :?> Worksheet)
                    if sheet.Name.Contains("G (") then
                        let array =
                            if (sheet.Range("A26").Value2 :?> string) = "MARK" then
                                sheet.Range("A28","Z45").Value2 :?> obj [,]
                            else
                                sheet.Range("A14", "Z45").Value2 :?> obj [,]

                        let startRowIndex = Array2D.base1 array
                        let endRowIndex = (array |> Array2D.length1) - (if startRowIndex = 0 then 1 else 0) 
                        let colIndex = Array2D.base2 array

                        for i = startRowIndex to endRowIndex do
                            let girderMarksWithLC3Loads =
                                girdersWithLC3Loads |> List.map (fun girder -> girder.Mark)
                            let mark = string array.[i, colIndex]
                            //array.[i, colIndex + 2] <- box (removeLL_FromGirder mark (string array.[i, colIndex + 2]))
                            if (girderMarksWithLC3Loads |> List.contains mark) then
                                array.[i, colIndex + 25] <- box (addLoadNote mark (string array.[i, colIndex + 25]))
                            if (mark <> null && mark <> "" && array.[i, colIndex + 2].ToString().Contains("/") = false) then
                                let girder = girders |> List.find (fun g -> g.Mark = mark)
                                let size = girder.GirderSize
                                let sizeAsArray = size.Split( [|"N";"K"|], StringSplitOptions.RemoveEmptyEntries)
                                let totalLoad = float sizeAsArray.[1]
                                let liveLoad = (System.Math.Ceiling((totalLoad * 1000.0 - (girder.PDL liveLoadUNO liveLoadSpecialNotes)) / 100.0) * 100.0) / 1000.0
                                let netUplift = if sizeAsArray.Length = 2 then "" else sizeAsArray.[2]
                                let newDesignation = sizeAsArray.[0] + "N" + sizeAsArray.[1] + "/" + (string liveLoad) + "K" + netUplift
                                array.[i, colIndex + 2] <- box newDesignation
                        if (sheet.Range("A26").Value2 :?> string) = "MARK" then
                            sheet.Range("C28", "C45").Value2 <- array.[*, (colIndex + 2)..]
                            sheet.Range("Z28","Z45").Value2 <- array.[*,array.GetLength(1)..]              ///////////////////////////////////////////////////////////////////
                        else
                            sheet.Range("C14", "C45").Value2 <- array.[*, (colIndex + 2)..]
                            sheet.Range("Z14", "Z45").Value2 <- array.[*,array.GetLength(1)..]          ////////////////////////////////////////////////////////////////////


        let addLC3Loads() =

            let joists, girders, loads, SDS, liveLoadUNO, liveLoadSpecialNotes, inwardPressureUNO, inwardPressureSpecialNotes = bomInfo

            let joistsWithLC3Loads = joists |> List.filter (fun joist -> List.isEmpty (joist.LC3Loads loads SDS seperatorInfo) = false)
            let girdersWithLC3Loads = girders |> List.filter (fun girder -> List.isEmpty (girder.Lc3and4Loads loads SDS liveLoadUNO liveLoadSpecialNotes inwardPressureUNO inwardPressureSpecialNotes seperatorInfo) = false)

            let getJoistLoadArray (joist : Joist) = 
                let loadArray =
                    let mutable markAdded = false
                    [for load in (joist.LC3Loads loads SDS seperatorInfo) do
                         let refObj =
                            match load.Ref with
                            | Some r -> box r
                            | None -> null
                         if markAdded = false then
                             markAdded <- true
                             yield 
                                [|box(sprintf "%s%s" "S" joist.Mark);
                                  box load.Type;
                                  box load.Category;
                                  box load.Position;
                                  null;
                                  box load.Load1Value;
                                  floatOptionToObj load.Load1DistanceFt;
                                  floatOptionToObj load.Load1DistanceIn;
                                  floatOptionToObj load.Load2Value;
                                  floatOptionToObj load.Load2DistanceFt;
                                  floatOptionToObj load.Load2DistanceIn;
                                  refObj;
                                  box(load.LoadCaseString)|]
                         else
                             yield
                                  [|null;
                                    box load.Type;
                                    box load.Category;
                                    box load.Position;
                                    null;
                                    box load.Load1Value;
                                    floatOptionToObj load.Load1DistanceFt;
                                    floatOptionToObj load.Load1DistanceIn;
                                    floatOptionToObj load.Load2Value;
                                    floatOptionToObj load.Load2DistanceFt;
                                    floatOptionToObj load.Load2DistanceIn;
                                    refObj;
                                    box(load.LoadCaseString)|]] |> array2D
                loadArray

            let getGirderLoadArray (girder : Girder) = 
                let loadArray =
                    let mutable markAdded = false
                    [for load in (girder.Lc3and4Loads loads SDS liveLoadUNO liveLoadSpecialNotes inwardPressureUNO inwardPressureSpecialNotes seperatorInfo) do
                         let refObj =
                            match load.Ref with
                            | Some r -> box r
                            | None -> null
                         if markAdded = false then
                             markAdded <- true
                             yield 
                                [|box(sprintf "%s%s" "S" girder.Mark);
                                  box load.Type;
                                  box load.Category;
                                  box load.Position;
                                  null;
                                  box load.Load1Value;
                                  floatOptionToObj load.Load1DistanceFt;
                                  floatOptionToObj load.Load1DistanceIn;
                                  floatOptionToObj load.Load2Value;
                                  floatOptionToObj load.Load2DistanceFt;
                                  floatOptionToObj load.Load2DistanceIn;
                                  refObj;
                                  box(load.LoadCaseString)|]
                         else
                             yield
                                  [|null;
                                    box load.Type;
                                    box load.Category;
                                    box load.Position;
                                    null;
                                    box load.Load1Value;
                                    floatOptionToObj load.Load1DistanceFt;
                                    floatOptionToObj load.Load1DistanceIn;
                                    floatOptionToObj load.Load2Value;
                                    floatOptionToObj load.Load2DistanceFt;
                                    floatOptionToObj load.Load2DistanceIn;
                                    refObj;
                                    box(load.LoadCaseString)|]] |> array2D
                loadArray

            let allLoads =
                let joistLoads = 
                    if List.isEmpty joistsWithLC3Loads then
                        None 
                    else
                        joists
                        |> List.filter
                            (fun j ->
                                not (Seq.isEmpty (j.LC3Loads loads SDS seperatorInfo)))
                        |> List.map getJoistLoadArray
                        |> Array2D.joinMany (Array2D.joinByRows)
                        |> Some
                let girderLoads = 
                    if List.isEmpty girdersWithLC3Loads then
                        None
                    else
                        girders
                        |> List.filter
                            (fun g ->
                                not (Seq.isEmpty (g.Lc3and4Loads loads SDS liveLoadUNO liveLoadSpecialNotes inwardPressureUNO inwardPressureSpecialNotes seperatorInfo)))
                        |> List.map getGirderLoadArray
                        |> Array2D.joinMany (Array2D.joinByRows)
                        |> Some
                
                match joistLoads, girderLoads with
                    | Some joistLoads, Some girderLoads -> Array2D.joinByRows joistLoads girderLoads
                    | Some joistLoads, None -> joistLoads
                    | None, Some girderLoads -> girderLoads
                    | None, None -> [] |> array2D

            
            
            let rec divideArray maxRow array =
                let rows = Array2D.length1 array
                if rows <= maxRow then
                    [array]             
                else
                    let first = array.[(Array2D.base1 array)..(maxRow - ((Array2D.base1 array) + 1)), *]
                    let rest = array.[maxRow..rows - ((Array2D.base1 array) + 1), *]
                    first :: (divideArray maxRow rest)


            let loadPagesAsArray = divideArray 42 allLoads

            let addLoadSheet() =
                let workSheetNames = [for sheet in bom.Worksheets -> (sheet :?> Worksheet).Name] 
                let indexOfLastLoadSheet, lastLoadSheetNumber =
                    let lastLoadSheetNumber = workSheetNames
                                              |> List.filter (fun sheet -> sheet.Contains("L ("))
                                              |> List.map (fun sheet -> System.Int32.Parse(sheet.Split([|"(";")"|], StringSplitOptions.RemoveEmptyEntries).[1]))
                                              |> List.max
                    (bom.Worksheets.[sprintf "L (%i)" lastLoadSheetNumber] :?> Worksheet).Index, lastLoadSheetNumber
                             

                let blankLoadWorksheet = bom.Worksheets.["L_A"] :?> Worksheet
                blankLoadWorksheet.Visible <- Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetVisible
                blankLoadWorksheet.Copy(Type.Missing, bom.Worksheets.[indexOfLastLoadSheet])
                blankLoadWorksheet.Visible <- Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetHidden
                let newLoadSheet = (bom.Worksheets.["L_A (2)"]) :?> Worksheet
                newLoadSheet.Name <- "L (" + string(lastLoadSheetNumber + 1) + ")"
                newLoadSheet

            let numLoadPages = List.length loadPagesAsArray
            let mutable loadPageCounter = 0
            for loadPage in loadPagesAsArray do
                loadPageCounter <- loadPageCounter + 1
                let newLoadSheet = addLoadSheet()
                //newLoadSheet.Range("A14", "D55").Value2 <- loadPage.[*, Array2D.base2 loadPage..(Array2D.base2 loadPage) + 3]
                //newLoadSheet.Range("F14", "M55").Value2 <- loadPage.[*, (Array2D.base2 loadPage) + 5..(Array2D.base2 loadPage) + 12]
                let numRows = Array2D.length1 loadPage
                let finalRow = numRows - 1 + 14
                newLoadSheet.Range("A14", sprintf "M%i" finalRow).Value2 <- loadPage
                

        
        if seperatorInfo.SeperateSeismic then
            changeSmLoadsToLC3()
        addLC3LoadsToLoadNotes()
        addLC3Loads()

    let adjustSinglePitchJoists (bom : Workbook) (bomInfo : Joist list * Girder list * LoadNote list * float) : Unit =
        ()

let seperateSeismic bomPath seperatorInfo=
    getAllInfo bomPath getInfo [Modifiers.seperateSeismic seperatorInfo]