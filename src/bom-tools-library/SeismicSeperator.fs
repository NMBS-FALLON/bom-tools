module DESign.BomTools.SeismicSeperator

open System
open System.IO
open Microsoft.Office.Interop.Excel
open System.Runtime.InteropServices
open DESign.ArrayExtensions
open System.Text.RegularExpressions
open Ookii.Dialogs.WinForms;



let (|Regex|_|) pattern input =
        let m = Regex.Match(input, pattern)
        if m.Success then Some(List.tail [ for g in m.Groups -> g.Value ])
        else None

type Result<'T,'TError> = 
    | Ok of ResultValue:'T 
    | Error of ErrorValue:'TError

let handleResultWithFailure result =
    match result with
    | Ok r -> r
    | Error s -> failwith s

let stringOptionFromObj (o: obj) =
    match o with
    | null -> None
    | _ -> Some (string o)        

let floatFromObj (o: obj) =
    match o with
    | null -> Error (sprintf "Expecting float, received %O" o)
    | :? string ->
        match string o with
        | "" -> Error (sprintf "Expecting float, received %O" o)
        | _ ->
            match System.Double.TryParse(string o) with
            | true, n -> Ok n
            | _ -> Error (sprintf "Expecting float, received %O" o)
    | :? float -> Ok (System.Convert.ToDouble o)
    | _ -> Error (sprintf "Expecting float, received %O" o)



let floatOptionFromObj (o: obj) =
    match o with
    | null -> Ok None
    | :? string ->
        match string o with
        | "" -> Ok None
        | _ ->
            match System.Double.TryParse(string o) with
            | true, n -> Ok (Some n)
            | _ -> Error (sprintf "Expecting float, received %O" o)
    | :? float -> Ok (Some (System.Convert.ToDouble o))
    | _ -> Error (sprintf "Expecting float, received %O" o)

let floatOptionToObj (fo: float option) =
    match fo with
    | Some v -> box v
    | None -> null


let parseObjToFloatOptionWithFailure (o: obj) =
    let result = floatOptionFromObj o
    handleResultWithFailure result



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
                let n = Convert.ToInt32 (a2D.[row- numPanels + j, startColIndex + 4])
                let feet = Convert.ToDouble (a2D.[row- numPanels + j, startColIndex + 5])
                let inch = Convert.ToDouble (a2D.[row- numPanels + j, startColIndex + 6])
                for k = 1 to Convert.ToInt32 (a2D.[row- numPanels + j, startColIndex + 4]) do
                    yield
                        {
                        Number = 1
                        LengthFt = Convert.ToDouble (a2D.[row- numPanels + j, startColIndex + 5])
                        LengthIn = Convert.ToDouble (a2D.[row- numPanels + j, startColIndex + 6])
                        }]
        if numPanels <> 0 then
            let mark = string (a2D.[row- numPanels + 1, startColIndex])
            let aFt = Convert.ToDouble (a2D.[row- numPanels + 1, startColIndex + 2])
            let aIn = Convert.ToDouble (a2D.[row- numPanels + 1, startColIndex + 3])
            let bFt = Convert.ToDouble (a2D.[row, startColIndex + 7])
            let bIn = Convert.ToDouble (a2D.[row, startColIndex + 8])
            yield
                 {
                 Mark = string (a2D.[row- numPanels + 1, startColIndex])
                 A_Ft = Convert.ToDouble (a2D.[row- numPanels + 1, startColIndex + 2])
                 A_In = Convert.ToDouble (a2D.[row- numPanels + 1, startColIndex + 3])
                 B_Ft = Convert.ToDouble (a2D.[row, startColIndex + 7])
                 B_In = Convert.ToDouble (a2D.[row, startColIndex + 8])
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



    member this.LC3Loads (loadNotes :LoadNote list) sds =
        
        match this.UDL, (this.Sds sds) with
        | Some udl, Some sds ->
            loadNotes
            |> List.filter (fun note -> this.LoadNoteList |> List.contains note.LoadNumber)
            |> List.map (fun note -> note.Load)
            |> List.filter (fun load -> load.Category <> "WL" && load.Category <> "SM" && (load.LoadCases = [] || (load.LoadCases |> List.contains 1)))
            |> List.map (fun load -> {load with LoadCases = [3]})
            |> List.append [udl; sds]
        | _ -> []


type _AdditionalJoist =
    {
    LocationFt : obj
    LocationIn : obj
    Load : float
    }

    member this.ToLoad() =            
        {
        Type = "C"
        Category = "CL"
        Position = "TC"
        Load1Value = this.Load * 1000.0
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
        Ref = None
        LoadCases = []
        }


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
    LiveLoadUNO : string
    LiveLoadSpecialNotes : Note List
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

    member this.UDL_PDL (liveLoadUNO : string) (liveLoadSpecialNotes : Note List)=
        let size = this.GirderSize
        let sizeAsArray = size.Split( [|"G"; "BG"; "VG"; "N"; "K"|], StringSplitOptions.RemoveEmptyEntries)
        let load = sizeAsArray.[2]
        let minSpace =
            let geometry = this.GirderGeometry
            let aSpace = geometry.A_Ft + geometry.A_In / 12.0
            let bSpace = geometry.B_Ft + geometry.B_In / 12.0
            let minPanelSpace = List.min (geometry.Panels |> List.map (fun geom -> geom.LengthFt + geom.LengthIn / 12.0))
            List.min [aSpace; bSpace; minPanelSpace]

        let TL = float load

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
        let UDL = DL / minSpace
        UDL, DL


    member this.SDS sds =
        let udl, _ = this.UDL_PDL this.LiveLoadUNO this.LiveLoadSpecialNotes
        let SDS = udl * 0.14 * sds
        Load.create("U", "SM", "TC", SDS,
                      None, None, None, None, None, None, [3])

    member this.DeadLoads =
        let _,dl = this.UDL_PDL this.LiveLoadUNO this.LiveLoadSpecialNotes
        let geom = this.GirderGeometry
        [for i = 1 to geom.NumPanels do
            let distanceFt, distanceIn = getPanelDim i geom
            yield
                Load.create("C", "CL", "TC", dl, Some distanceFt, Some distanceIn,
                             None, None, None, None, [3]) ]
        

    member this.LC3Loads (loadNotes :LoadNote list) (sds : float) =
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
                |> List.filter (fun load -> load.Category <> "WL" && load.Category <> "SM" && (load.LoadCases = [] || (load.LoadCases |> List.contains 1)))
                |> List.map (fun load -> {load with LoadCases = [3]})                 
            
            []
            |> List.append sds_From_UDLs_And_UTLs
            |> List.append [this.SDS sds]
            |> List.append this.DeadLoads
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
                            |> List.map (fun string -> System.Int32.Parse(string))
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
                                | Failure(msg) -> printfn "Issue with load line %i; %s" currentIndex msg; failwith ""
                                
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
                            {
                            LocationFt = string a.[col]
                            LocationIn = string a.[col + 1]
                            Load = Convert.ToDouble(a.[col + 2])
                            }
                        yield additionalJoist.ToLoad()
                col <- col + 4]

        let getAdditionalJoistsFromArray (a2D : obj [,]) =
            let mutable startRowIndex = Array2D.base1 a2D
            let colIndex = Array2D.base2 a2D
            let endIndex = (a2D |> Array2D.length1) - (if startRowIndex = 0 then 1 else 0)
            let additionalJoists : AdditionalJoist list =
                [for currentIndex = startRowIndex to endIndex do
                    if a2D.[currentIndex, colIndex] <> null && a2D.[currentIndex, colIndex] <> (box "") then
                        yield
                            {
                            Mark = string a2D.[currentIndex, colIndex]
                            AdditionalJoists = getAdditionalJoistsFromArraySlice a2D.[currentIndex, *]
                            } ]
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
                            LiveLoadUNO = ""
                            LiveLoadSpecialNotes = []
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

        let addLiveLoadInfoToGirders (girders: Girder list, liveLoadUNO : string, liveLoadSpecialNotes : Note list) =
            [for girder in girders do
                yield {girder with LiveLoadUNO = liveLoadUNO; LiveLoadSpecialNotes = liveLoadSpecialNotes}]
                
let saveWorkbook (title : string) (workbook : Workbook) =
        let title = title.Replace(".xlsm", " (IMPORT).xlsm")
        let title = title.Replace(".xlsx", " (IMPORT).xlsx")
        workbook.SaveAs(title)

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
        let stopWatch = System.Diagnostics.Stopwatch.StartNew()
        printfn "Opening Workbook"
        workbook <- workbooks.Open(tempReportPath)
        stopWatch.Stop()
        printfn "Workbook is opened (in %i seconds)" (stopWatch.Elapsed.Minutes * 60 + stopWatch.Elapsed.Seconds)
        stopWatch.Restart()
        printfn "Retrieving BOM information"
        let info = getInfoFunction workbook
        stopWatch.Stop()
        printfn "BOM information retrieved (in %i seconds)" stopWatch.Elapsed.Seconds
        tempExcelApp.EnableEvents <- false
        for modifyWorkbookFunction in modifyWorkbookFunctions do
            stopWatch.Restart()
            printfn "Applying Workbook Modification"
            modifyWorkbookFunction workbook info
            stopWatch.Stop()
            printfn "Workbook Modification complete (in %i seconds)" (stopWatch.Elapsed.Minutes * 60 + stopWatch.Elapsed.Seconds)

        tempExcelApp.EnableEvents <- true
        
        workbook |> saveWorkbook reportPath

        printfn "Finished processing %s." reportPath 
        printfn "Finished processing all files."
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
        | [] -> "L = 0.0K"
        | _ -> liveLoadNotes.[0].Text

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
            let girders = (CleanBomInfo.CleanGirders.addLiveLoadInfoToGirders (girders, liveLoadUNO, liveLoadSpecialNotes))
            girders

    (joists, girders, loads, SDS)


module Modifiers =
    let seperateSeismic (bom : Workbook) (bomInfo : Joist list * Girder list * LoadNote list * float) : Unit =

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
                if a2D.[currentIndex, startCol + 2] = (box "SM") && (lc = "1" || lc = "") then
                    a2D.[currentIndex, startCol + 12] <- box "3"

    

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
            let joists, girders, loads, SDS = bomInfo
            let joistsWithLC3Loads = joists |> List.filter (fun joist -> List.isEmpty (joist.LC3Loads loads SDS) = false)
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

            let girdersWithLC3Loads = girders |> List.filter (fun girder -> List.isEmpty (girder.LC3Loads loads SDS) = false)
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
                        if (sheet.Range("A26").Value2 :?> string) = "MARK" then
                            sheet.Range("Z28","Z45").Value2 <- array.[*,array.GetLength(1)..]              ///////////////////////////////////////////////////////////////////
                        else
                            sheet.Range("Z14", "Z45").Value2 <- array.[*,array.GetLength(1)..]          ////////////////////////////////////////////////////////////////////


        let addLC3Loads() =

            let joists, girders, loads, SDS = bomInfo

            let joistsWithLC3Loads = joists |> List.filter (fun joist -> List.isEmpty (joist.LC3Loads loads SDS) = false)
            let girdersWithLC3Loads = girders |> List.filter (fun girder -> List.isEmpty (girder.LC3Loads loads SDS) = false)

            let getJoistLoadArray (joist : Joist) = 
                let loadArray =
                    let mutable markAdded = false
                    [for load in (joist.LC3Loads loads SDS) do
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
                    [for load in (girder.LC3Loads loads SDS) do
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
                        Some ((joists |> List.map getJoistLoadArray) |> Array2D.joinMany (Array2D.joinByRows))
                let girderLoads = 
                    if List.isEmpty girdersWithLC3Loads then
                        None
                    else
                        Some ((girders |> List.map getGirderLoadArray) |> Array2D.joinMany (Array2D.joinByRows))
                
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
                printfn "Writing to Seismic Load Page %i of %i" loadPageCounter numLoadPages
                let newLoadSheet = addLoadSheet()
                //newLoadSheet.Range("A14", "D55").Value2 <- loadPage.[*, Array2D.base2 loadPage..(Array2D.base2 loadPage) + 3]
                //newLoadSheet.Range("F14", "M55").Value2 <- loadPage.[*, (Array2D.base2 loadPage) + 5..(Array2D.base2 loadPage) + 12]
                let numRows = Array2D.length1 loadPage
                let finalRow = numRows - 1 + 14
                newLoadSheet.Range("A14", sprintf "M%i" finalRow).Value2 <- loadPage
                

        
        changeSmLoadsToLC3()
        addLC3LoadsToLoadNotes()
        addLC3Loads()

    let adjustSinglePitchJoists (bom : Workbook) (bomInfo : Joist list * Girder list * LoadNote list * float) : Unit =
        ()

let SeperateSeismic bomPath =
    getAllInfo bomPath getInfo [Modifiers.seperateSeismic]


let seperateSeismicAndAdjustSinglePitches bomPath =
    getAllInfo bomPath getInfo [Modifiers.seperateSeismic; Modifiers.adjustSinglePitchJoists]

let adjustSinglePitchJoists bomPath =
    getAllInfo bomPath getInfo [Modifiers.adjustSinglePitchJoists]
