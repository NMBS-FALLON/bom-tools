module DESign.Helpers

open System.Text.RegularExpressions


let (|KCS|K|LH|G|) (joistSize : string) =
    match joistSize with
    | size when size.ToUpper().Contains("KCS") -> KCS
    | size when size.ToUpper().Contains("K") -> K
    | size when size.ToUpper().Contains("LH") -> LH
    | size when size.ToUpper().Contains("G") -> G
    | _ -> failwith (sprintf "Unknown joist type: %s" joistSize)
    
let handleWithFailwith result =
    match result with
    | Ok v -> v
    | Error s -> failwith s


let handleWithPrint result =
    match result with
    | Ok v -> v
    | Error s -> printfn "Error: %s." s

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



let ConvertToCommaSeparatedString (value:seq<string>) =
    let rec convert (innerVal:List<string>) acc =
        match innerVal with
            | [] -> acc
            | hd::[] -> convert [] (acc + hd)
            | hd::tl -> convert tl (acc + hd + ",")
    convert (Seq.toList value) ""


type FtIn = 
    { 
        Ft : float
        In : float
    }
    member this.ToFeet =
        this.Ft + this.In / 12.0

    member this.ToInches =
        this.Ft * 12.0 + this.In

    static member Create (Ft: float, In: float) = {Ft = Ft; In = In}
