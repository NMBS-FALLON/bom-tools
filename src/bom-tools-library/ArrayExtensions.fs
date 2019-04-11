module DESign.ArrayExtensions

module Array2D =

// Some generic functions for a single array that are required for the merge function but may also be helpful in some other cases:
    
// arrayDimension = 0 for rows and 1 for columns,
// index represents the index of the row or column that you want 
let getSlice arrayDimension index (array: 'T[,]) = 
    match arrayDimension with
    | 0 -> array.[index..index, *] |> Seq.cast<'T> |> Seq.toArray
    | 1 -> array.[*,index..index] |> Seq.cast<'T> |> Seq.toArray
    | _ -> failwith "arrayDimension must be either 0 for rows and 1 for columns"
    // getSlice 0 1 read "get the second row"
    // getSlice 1 0 reads "get the first column"

// Converts an Array2D into a list of array's
// arrayDimension = 0 for rows and 1 for columns, 
let toArrayOfSlices arrayDimension (array : 'T[,]) : 'T[] list =
    let iteratorDimension =
        match arrayDimension with
        | 0 -> 1
        | 1 -> 0
        | _ -> failwith "arrayDimension must be either 0 for rows and 1 for columns"
    [for i = 0 to array.GetLength(arrayDimension) - 1 do
        yield array |> getSlice arrayDimension i ]

let joinByRows (a1: 'a[,]) (a2: 'a[,]) =
    let a1l1,a1l2,a2l1,a2l2 = (Array2D.length1 a1),(Array2D.length2 a1),(Array2D.length1 a2),(Array2D.length2 a2)
    if a1l2 <> a2l2 then failwith "arrays have different column sizes"
    let result = Array2D.zeroCreate (a1l1 + a2l1) a1l2
    let index =
        if typeof<'a> = typeof<obj> then 1 else 0
    Array2D.blit a1 (Array2D.base1 a1) (Array2D.base2 a1) result 0 0 a1l1 a1l2
    Array2D.blit a2 (Array2D.base1 a2) (Array2D.base2 a2) result a1l1 0 a2l1 a2l2
    result

let joinByCols (a1: 'a[,]) (a2: 'a[,]) =
    let a1l1,a1l2,a2l1,a2l2 = (Array2D.length1 a1),(Array2D.length2 a1),(Array2D.length1 a2),(Array2D.length2 a2)
    if a1l1 <> a2l1 then failwith "arrays have different row sizes"
    let result = Array2D.zeroCreate a1l1 (a1l2 + a2l2)
    Array2D.blit a1 (Array2D.base1 a1) (Array2D.base2 a1) result 0 0 a1l1 a1l2
    Array2D.blit a2 (Array2D.base1 a2) (Array2D.base2 a2) result 0 a1l2 a2l1 a2l2
    result

    // here joiner function must be Array2D.joinByRows or Array2D.joinByCols
let joinMany joiner (a: seq<'a[,]>)  = 
    Seq.fold joiner (Seq.head a) (Seq.tail a)
    (*
    let arrays = a |> Array.ofSeq
    if Array.length arrays = 0 then 
        failwith "no arrays"
    elif Array.length arrays = 1 then 
        Array.head arrays
    else
        let rec doJoin acc arrays = 
            if Array.length arrays = 0 then
                acc
            elif Array.length arrays = 1 then
                joiner acc (Array.head arrays)
            else
                let acc = joiner acc (Array.head arrays)
                doJoin acc (Array.tail arrays)
        doJoin <|| (Array.head arrays, Array.tail arrays)
        // or doJoin arrays.[0] arrays.[1..] 
            
        *)



