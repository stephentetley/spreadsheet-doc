// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace SheetDoc.Internal


module Common = 

    // Input is zero indexed
    let columnName (ix : int) : string = 
        let getLetter (i : int) : char = i + 65 |> char
        let rec work (i : int) ac cont = 
            let (d,r) = System.Math.DivRem(i, 26)
            printfn "(%d, %d)" d r
            let ch = getLetter r
            if d <= 0 then 
                cont (ch :: ac)
            else
                printfn "%d" d
                work (d-1) (ch :: ac) cont
        work ix [] (fun x -> System.String  (List.toArray x))


    let cellName (row : int) (column : int) : string = 
        columnName(column) + string(row)


    
    