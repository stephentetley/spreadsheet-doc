// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace SpreadsheetDoc.Internal


module Common = 

    // Input is zero indexed
    let columnName (ix : int) : string = 
        let getLetter (i : int) : char = i + 65 |> char
        let rec work (i : int) ac cont = 
            let (d,r) = System.Math.DivRem(i, 26)
            let ch = getLetter r
            if d <= 0 then 
                cont (ch :: ac)
            else
                work (d-1) (ch :: ac) cont
        work ix [] (fun x -> System.String  (List.toArray x))

    // Input is zero indexed
    let cellName (row : int) (column : int) : string = 
        columnName(column) + string(row + 1)


    
    