// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace SheetDoc.Internal


module Syntax = 

    open System


    // TODO - text data should allow runs so we can have fairly simple formatting.

    type Value = 
        | Blank
        | Int64Val of int64
        | DecimalVal of decimal
        | StringVal of string
        | DateTimeVal of DateTime

    type CellDoc = 
        { CellValue : Value
          StyleIndex : uint32 option 
        }

    type RowDoc = 
        { RowCells : CellDoc list }
        

    type SheetDoc = 
        { SheetName : string
          SheetRows : RowDoc list
        }

    type SpreadSheetDoc = 
        { Sheets : SheetDoc list 
        }