// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace SheetDoc

module SheetDoc =
    
    open System
    open SheetDoc.Internal.Syntax

    type ValueDoc = Value

    /// Design Principle
    /// The format of a spreadsheet is complicated, so we will build 
    /// a simple 'model' of a spreadsheet and render it.

    let intDoc (d : int) : ValueDoc = IntValue d

    let text (s : string) : ValueDoc = StrValue s

    let datetime (dt : DateTime) : ValueDoc = DateTimeValue dt

    let cell (value : ValueDoc) : CellDoc = { CellValue = value }

    let blank : CellDoc = { CellValue = Blank }

    let row (cells : CellDoc list) : RowDoc = { RowCells = cells }

    let sheet (name : string) (rows : RowDoc list) = 
        { SheetName = name
          SheetRows = rows
        }

    let spreadsheet (sheets : SheetDoc list) = 
        { Sheets = sheets
        }



