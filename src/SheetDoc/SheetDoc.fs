// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace SheetDoc

module SheetDoc =
    
    open System

    open SheetDoc.Internal
    open SheetDoc.Internal.Syntax

    type ValueDoc = Value

    /// Design Principle
    /// The format of a spreadsheet is complicated, so we will build 
    /// a simple 'model' of a spreadsheet and render it.

    let int64Value (i : int64) : ValueDoc = Int64Val i

    let intValue (i : int) : ValueDoc = Int64Val (int64 i)

    let decimalValue (d : decimal) : ValueDoc = DecimalVal d

    let stringValue (s : string) : ValueDoc = StringVal s

    let dateTimeValue (dt : DateTime) : ValueDoc = DateTimeVal dt

    let cell (value : ValueDoc) : CellDoc = 
        { CellValue = value; StyleIndex = None }
    
    let blankCell : CellDoc = { CellValue = Blank; StyleIndex = None }

    /// Note - currently this will invalidate any existing formatting e.g. datetime
    let bold : CellDoc -> CellDoc = 
        fun cell -> { cell with StyleIndex = Some 2u }

    let text (s : string) : CellDoc = 
        match s with
        | null | "" -> blankCell
        | _ -> StringVal s |> cell


    

    let row (cells : CellDoc list) : RowDoc = { RowCells = cells }

    let sheet (name : string) (rows : RowDoc list) = 
        { SheetName = name
          SheetRows = rows
        }

    let spreadsheet (sheets : SheetDoc list) = 
        { Sheets = sheets
        }


    let writeSpreadSheetDoc (spreadSheetDoc: SpreadSheetDoc) (outputPath : string) = 
        Render.renderSpreadSheetDoc spreadSheetDoc outputPath


