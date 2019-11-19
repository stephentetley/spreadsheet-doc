// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace SheetDoc.Internal


module Render = 
    
    open DocumentFormat.OpenXml
    open DocumentFormat.OpenXml.Spreadsheet
    open DocumentFormat.OpenXml.Packaging

    open SheetDoc.Internal.Common
    open SheetDoc.Internal.Syntax

    let renderTextCell (row : int) (col: int) (value : string) : OpenXmlElement = 
        let cellRef = cellName row col |> StringValue
        let cell = new Cell(DataType = EnumValue(CellValues.InlineString), CellReference = cellRef)
        let inlineStr = new InlineString()
        let text = new Text(Text = value)
        inlineStr.AppendChild(text :> OpenXmlElement) |> ignore
        cell.AppendChild(inlineStr) |> ignore
        cell :> OpenXmlElement

    let renderIntCell (row : int) (col: int) (value : int) : OpenXmlElement = 
        let cellRef = cellName row col |> StringValue
        let cell = new Cell(DataType = EnumValue(CellValues.Number)
                            , CellReference = cellRef)
        cell.CellValue <- new CellValue(value.ToString())
        cell :> OpenXmlElement

    let renderCell (row : int) (col: int) (value : Value) : OpenXmlElement = 
        match value with
        | Str s -> renderTextCell row col s
        | Int i -> renderIntCell row col i

    let renderRow (rowIx : int) (cellDocs : CellDoc list ) : OpenXmlElement = 
        let cells = cellDocs |> List.mapi (fun i x -> renderCell rowIx i x.CellValue) 
        let row = new Row(RowIndex = UInt32Value(uint32 <| rowIx + 1))
        cells |> List.iter (fun cell -> row.Append(cell))
        row :> OpenXmlElement

    let fillSheetData (sheetDoc : SheetDoc) : OpenXmlElement = 
        let sheetData = new SheetData()
        sheetDoc.SheetRows |> 
            List.iteri (fun i x -> 
                let row = renderRow i x.RowCells 
                sheetData.AppendChild(row) |> ignore)
        sheetData :> OpenXmlElement
        


    let renderSheetDoc (spreadsheetDocument : SpreadsheetDocument) 
                        (workbookPart : WorkbookPart) 
                        (sheets : Sheets) 
                        (ix : int) 
                        (sheetDoc : SheetDoc) : unit = 
        let worksheetPart = workbookPart.AddNewPart<WorksheetPart>() 
        let sheetData = fillSheetData sheetDoc

        // Add Sheetdata to the worksheet
        worksheetPart.Worksheet <- new Worksheet(sheetData)

        let partId = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart)
               
               
        let sheet = new Sheet()
        sheet.Id <- StringValue(partId)
        sheet.SheetId <- UInt32Value(uint32 <| ix + 1)
        sheet.Name <- StringValue(sheetDoc.SheetName)
        sheets.Append ([sheet :> OpenXmlElement])
        

    let renderSpreadSheetDoc (spreadsheet : SpreadSheetDoc) (outputPath : string) = 
        let spreadsheetDocument : SpreadsheetDocument = 
            SpreadsheetDocument.Create(outputPath, SpreadsheetDocumentType.Workbook)

        // Add a WorkbookPart to the document
        let workbookPart : WorkbookPart = spreadsheetDocument.AddWorkbookPart()
        workbookPart.Workbook <- new Workbook()
    
        let sheets : Sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets())
        

        spreadsheet.Sheets 
            |> List.iteri (renderSheetDoc spreadsheetDocument workbookPart sheets)


        workbookPart.Workbook.Save()
        spreadsheetDocument.Close()