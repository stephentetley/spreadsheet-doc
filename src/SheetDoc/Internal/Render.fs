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
        let cell = new Cell(DataType = EnumValue(CellValues.Number), CellReference = cellRef)
        let body = new Text(Text = value.ToString())
        cell.AppendChild(body) |> ignore
        cell :> OpenXmlElement

    let renderCell (row : int) (col: int) (value : Value) : OpenXmlElement = 
        match value with
        | Str s -> renderTextCell row col s
        | Int i -> renderIntCell row col i

    let renderRow (rowIx : int) (cellDocs : CellDoc list ) : OpenXmlElement = 
        let cells = cellDocs |> List.mapi (fun i x -> renderCell rowIx i x.CellValue) 
        let row = new Row(RowIndex = UInt32Value(uint32 <| rowIx + 1))
        row.Append(cells)
        row :> OpenXmlElement

    let fillSheetData (sheetData : SheetData) (sheetDoc : SheetDoc) : unit = 
        sheetDoc.SheetRows |> 
            List.iteri (fun i x -> 
                let row = renderRow i x.RowCells 
                sheetData.AppendChild(row) |> ignore)
        


    let renderSheetDoc (worksheetPartId : string) (ix : int) (sheetDoc : SheetDoc) : OpenXmlElement = 
        let sheet = new Sheet()
        sheet.Id <- StringValue(worksheetPartId)
        sheet.SheetId <- UInt32Value(uint32 ix)
        sheet.Name <- StringValue(sheetDoc.SheetName)
        sheet :> OpenXmlElement
        

    let renderSpreadSheetDoc (spreadsheet : SpreadSheetDoc) (outputPath : string) = 
        let spreadsheetDocument : SpreadsheetDocument = 
            SpreadsheetDocument.Create(outputPath, SpreadsheetDocumentType.Workbook)

        // Add a WorkbookPart to the document
        let workbookPart = spreadsheetDocument.AddWorkbookPart()
        workbookPart.Workbook <- new Workbook()
    
        // Add a WorksheetPart to the WorkbookPart
        let worksheetPart = workbookPart.AddNewPart<WorksheetPart>()
        worksheetPart.Worksheet <- new Worksheet(new SheetData() :> OpenXmlElement)

        let sheets : Sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets())

        let sheetList = 
            spreadsheet.Sheets
                |> List.mapi (fun i x -> 
                    let worksheetPartId = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart)
                    renderSheetDoc worksheetPartId (i+1) x) 
                

        sheets.Append(sheetList)


        workbookPart.Workbook.Save()
        spreadsheetDocument.Close()