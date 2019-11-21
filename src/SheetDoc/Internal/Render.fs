// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace SheetDoc.Internal


module Render = 
    
    open System
    open System.Globalization

    open DocumentFormat.OpenXml
    open DocumentFormat.OpenXml.Spreadsheet
    open DocumentFormat.OpenXml.Packaging

    open SheetDoc.Internal.Common
    open SheetDoc.Internal.Syntax
    open SheetDoc.Internal.Stylesheet

    let renderTextCell (row : int) (col: int) (value : string) (styleIx : uint32 option) : OpenXmlElement = 
        let cellRef = cellName row col |> StringValue
        let cell = new Cell(DataType = EnumValue(CellValues.InlineString), CellReference = cellRef)
        let inlineStr = new InlineString()
        let text = new Text(Text = value)
        inlineStr.AppendChild(text :> OpenXmlElement) |> ignore
        cell.AppendChild(inlineStr) |> ignore
        match styleIx with
        | Some ix -> 
            cell.StyleIndex <- UInt32Value(ix)
        | _ -> ()
        cell :> OpenXmlElement

    let renderNumberCell (row : int) (col: int) 
                            (value : string) (styleIx : uint32 option) : OpenXmlElement = 
        let cellRef = cellName row col |> StringValue
        let cell = new Cell(DataType = EnumValue(CellValues.Number)
                            , CellReference = cellRef)
        cell.CellValue <- new CellValue(value.ToString())
        match styleIx with
        | Some ix -> 
            cell.StyleIndex <- UInt32Value(ix)
        | _ -> ()
        cell :> OpenXmlElement


    let renderDateTimeCell (row : int) (col: int) (value : DateTime) : OpenXmlElement = 
        let cellRef = cellName row col |> StringValue
        let cell = new Cell(DataType = EnumValue(CellValues.Number)
                            , CellReference = cellRef)
        cell.CellValue <- new CellValue(value.ToOADate().ToString(CultureInfo.InvariantCulture))
        cell.StyleIndex <- UInt32Value(1u)
        cell :> OpenXmlElement


    // Note - returning Some/None might not be right if we add styling
    // and we can have e.g a colour filled cell with no value.
    let renderCell (row : int) (col: int) (value : Value) (styleIx : uint32 option) : OpenXmlElement option = 
        match value with
        | Blank -> None
        | Int64Val i -> renderNumberCell row col (i.ToString()) styleIx |> Some
        | DecimalVal d -> renderNumberCell row col (d.ToString()) styleIx |> Some
        | StringVal s -> renderTextCell row col s styleIx |> Some
        | DateTimeVal dt -> renderDateTimeCell row col dt |> Some

    let renderRow (rowIx : int) (cellDocs : CellDoc list ) : OpenXmlElement = 
        let cells = cellDocs |> List.mapi (fun i x -> renderCell rowIx i x.CellValue x.StyleIndex) 
        let row = new Row(RowIndex = UInt32Value(uint32 <| rowIx + 1))
        cells |> List.choose id |> List.iter (fun cell -> row.Append(cell))
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

        let stylesPart : WorkbookStylesPart = 
            spreadsheetDocument.WorkbookPart.AddNewPart<WorkbookStylesPart>()

        let stylesheet = sheetDocStylesheet ()

        stylesPart.Stylesheet <- stylesheet

    
        let sheets : Sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets())
        

        spreadsheet.Sheets 
            |> List.iteri (renderSheetDoc spreadsheetDocument workbookPart sheets)


        workbookPart.Workbook.Save()
        spreadsheetDocument.Close()