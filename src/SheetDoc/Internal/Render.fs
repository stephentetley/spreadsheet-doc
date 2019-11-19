// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace SheetDoc.Internal


module Render = 
    
    open DocumentFormat.OpenXml
    open DocumentFormat.OpenXml.Spreadsheet
    open DocumentFormat.OpenXml.Packaging

    open SheetDoc.Internal.Syntax


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
            List.mapi (fun i x -> 
                let worksheetPartId = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart)
                renderSheetDoc worksheetPartId (i+1) x) 
                spreadsheet.Sheets

        sheets.Append(sheetList)

        workbookPart.Workbook.Save()
        spreadsheetDocument.Close()