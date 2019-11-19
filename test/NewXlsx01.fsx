// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

#r "netstandard"
#r "System.Xml.Linq"
#r "System.Xml.ReaderWriter"
#r "System.Xml.XDocument"
#r "System.IO.FileSystem.Primitives"
open System.IO


#I @"C:\Users\stephen\.nuget\packages\system.io.packaging\4.5.0\lib\netstandard1.3"
#r "System.IO.Packaging"
#I @"C:\Users\stephen\.nuget\packages\DocumentFormat.OpenXml\2.9.1\lib\netstandard1.3"
#r "DocumentFormat.OpenXml"
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Spreadsheet
open DocumentFormat.OpenXml.Packaging




let outputDirectory () : string = 
    Path.Combine(__SOURCE_DIRECTORY__, @"..\output")


let test01 () = 
    let outfile = Path.Combine(outputDirectory(), "sample.xlsx")
    let spreadsheetDocument : SpreadsheetDocument = 
        SpreadsheetDocument.Create(outfile, SpreadsheetDocumentType.Workbook)
    
    // Add a WorkbookPart to the document
    let workbookPart = spreadsheetDocument.AddWorkbookPart()
    workbookPart.Workbook <- new Workbook()


    // Add a WorksheetPart to the WorkbookPart
    let worksheetPart = workbookPart.AddNewPart<WorksheetPart>()
    worksheetPart.Worksheet <- new Worksheet(new SheetData() :> OpenXmlElement)

    let sheets : Sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets())

    let sheet : Sheet = new Sheet()
    sheet.Id <- StringValue(spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart))
    sheet.SheetId <- UInt32Value(1u)
    sheet.Name <- StringValue("Sheet_1")

    sheets.Append([sheet :> OpenXmlElement])

    printfn "sheets.HasChildren %b" sheets.HasChildren


    workbookPart.Workbook.Save()
    spreadsheetDocument.Close()

