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
    let outfile = Path.Combine(outputDirectory(), "new_two_sheets.xlsx")
    let spreadsheetDocument : SpreadsheetDocument = 
        SpreadsheetDocument.Create(outfile, SpreadsheetDocumentType.Workbook)
    
    // Add a WorkbookPart to the document
    let workbookPart = spreadsheetDocument.AddWorkbookPart()
    workbookPart.Workbook <- new Workbook()


    // Add a WorksheetPart to the WorkbookPart
    let worksheetPart1 = workbookPart.AddNewPart<WorksheetPart>()
    let sheetData1 = new SheetData()
    let row1 = new Row(RowIndex = UInt32Value(1u))
    let cell1 = new Cell(DataType = EnumValue(CellValues.Number), CellReference = StringValue "A1")
    cell1.AppendChild (new CellValue(Text = "1001")) |> ignore
    row1.Append (cell1 :> OpenXmlElement)
    sheetData1.AppendChild (row1 :> OpenXmlElement) |> ignore

    let row2 = new Row(RowIndex = UInt32Value(2u))
    let cell2 = new Cell(DataType = EnumValue(CellValues.Number), CellReference = StringValue "A2")
    cell2.AppendChild (new CellValue(Text = "1002")) |> ignore
    row2.Append (cell2 :> OpenXmlElement)
    sheetData1.AppendChild (row2 :> OpenXmlElement) |> ignore


    worksheetPart1.Worksheet <- new Worksheet(sheetData1 :> OpenXmlElement)

    // Add a WorksheetPart to the WorkbookPart
    let worksheetPart2 = workbookPart.AddNewPart<WorksheetPart>()
    worksheetPart2.Worksheet <- new Worksheet(new SheetData() :> OpenXmlElement)


    let sheets : Sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets())

    let sheet1 : Sheet = new Sheet()
    sheet1.Id <- StringValue(spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart1))
    sheet1.SheetId <- UInt32Value(1u)
    sheet1.Name <- StringValue("Sheet_1")


    let sheet2 : Sheet = new Sheet()
    sheet2.Id <- StringValue(spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart2))
    sheet2.SheetId <- UInt32Value(2u)
    sheet2.Name <- StringValue("Sheet_2")

    sheets.Append([sheet1 :> OpenXmlElement; sheet2 :> OpenXmlElement])

    printfn "sheets.HasChildren %b" sheets.HasChildren


    workbookPart.Workbook.Save()
    spreadsheetDocument.Close()

