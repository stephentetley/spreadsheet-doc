#r "netstandard"
#r "System.Xml.Linq"
#r "System.Xml.ReaderWriter"
#r "System.Xml.XDocument"
#r "System.IO.FileSystem.Primitives"
open System.IO
open System.Linq
// open System.Text.RegularExpressions


#I @"C:\Users\stephen\.nuget\packages\system.io.packaging\4.5.0\lib\netstandard1.3"
#r "System.IO.Packaging"
#I @"C:\Users\stephen\.nuget\packages\DocumentFormat.OpenXml\2.9.1\lib\netstandard1.3"
#r "DocumentFormat.OpenXml"
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Spreadsheet
open DocumentFormat.OpenXml.Packaging



let outputDirectory () : string = 
    Path.Combine(__SOURCE_DIRECTORY__, @"..\output")

let createSpreadsheet (filepath:string) (sheetName:string) (sheetData:SheetData) =
    // Create a spreadsheet document by supplying the filepath.
    // By default, AutoSave = true, Editable = true, and Type = xlsx.
   
    using (SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook)) (fun spreadsheetDocument ->

    // Add a WorkbookPart to the document.
    let workbookPart = spreadsheetDocument.AddWorkbookPart(Workbook = new Workbook())

    // Add a WorksheetPart to the WorkbookPart.
    // http://stackoverflow.com/questions/5702939/unable-to-append-a-sheet-using-openxml-with-f-fsharp
    let worksheetPart = workbookPart.AddNewPart<WorksheetPart>()
    
    worksheetPart.Worksheet <- new Worksheet(sheetData:> OpenXmlElement)

    // Add Sheets to the Workbook.
    let sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets())

    // Append a new worksheet and associate it with the workbook.
    let sheet = new Sheet(  Id =  StringValue(spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart)),
                            SheetId =  UInt32Value(1u),
                            Name = StringValue(sheetName)
                            )
    [sheet :> OpenXmlElement] |> sheets.Append
    )

//helpers
let createCellReference (header:string) (index:int) =
    StringValue(header + string(index))

let createNumberCell number (header:string) (index:int) =
    let cell = new Cell(DataType = EnumValue(CellValues.Number), CellReference = createCellReference header index)
    let value = new CellValue(Text = number.ToString())
    value |> cell.AppendChild|> ignore
    cell :> OpenXmlElement

let createTextCell text (header:string) (index:int) =
    let cell = new Cell(DataType = EnumValue(CellValues.InlineString), CellReference = createCellReference header index)
    let inlineString = new InlineString()
    let t = new Text(Text = text)
    t |> inlineString.AppendChild |> ignore
    inlineString |> cell.AppendChild|> ignore
    cell :> OpenXmlElement

let createContentRow (text, (number1:int), (number2:int), (index:int)) =
    let row = new Row(RowIndex = UInt32Value(uint32(index)))
    let cell1 = createTextCell text "A" index
    let cell2 = createNumberCell number1 "B" index
    let cell3 = createNumberCell number2 "C" index 
    cell1 |> row.Append
    //cell2 |> row.Append
    //cell3 |> row.Append
    row :> OpenXmlElement

//test
let createTestSheetData = 
    let sheetData = new SheetData()
    ("test1", 123, 456, 1) |> createContentRow |> sheetData.AppendChild |> ignore
    //("test2", 35, 1231, 2) |> createContentRow |> sheetData.AppendChild |> ignore
    //("test3", 345, 21, 3) |> createContentRow |> sheetData.AppendChild |> ignore
    sheetData

let testData = createTestSheetData
let outpath = Path.Combine(outputDirectory(), "test.xlsx")
let result = createSpreadsheet outpath "test" testData;;