// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

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


#load "..\src\SheetDoc\SheetDoc.fs"
open SheetDoc.SheetDoc


let outputDirectory () : string = 
    Path.Combine(__SOURCE_DIRECTORY__, @"..\output")

let cellReference (column : string) (row : uint32) : string = 
    column + string(row)


let insertSharedStringItem (text : string) (shareStringPart : SharedStringTablePart) : int = 
    let ss : seq<SharedStringItem> = shareStringPart.SharedStringTable.Elements<SharedStringItem>()
    match Seq.tryFindIndex (fun (item: SharedStringItem) -> item.InnerText = text) ss with
    | Some i -> i
    | None -> 
        let xmltext = new Spreadsheet.Text(text) 
        shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(xmltext)) |> ignore
        shareStringPart.SharedStringTable.Save() |> ignore
        ss.Count()

let insertCellInWorkSheet (sheetData : SheetData) (columnId : string) (rowIx : uint32) = 


    let row : Row = 
        let rows = sheetData.Elements<Row>() 
        match Seq.tryFind (fun (r : Row) -> r.RowIndex = UInt32Value(rowIx)) rows with
        | Some rowx -> rowx
        | None ->
            let rowx = new Row()
            rowx.RowIndex <- UInt32Value(rowIx)
            sheetData.Append(rowx) 
            rowx

    let cells : seq<Cell> = row.Elements<Cell>()
    let cellRef = cellReference columnId rowIx
    match Seq.tryFind (fun (c : Cell) -> c.CellReference.Value = cellRef) cells with
    | Some cellx -> cellx
    | None -> 
        let newCell : Cell = new Cell()
        newCell.CellReference <- StringValue(cellRef)
        row.InsertAt(newCell, 0) |> ignore
        newCell



let test01 () = 
    let outfile = Path.Combine(outputDirectory(), "sample.xlsx")
    use spreadsheetDocument : SpreadsheetDocument = 
        SpreadsheetDocument.Create( outfile, SpreadsheetDocumentType.Workbook, autoSave=true)
    
    // Add a WorkbookPart to the document
    let workbookPart = spreadsheetDocument.AddWorkbookPart()
    workbookPart.Workbook <- new Workbook()

    let sharedStringPart : SharedStringTablePart = spreadsheetDocument.WorkbookPart.AddNewPart<SharedStringTablePart>()


    // Add a WorksheetPart to the WorkbookPart
    let worksheetPart = workbookPart.AddNewPart<WorksheetPart>()
    worksheetPart.Worksheet <- new Worksheet(new SheetData())

    let sheets : Sheets = workbookPart.Workbook.AppendChild<Sheets>(new Sheets())

    let sheet : Sheet = new Sheet()
    sheet.Id <- StringValue(workbookPart.GetIdOfPart(worksheetPart) )
    sheet.SheetId <- UInt32Value(1u)
    sheet.Name <- StringValue("Test Sheet")
    sheets.Append(sheet)


    let sheetData : SheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()

    match sheetData with
    | null -> 
        printfn "This won't work,..."
        workbookPart.Workbook.Save()
    | _ -> 
        let cell : Cell = new Cell()
        cell.CellReference <- StringValue("A1")
        cell.CellValue <- new CellValue(text = "1000")

        cell.DataType <- new EnumValue<CellValues> (CellValues.Number)
        workbookPart.Workbook.Save()
        // spreadsheetDocument.Close()








