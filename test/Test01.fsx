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

#load "..\src\SheetDoc\Internal\Common.fs"
#load "..\src\SheetDoc\Internal\Syntax.fs"
#load "..\src\SheetDoc\Internal\Render.fs"
#load "..\src\SheetDoc\SheetDoc.fs"
open SheetDoc.Internal.Syntax
open SheetDoc.Internal.Render
open SheetDoc.SheetDoc

let outputFile (name : string) : string = 
    Path.Combine(__SOURCE_DIRECTORY__, @"..\output", name)

let test01 () : unit = 
    let doc1 = 
        { Sheets = [ 
            { SheetName = "Sheet_1"
            ; SheetRows = 
                [ {RowCells = [ {CellValue = Int64Val 1000L}; {CellValue = StringVal "hello"} ] }
                ; {RowCells = [ {CellValue = Int64Val 1001L}; {CellValue = StringVal "world"} ] }
                ] 

            }
            { SheetName = "Sheet_2"
            ; SheetRows = 
                [ {RowCells = [ {CellValue = StringVal "world"} ] }
                ] 
            } 
        ]
        }
    renderSpreadSheetDoc doc1 (outputFile "test01.xlsx")


let test02 () : unit = 
    let doc1 = 
        spreadsheet 
            [ sheet "Hello"
                [ row [cell <| intValue 1000; blankCell ; text "hello"]
                ; row [cell <| intValue 1001; blankCell ; text "world"]
                ] 

            ; sheet "World"
                [ row [ text "hello world"]
                ; row [ cell <| dateTimeValue System.DateTime.Now ] // see stackoverflow 2792304 how-to-insert-a-date-into-an-open-xml-worksheet
                ]
            ]
    renderSpreadSheetDoc doc1 (outputFile "test02.xlsx")