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

#load "..\src\SpreadsheetDoc\Internal\Common.fs"
#load "..\src\SpreadsheetDoc\Internal\Syntax.fs"
#load "..\src\SpreadsheetDoc\Internal\Stylesheet.fs"
#load "..\src\SpreadsheetDoc\Internal\Render.fs"
#load "..\src\SpreadsheetDoc\SpreadsheetDoc.fs"
open SpreadsheetDoc.Internal.Render
open SpreadsheetDoc.SpreadsheetDoc

let outputFile (name : string) : string = 
    Path.Combine(__SOURCE_DIRECTORY__, @"..\output", name)

let test01 () : unit = 
    let doc1 = 
        spreadsheet 
            [ sheet "Hello"
                [ row [cell <| intValue 1000; blankCell ; text "hello"]
                ; row [cell <| intValue 1001; blankCell ; text "world"]
                ] 

            ; sheet "World"
                [ row [ text "hello world" |> bold ]
                ; row [ cell <| dateTimeValue System.DateTime.Now ] // see stackoverflow 2792304 how-to-insert-a-date-into-an-open-xml-worksheet
                ]
            ]
    renderSpreadsheetDoc doc1 (outputFile "test01.xlsx")