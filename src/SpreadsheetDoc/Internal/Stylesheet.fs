// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace SpreadsheetDoc.Internal


module Stylesheet = 
   
    open DocumentFormat.OpenXml
    open DocumentFormat.OpenXml.Spreadsheet


    let fontCollection () : Fonts = 
        let font0 = new Font()
        let font1 = new Font()
        font1.Bold <- new Bold()
        new Fonts (List.map (fun x -> x :> OpenXmlElement) [ font0; font1 ])

    /// TODO - a stylesheet is "very static" in Excel 
    /// It looks like the most appropriate strategy to have them
    /// would be a "collection pass" that finds all styles used in a 
    /// spreadsheet and generates a table accordingly.
    /// The alternative - implement a fixed set of styles - which we 
    /// currently use will result in a combinatorial explosion as we
    /// add more styles
    let sheetDocStylesheet () = 
        let stylesheet = new Stylesheet()
        let cellFormats : OpenXmlElement list = 
            let dateformat = new CellFormat()
            dateformat.NumberFormatId <- UInt32Value(14u)
            dateformat.ApplyNumberFormat <- BooleanValue(true)

            let boldformat = new CellFormat()
            boldformat.FontId <- UInt32Value(1u)

            [ new CellFormat() :> OpenXmlElement
            ; dateformat :> OpenXmlElement
            ; boldformat :> OpenXmlElement
            ]
        stylesheet.Fonts <- fontCollection ()
        stylesheet.Fills <- new Fills([new Fill() :> OpenXmlElement])
        stylesheet.Borders <- new Borders([new Border() :> OpenXmlElement])
        stylesheet.CellStyleFormats <- new CellStyleFormats([new CellFormat() :> OpenXmlElement])
        stylesheet.CellFormats <- new CellFormats(cellFormats)
        stylesheet