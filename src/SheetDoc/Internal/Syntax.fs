// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace SheetDoc.Internal


module Syntax = 


    type SheetDoc = 
        { SheetName : string
        }

    type SpreadSheetDoc = 
        { Sheets : SheetDoc list 
        }