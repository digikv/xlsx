// A printable report: title, styled table header, number/date/percent formats,
// totals with formulas, page setup, header and footer.
// Build: hbmk2 xlsxreport.hbp
#INCLUDE "hbclass.ch"
REQUEST HB_GT_WIN_DEFAULT
REQUEST HB_CODEPAGE_UTF8EX

PROCEDURE MAIN()
LOCAL oExcel, oSheet, i, nRow, nLast
LOCAL nTitle, nHead, nText, nInt, nNum, nDate, nPct, nBool, nTotal, nTotalLbl
LOCAL nBorder, nFill
LOCAL aData := { ;
   { "Laptop 15.6",            4, 899.90,  HB_STOD( "20261103" ), 0.20, .T. }, ;
   { "Monitor 27 inch",        6, 279.50,  HB_STOD( "20261104" ), 0.20, .T. }, ;
   { "Keyboard & mouse set",  25,  39.99,  HB_STOD( "20261105" ), 0.20, .F. }, ;
   { "USB-C docking station",  8, 149.00,  HB_STOD( "20261107" ), 0.10, .T. }, ;
   { "=not a formula, just text", 1, 1.00, CTOD( "" ),            0.00, .F. } }

hb_cdpSelect( 'UTF8EX' )
SET DATE FORMAT "dd.mm.yyyy"

oExcel := WorkBook():New( "report.xlsx" )

// --- styles -----------------------------------------------------------------
// NewBorder( nLeft, nRight, nTop, nBottom, nDiagonal [, cColorLeft, ... ] )
//   1 thin, 2 medium, 3 thick, 4 double, 5 dashed, 6 dotted, 7 hair, ...
nBorder := oExcel:NewBorder( 1, 1, 1, 1 )
// NewFillPattern( nPattern, cForeground, cBackground ) - 1 = solid
nFill   := oExcel:NewFillPattern( 1, "FF305496", "FF305496" )

// NewStyle( nFont, nBorder, nFill, nVertAlign, nHorzAlign, nNumFormat, nRotation, lWrap )
//   align: 1 left/top, 2 center, 3 right/bottom
nTitle    := oExcel:NewStyle( oExcel:NewFont( "Calibri", 16, .T., , , , "FF1F3864" ), , , 2, 2 )
nHead     := oExcel:NewStyle( oExcel:NewFont( , 11, .T., , , , "FFFFFFFF" ), nBorder, nFill, 2, 2, , , .T. )
nText     := oExcel:NewStyle( , nBorder, , , , , , .T. )
nInt      := oExcel:NewStyle( , nBorder, , , 2, oExcel:NewFormat( "#,##0" ) )
nNum      := oExcel:NewStyle( , nBorder, , , , oExcel:NewFormat( "#,##0.00" ) )
nDate     := oExcel:NewStyle( , nBorder, , , 2, oExcel:NewFormat( "dd.mm.yyyy" ) )
nPct      := oExcel:NewStyle( , nBorder, , , 2, oExcel:NewFormat( "0%" ) )
nBool     := oExcel:NewStyle( , nBorder, , , 2 )
nTotalLbl := oExcel:NewStyle( oExcel:NewFont( , 11, .T. ), oExcel:NewBorder( 1, 1, 2, 4 ), , , 3 )
nTotal    := oExcel:NewStyle( oExcel:NewFont( , 11, .T. ), oExcel:NewBorder( 1, 1, 2, 4 ), , , , oExcel:NewFormat( "#,##0.00" ) )

// --- sheet and page setup ----------------------------------------------------
oSheet := oExcel:WorkSheet( "Invoice" )
oSheet:paperSize   := 9          // A4
oSheet:lLandscape  := .T.
oSheet:fitToPage   := .T.        // fit all columns on one page width
oSheet:leftMargin  := 0.5
oSheet:rightMargin := 0.5
oSheet:AddHeader( "My Company Ltd.", "Sales report", DToC( Date() ) )
oSheet:AddFooter( "", "", "Page &P of &N" )

// ColumnsWidth( nFromCol, nToCol, nWidth )
oSheet:ColumnsWidth( 1, 1, 30 )
oSheet:ColumnsWidth( 2, 2, 8 )
oSheet:ColumnsWidth( 3, 4, 12 )
oSheet:ColumnsWidth( 5, 7, 11 )
oSheet:ColumnsWidth( 8, 8, 14 )

// --- title (merged cells) ----------------------------------------------------
oSheet:MergeCell( "A1:H1" )
oSheet:Cell( "A1", "Sales report - November 2026", nTitle )
oSheet:RowDetail( 1, 30 )       // row height

// --- table header ---------------------------------------------------------------
oSheet:RowDetail( 3, 32 )
oSheet:Cell( "A3", "Item",           nHead )
oSheet:Cell( "B3", "Qty",            nHead )
oSheet:Cell( "C3", "Price",          nHead )
oSheet:Cell( "D3", "Date",           nHead )
oSheet:Cell( "E3", "VAT",            nHead )
oSheet:Cell( "F3", "In stock",       nHead )
oSheet:Cell( "G3", "Amount",         nHead )
oSheet:Cell( "H3", "Amount with VAT", nHead )

// --- rows -------------------------------------------------------------------
nRow := 3
FOR i := 1 TO LEN( aData )
   nRow++
   // text that starts with "=" is written as text; use Formula() for formulas
   oSheet:Cell( { nRow, 1 }, aData[ i, 1 ], nText )
   oSheet:Cell( { nRow, 2 }, aData[ i, 2 ], nInt )
   oSheet:Cell( { nRow, 3 }, aData[ i, 3 ], nNum )
   oSheet:Cell( { nRow, 4 }, aData[ i, 4 ], nDate )   // empty date -> empty cell
   oSheet:Cell( { nRow, 5 }, aData[ i, 5 ], nPct )
   oSheet:Cell( { nRow, 6 }, aData[ i, 6 ], nBool )
   oSheet:Formula( { nRow, 7 }, "B" + hb_ntos( nRow ) + "*C" + hb_ntos( nRow ), nNum )
   oSheet:Formula( { nRow, 8 }, "G" + hb_ntos( nRow ) + "*(1+E" + hb_ntos( nRow ) + ")", nNum )
NEXT
nLast := nRow

// --- totals -------------------------------------------------------------------
nRow++
oSheet:MergeCell( "A" + hb_ntos( nRow ) + ":F" + hb_ntos( nRow ) )
oSheet:Cell( { nRow, 1 }, "Total", nTotalLbl )
oSheet:Formula( { nRow, 7 }, "SUM(G4:G" + hb_ntos( nLast ) + ")", nTotal )
oSheet:Formula( { nRow, 8 }, "SUM(H4:H" + hb_ntos( nLast ) + ")", nTotal )

IF oExcel:Save()
   ? "report.xlsx created"
ELSE
   ? oExcel:cError        // e.g. the file is open in Excel
ENDIF

RETURN
