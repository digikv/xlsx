#INCLUDE "hbclass.ch"
REQUEST HB_GT_WIN_DEFAULT                                             // Console

REQUEST HB_CODEPAGE_UTF8EX

PROCEDURE MAIN()
LOCAL oExcel, oSheet1, oSheet2
LOCAL nBorder, nFont, nNumFmt, nNumFmt1, nFill1, nFill2
LOCAL nStyle, nStyle1, nStyle2
LOCAL oDrawing, cPicturePath
hb_cdpSelect( 'UTF8EX' )
setmode(25,80)
SET DATE FORMAT "dd.mm.yyyy"
oExcel := WorkBook():New("sample.xlsx")
nNumFmt := oExcel:NewFormat("#,##0.000")
nNumFmt1 := oExcel:NewFormat("#,##0.00")
oSheet1:= oExcel:WorkSheet("Test1")
oSheet1:paperSize := 9 // A4
oSheet1:lLandscape := .T.
oSheet1:leftMargin := 0.5
oSheet1:rightMargin := 0.5
oSheet1:topMargin := 0.5
oSheet1:bottomMargin := 0.5
oSheet1:RowDetail( 5, 49.5 )    // NEW method
oSheet2:= oExcel:WorkSheet("Proba")
nFill1 := oExcel:NewFillPattern( 1, "FFEEEEEE", "FFFFFFCC" )
nFill2 := oExcel:NewFillPattern( 1, "FFDDDDDD", "FFEEEEEE" )

// NewFont( cFont, nFontSize, lBold, lItalic, lUnderline, lStrike, cRGB )
nFont := oExcel:NewFont( "Tahoma", 16, .T., .F., .F., .F., "FFFF0000" )
nBorder := oExcel:NewBorder( 1, 1, 1, 1, 0 )

// NewStyle( nFont, nBorder, nFill, nVA, nHA, nNumFormat, nRotation, lWrap )
nStyle :=  oExcel:NewStyle( nFont, nBorder, nFill1, 2, 2 )
nStyle1 := oExcel:NewStyle( , , , , , nNumFmt )
nStyle2 := oExcel:NewStyle( , nBorder, , , , nNumFmt1 )

// AddHeader( cLeft, cCenter, cRight )
oSheet1:AddHeader( "", "Kosovo is Serbia", "" )
oSheet1:AddFooter( "Novus Ordo Mundi-BRICS", "", "&P of &N" ) 
oSheet1:Cell("A1", 3 )
oSheet1:Cell("B1", 2 )
oSheet1:Cell("C1", "=A1*B1" )
oSheet1:Cell("A3", .T. )
oSheet1:Cell("C3", "TEKST1", oExcel:NewStyle( 1, 1, nFill2, 2, 2 ) )
oSheet1:ColumnsWidth( 4, 4, 10 )
oSheet1:Cell("D3", DATE() )



// ColumnsWidth( fromCol, toCol, nWidth )
oSheet2:ColumnsWidth( 3, 3, 40 )
oSheet2:Cell("C2", "בשנה הבאה בקוסובו", nStyle )

// change hight of row 2 in sheet2
oSheet2:RowDetail( 2, 29.5 ) 

oSheet1:Cell("N3", 35.2, nStyle1 ) 
oSheet1:MergeCell("B7:F8")
oSheet1:Cell("B7", "UNSC RESOLUTION 1244", nStyle )
oSheet2:Cell("O12", 22 )
oSheet1:Cell("G1", 1 )
oSheet1:Cell("G2", 2 )
oSheet1:Cell("G3", 3 )
oSheet1:Cell("G4", -3.5 )
oSheet1:Cell("G5", 5 )
oSheet1:Cell("G6", 6 )
oSheet1:Cell("G7", -7 )
oSheet1:Cell("G8", 8 )
oSheet1:Cell("G9", 9 )
oSheet1:Cell("G10", 10, nStyle2 )

// Put JPG drawing 
oDrawing := oSheet1:Drawing("logo")
cPicturePath := DiskName() + hb_OSDriveSeparator() + hb_PS() + CurDir() + hb_PS()+"MSLogoMalo.jpg"
oDrawing:TwoCellAnchor( "I6", "K9", cPicturePath )

oExcel:Save()

Return