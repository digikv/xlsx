// Reading an .xlsx file: sheet names, values, dates, formulas, ranges, merged cells.
// Build: hbmk2 xlsxreader.hbp   (run xlsxwriter first - it creates sample.xlsx)
#INCLUDE "hbclass.ch"
REQUEST HB_GT_WIN_DEFAULT
REQUEST HB_CODEPAGE_UTF8EX

PROCEDURE MAIN()
LOCAL oExcel, oSheet, cName, aRow

hb_cdpSelect( 'UTF8EX' )
setmode( 25, 80 )
SET DATE FORMAT "dd.mm.yyyy"

oExcel := WorkBook():New()
IF ! oExcel:Read( "sample.xlsx" )
   ? oExcel:cError
   RETURN
ENDIF

FOR EACH cName IN oExcel:SheetNames()
   oSheet := oExcel:GetSheet( cName )
   ? cName, "-", hb_ntos( oSheet:LastRow() ), "rows,", hb_ntos( oSheet:LastCol() ), "columns"
NEXT
?

oSheet := oExcel:GetSheet( "Test1" )
? "A1 =", oSheet:Cell( "A1" ), " B1 =", oSheet:Cell( "B1" )
? "C1 =", oSheet:Cell( "C1" ), " formula:", oSheet:CellFormula( "C1" )
? "A3 =", oSheet:Cell( "A3" ), " C3 =", oSheet:Cell( "C3" )
? "D3 =", oSheet:Cell( "D3" ), " type:", ValType( oSheet:Cell( "D3" ) )   // D = Date
? "B23 =", oSheet:Cell( "B23" )
? "A5 =", oSheet:Cell( "A5" )                                             // NIL - empty
? "merged:", hb_ValToExp( oSheet:MergedCells() )

? "G1:G10 ="
FOR EACH aRow IN oSheet:Range( "G1:G10" )
   ?? "", aRow[ 1 ]
NEXT

? "Proba C2 =", oExcel:GetSheet( 2 ):Cell( "C2" )

RETURN
