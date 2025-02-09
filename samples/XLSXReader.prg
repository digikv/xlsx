#INCLUDE "hbclass.ch"
REQUEST HB_GT_WIN_DEFAULT
REQUEST HB_CODEPAGE_UTF8EX

PROCEDURE MAIN()
LOCAL oExcel, oSheet1, oSheet2, nFont
hb_cdpSelect( 'UTF8EX' )
setmode(25,80)
SET DATE FORMAT "dd.mm.yyyy"
oExcel := WorkBook():New()
oExcel:Read("sample.xlsx")
oSheet1:= oExcel:WorkSheet("Test1")
? "A1=", oSheet1:Cell("A1"), "B1=", oSheet1:Cell("B1"), "C1=", oSheet1:Cell("C1") 
? "A3=", oSheet1:Cell("A3"), "C3=", oSheet1:Cell("C3")
? "D3=", oSheet1:Cell("D3")
? "A5=", oSheet1:Cell("A5")
RETURN