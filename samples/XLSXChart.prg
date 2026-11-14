// Charts: column ("bar"), horizontal bar ("hbar"), line and pie.
// Build: hbmk2 xlsxchart.hbp
#INCLUDE "hbclass.ch"
REQUEST HB_GT_WIN_DEFAULT
REQUEST HB_CODEPAGE_UTF8EX

PROCEDURE MAIN()
LOCAL oExcel, oSheet, oSheet2, oChart, i
LOCAL nHead, nNum, nTotal
LOCAL aMonths := { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" }
LOCAL aSales1 := { 1200, 1350, 1100, 1500, 1720, 1680, 1900, 2050, 1800, 1650, 1400, 2300 }
LOCAL aSales2 := { 1300, 1420, 1380, 1610, 1800, 1950, 2100, 2200, 1990, 1870, 1720, 2600 }
LOCAL aRegions := { "North", "South", "East", "West" }
LOCAL aShare := { 35, 25, 22, 18 }

hb_cdpSelect( 'UTF8EX' )
SET DATE FORMAT "dd.mm.yyyy"

oExcel := WorkBook():New( "charts.xlsx" )

nHead  := oExcel:NewStyle( oExcel:NewFont( , 11, .T. ), oExcel:NewBorder( 1, 1, 1, 1 ), ;
                           oExcel:NewFillPattern( 1, "FFD9E1F2", "FFD9E1F2" ), 2, 2 )
nNum   := oExcel:NewStyle( , oExcel:NewBorder( 1, 1, 1, 1 ), , , , oExcel:NewFormat( "#,##0" ) )
nTotal := oExcel:NewStyle( oExcel:NewFont( , 11, .T. ), oExcel:NewBorder( 1, 1, 2, 2 ), , , , oExcel:NewFormat( "#,##0" ) )

// --- data -------------------------------------------------------------------
oSheet := oExcel:WorkSheet( "Sales" )
oSheet:ColumnsWidth( 1, 1, 12 )
oSheet:ColumnsWidth( 2, 3, 10 )

oSheet:Cell( "A1", "Month", nHead )
oSheet:Cell( "B1", "2025", nHead )      // header text becomes the series name
oSheet:Cell( "C1", "2026", nHead )
FOR i := 1 TO 12
   oSheet:Cell( { i + 1, 1 }, aMonths[ i ] )          // address as { nRow, nCol }
   oSheet:Cell( { i + 1, 2 }, aSales1[ i ], nNum )
   oSheet:Cell( { i + 1, 3 }, aSales2[ i ], nNum )
NEXT
oSheet:Cell( "A14", "Total", nHead )
oSheet:Formula( "B14", "SUM(B2:B13)", nTotal )
oSheet:Formula( "C14", "SUM(C2:C13)", nTotal )

// --- charts -----------------------------------------------------------------
// AddChart( cTitle, cType, cCategories, xValues, cFromCell, cToCell, aSeriesNames )

// column chart with two series, names taken from B1 and C1
oSheet:AddChart( "Sales by month", "bar", "A2:A13", { "B2:B13", "C2:C13" }, "E1", "M18" )

// line chart
oSheet:AddChart( "Trend", "line", "A2:A13", { "B2:B13", "C2:C13" }, "E20", "M37" )

// horizontal bars, own series name and color, no legend
oChart := oSheet:AddChart( "2026 only", "hbar", "A2:A13", "C2:C13", "O1", "V18", { "Sales 2026" } )
oChart:aColors := { "70AD47" }
oChart:lLegend := .F.

// --- pie on a second sheet --------------------------------------------------
oSheet2 := oExcel:WorkSheet( "Regions" )
oSheet2:Cell( "A1", "Region", nHead )
oSheet2:Cell( "B1", "Share %", nHead )
FOR i := 1 TO LEN( aRegions )
   oSheet2:Cell( { i + 1, 1 }, aRegions[ i ] )
   oSheet2:Cell( { i + 1, 2 }, aShare[ i ], nNum )
NEXT
// pie shows only the first series; data can also come from another sheet: "'Sales'!A2:A13"
oSheet2:AddChart( "Market share", "pie", "A2:A5", "B2:B5", "D1", "J16" )

IF oExcel:Save()
   ? "charts.xlsx created"
ELSE
   ? oExcel:cError
ENDIF

RETURN
