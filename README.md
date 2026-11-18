# xlsx

Create and read `.xlsx` files with a pure Harbour class. No Excel, OLE/ActiveX or external DLLs are needed.
Everything is in one source file: `XLSXClass.prg`.

## Features

- Numbers, text, dates, logical values and formulas
- Text is written as UTF-8; text in another codepage (e.g. from a DBF) is converted automatically
- Styles: fonts, colors, borders, fill patterns, alignment, text rotation, wrap text
- Built-in and custom number / date formats
- Merged cells, column widths, row heights
- Several worksheets in one workbook
- Page setup: paper size, orientation, margins, fit to page width, header and footer
- Pictures (PNG, JPG, GIF, BMP, EMF, WMF)
- Charts: column, horizontal bar, line and pie
- Reading existing `.xlsx` files: values, dates, formulas, ranges, merged cells and formatting
  (fonts, colors, borders, number formats, column widths, row heights, page setup),
  so a file can be read, changed and saved again

## Building

Add `XLSXClass.prg` to your project and link the `hbziparc` contrib (Harbour 3.2):

```
hbmk2 myapp.prg XLSXClass.prg hbziparc.hbc
```

The `samples` folder has ready-made `.hbp` files:

```
cd samples
hbmk2 xlsxwriter.hbp
hbmk2 xlsxreader.hbp
hbmk2 xlsxreport.hbp
hbmk2 xlsxchart.hbp
```

## Quick start

```harbour
#include "hbclass.ch"
REQUEST HB_CODEPAGE_UTF8EX

PROCEDURE Main()
   LOCAL oExcel, oSheet

   hb_cdpSelect( "UTF8EX" )

   oExcel := WorkBook():New( "hello.xlsx" )
   oSheet := oExcel:WorkSheet( "Sheet1" )

   oSheet:Cell( "A1", "Hello world" )      // text
   oSheet:Cell( "B1", 1234.56 )            // number
   oSheet:Cell( "C1", Date() )             // date (gets a date format automatically)
   oSheet:Cell( "C2", hb_DateTime() )      // date and time
   oSheet:Cell( "D1", .T. )                // logical
   oSheet:Formula( "E1", "B1*2" )          // formula

   IF ! oExcel:Save()
      ? oExcel:cError                      // e.g. the file is open in Excel
   ENDIF

   RETURN
```

## Examples

### Cell addresses

A cell can be addressed as `"B3"` or as `{ nRow, nCol }`. The array form is handy in loops:

```harbour
oSheet:Cell( "B3", "text" )
oSheet:Cell( { 3, 2 }, "same cell as B3" )

FOR i := 1 TO Len( aData )
   oSheet:Cell( { i + 1, 1 }, aData[ i, 1 ] )
   oSheet:Cell( { i + 1, 2 }, aData[ i, 2 ] )
NEXT
```

### Formulas

Use `Formula()`. The leading `=` is optional:

```harbour
oSheet:Formula( "C1", "SUM(A1:A10)" )
oSheet:Formula( { 2, 3 }, "=A1*1.2" )
```

Text that starts with `=` and is passed to `Cell()` is written as **text**, so a value from a database never becomes a formula by accident:

```harbour
oSheet:Cell( "C3", "=1+1" )        // the cell shows =1+1
```

For the old behaviour (text starting with `=` is a formula) set `oExcel:lFormulaPrefix := .T.`

### Styles

A style is built from a font, a border, a fill, alignment, a number format, rotation and wrap.
Each `New...()` method returns a number; identical definitions are created only once, so they can be called freely.

```harbour
// NewFont( cFont, nSize, lBold, lItalic, lUnderline, lStrike, cRGB )
// NewBorder( nLeft, nRight, nTop, nBottom, nDiagonal [, cColorLeft, cColorRight, cColorTop, cColorBottom, cColorDiagonal ] )
// NewFillPattern( nPattern, cForeground, cBackground )
// NewStyle( nFont, nBorder, nFill, nVertAlign, nHorzAlign, nNumFormat, nRotation, lWrap )

nBold  := oExcel:NewStyle( oExcel:NewFont( "Arial", 12, .T. ) )
nMoney := oExcel:NewStyle( , oExcel:NewBorder( 1, 1, 1, 1 ), , , 3, oExcel:NewFormat( "#,##0.00" ) )
nDate  := oExcel:NewStyle( , , , , 2, oExcel:NewFormat( "dd.mm.yyyy" ) )
nHead  := oExcel:NewStyle( oExcel:NewFont( , 11, .T., , , , "FFFFFFFF" ), ;   // white bold font
                           oExcel:NewBorder( 1, 1, 1, 1 ), ;                    // thin border
                           oExcel:NewFillPattern( 1, "FF305496", "FF305496" ), ;// solid blue
                           2, 2 )                                               // centered
nWrap  := oExcel:NewStyle( , , , 1, , , , .T. )       // top aligned, wrap text
nRot   := oExcel:NewStyle( , , , , 2, , 90 )          // text rotated 90 degrees

oSheet:Cell( "A1", "Title", nBold )
oSheet:Cell( "A2", 1500, nMoney )
oSheet:Cell( "A3", Date(), nDate )
oSheet:Cell( "A4", "Header", nHead )
oSheet:Cell( "A5", "A long text that wraps inside the cell", nWrap )
oSheet:Cell( "A6", "Up", nRot )
```

| Parameter | Values |
|---|---|
| alignment (`nVertAlign`, `nHorzAlign`) | 1 left / top, 2 center, 3 right / bottom |
| border style | 1 thin, 2 medium, 3 thick, 4 double, 5 dashed, 6 dotted, 7 hair, 8 mediumDashed, 9 dashDot, 10 mediumDashDot, 11 dashDotDot, 12 mediumDashDotDot, 13 slantDashDot |
| fill pattern | 1 solid, 2 darkHorizontal, 3 darkVertical, 4 darkDown, 5 darkUp, 6 darkGrid, 7 darkTrellis, 8 lightHorizontal, 9 lightVertical, 10 lightDown, 11 lightUp, 12 lightGrid, 13 lightTrellis, 14 mediumGray, 15 darkGray, 16 lightGray, 17 gray125, 18 gray0625 |
| colors | ARGB hex, e.g. `"FFFF0000"` (red) |

### Number and date formats

Common formats map to Excel's built-in ones (`"0"`, `"0.00"`, `"#,##0"`, `"#,##0.00"`, `"0%"`, `"0.00%"`, `"0.00E+00"`, `"h:mm"`, `"@"`, ...).
`"dd.mm.yyyy"`, `"dd/mm/yyyy"` and `"m/d/yyyy"` all map to the built-in short date, which Excel shows in the user's regional date format.
Any other string becomes a custom format:

```harbour
nPct    := oExcel:NewStyle( , , , , , oExcel:NewFormat( "0.00%" ) )
nRed    := oExcel:NewStyle( , , , , , oExcel:NewFormat( "#,##0.00;[Red]-#,##0.00" ) )
nCustom := oExcel:NewStyle( , , , , , oExcel:NewFormat( "dd.mm.yyyy hh:mm" ) )

oSheet:Cell( "D1", 0.256, nPct )        // 25.60%
oSheet:Cell( "D2", -1500, nRed )        // -1,500.00 in red
```

An empty date (`CToD( "" )`) is written as an empty cell.

### Column widths, row heights, merged cells

```harbour
oSheet:ColumnsWidth( 1, 1, 30 )       // column A
oSheet:ColumnsWidth( 2, 5, 12 )       // columns B..E
oSheet:ColumnsWidth( 8, 8, 10, .T. )  // hide column H
// RowDetail( nRow, nHeight, nStyle, lHide )
oSheet:RowDetail( 1, 28 )             // height of row 1
oSheet:RowDetail( 7, , , .T. )        // hide row 7

oSheet:MergeCell( "A10:E10" )         // call MergeCell() before Cell(),
oSheet:Cell( "A10", "Merged title", nHead )   // the style is applied to the whole range
```

### Page setup, header and footer

```harbour
oSheet:paperSize    := 9      // 9 = A4, 1 = Letter
oSheet:lLandscape   := .T.
oSheet:fitToPage    := .T.    // all columns on one page width
oSheet:leftMargin   := 0.5    // inches
oSheet:rightMargin  := 0.5
oSheet:topMargin    := 0.75
oSheet:bottomMargin := 0.75

// AddHeader( cLeft, cCenter, cRight ), AddFooter( cLeft, cCenter, cRight )
// &P page, &N number of pages, &D date, &T time, &F file name, &A sheet name
// (a single & in the text, e.g. "Procter & Gamble", is written correctly)
oSheet:AddHeader( "My Company", "&A", "&D" )
oSheet:AddFooter( "&F", "", "Page &P of &N" )

oSheet:nTitlesRows := 3       // rows 1..3 are repeated at the top of every printed page
```

### Several sheets

`WorkSheet( cName )` creates the sheet the first time and returns the same object later.
`GetSheet( cName )` returns an existing sheet or `NIL`.
Like in Excel, sheet names are not case sensitive, characters `: \ / ? * [ ]` are replaced with `_`
and a name is cut to 31 characters.

```harbour
oData    := oExcel:WorkSheet( "Data" )
oSummary := oExcel:WorkSheet( "Summary" )
oSummary:Formula( "A1", "SUM(Data!B2:B100)" )
```

### Pictures

```harbour
oDrawing := oSheet:Drawing( "logo" )
// TwoCellAnchor( cFromCell, cToCell, cPath ): the picture fills C3..E8
oDrawing:TwoCellAnchor( "C3", "E8", hb_cwd() + "logo.png" )

// OneCellAnchor( cCell, nWidth, nHeight, cPath ): fixed size in EMU (1 pixel = 9525, 1 cm = 360000)
oSheet:Drawing( "stamp" ):OneCellAnchor( "G3", 200 * 9525, 80 * 9525, hb_cwd() + "stamp.png" )
```

Cells can also be given as `{ nRow, nCol }`.

### Charts

```harbour
// AddChart( cTitle, cType, cCategories, xValues, cFromCell, cToCell, aSeriesNames )

oSheet:Cell( "A1", "Month" ) ; oSheet:Cell( "B1", "2025" ) ; oSheet:Cell( "C1", "2026" )
// ... months in A2:A13, values in B2:C13 ...

// column chart with two series; series names are read from B1 and C1
oSheet:AddChart( "Sales by month", "bar", "A2:A13", { "B2:B13", "C2:C13" }, "E1", "M18" )

// line chart
oSheet:AddChart( "Trend", "line", "A2:A13", { "B2:B13", "C2:C13" }, "E20", "M37" )

// horizontal bars with own series name and color, without legend
oChart := oSheet:AddChart( "2026", "hbar", "A2:A13", "C2:C13", "O1", "V18", { "Sales 2026" } )
oChart:aColors := { "70AD47" }
oChart:lLegend := .F.

// pie chart (only the first series), data from another sheet
oOther:AddChart( "Share", "pie", "'Regions'!A2:A5", "'Regions'!B2:B5", "D1", "J16" )
```

| Parameter | Meaning |
|---|---|
| `cTitle` | chart title, `""` for none |
| `cType` | `"bar"` (columns), `"hbar"` (horizontal bars), `"line"`, `"pie"` |
| `cCategories` | range with category labels: `"A2:A13"` or `"'Other sheet'!A2:A13"` |
| `xValues` | range with values, or an array of ranges for several series |
| `cFromCell`, `cToCell` | top-left and bottom-right cell of the chart (`"E1"` or `{ nRow, nCol }`); default `H2`..`P20` |
| `aSeriesNames` | optional; without it the text in the cell above (or left of) each range is used |

`AddChart()` returns the chart object: `aColors` (RGB colors of series / pie slices) and `lLegend` can be changed.

### Text in another codepage (DBF data)

Text that is not valid UTF-8 is converted from `oExcel:cCodePage`.
The default is the application codepage at the time `WorkBook():New()` is called, so in most programs nothing has to be done.
To set it explicitly:

```harbour
REQUEST HB_CODEPAGE_HRWIN

oExcel := WorkBook():New( "customers.xlsx" )
oExcel:cCodePage := "HRWIN"          // DBF data in Windows-1250
oSheet:Cell( { nRow, 1 }, customers->name )
```

Control characters (e.g. `Chr(0)` or `Chr(26)` left in DBF fields) are removed, because they are not allowed in XML.

### Reading a file

```harbour
oExcel := WorkBook():New()
IF ! oExcel:Read( "data.xlsx" )
   ? oExcel:cError                         // file not found or not an .xlsx
   RETURN
ENDIF

? hb_ValToExp( oExcel:SheetNames() )      // {"Sheet1", "Prices"}
oSheet := oExcel:GetSheet( 1 )            // by position, or by name: GetSheet( "Prices" )

? oSheet:Cell( "A1" )                     // number, text, logical, Date or DateTime; NIL if empty
? oSheet:Cell( { 3, 4 } )
? oSheet:LastRow(), oSheet:LastCol()      // last used row and column

// formula cells: Cell() returns the result saved in the file, CellFormula() the formula
? oSheet:Cell( "C10" ), oSheet:CellFormula( "C10" )    // 1250.5  =SUM(C2:C9)

// a range or the whole sheet as an array of rows
FOR EACH aRow IN oSheet:Range( "A2:D100" )
   ? aRow[ 1 ], aRow[ 2 ], aRow[ 3 ], aRow[ 4 ]
NEXT
aAll := oSheet:ToArray()                  // A1 .. LastRow() / LastCol()

? hb_ValToExp( oSheet:MergedCells() )     // {"A1:D1"}
```

- Cells with a date format are returned as `Date`, with date and time as `DateTime` (timestamp).
  A time-only format stays a number (fraction of a day). `oExcel:lReadDates := .F.` returns all dates as Excel serial numbers.
- `Range()` accepts `"A1:C10"`, `"A:C"` (all used rows), `"2:5"` (all used columns) or `{ nRow1, nCol1, nRow2, nCol2 }`.
- Text is returned in UTF-8; for another codepage use e.g. `hb_Translate( x, "UTF8", "HRWIN" )`.
- Rich text, inline strings, error values (`"#DIV/0!"`) and the 1904 date system (Excel for Mac) are handled.
- A formula without a saved result (e.g. a file created by this class) is returned by `Cell()` as text: `"=A1*B1"`.

### Formatting of a read file

`Read()` also reads the formatting: fonts, colors (also theme and indexed colors), fills, borders,
number formats, alignment, column widths and hidden columns, row heights and hidden rows, margins,
paper, orientation, header and footer. A file can be read, changed and saved with the same look:

```harbour
oExcel := WorkBook():New()
oExcel:Read( "invoice.xlsx" )
oSheet := oExcel:GetSheet( 1 )
oSheet:Cell( "B10", 1250 )              // new value, the cell keeps its format
oExcel:cName := "invoice_new.xlsx"      // save under another name
oExcel:Save()
```

The style of a cell can be read and set:

```harbour
nStyle := oSheet:CellStyle( "B3" )      // 0 = no style
h := oExcel:StyleInfo( nStyle )
? h[ "font" ], h[ "size" ], h[ "bold" ], h[ "color" ]   // Calibri  11  .T.  FFFF0000
? h[ "fill" ], h[ "numformat" ], h[ "halign" ]         // FFFFFF99  #,##0.00  center
? h[ "border" ][ "bottom" ], h[ "bordercolor" ][ "bottom" ]   // thin  FF000000

oSheet:CellStyle( "D3", nStyle )        // same style on another cell (also an empty one)
```

`StyleInfo()` keys: `font, size, bold, italic, underline, strike, color, fill, pattern, border, bordercolor,
numformat, halign, valign, rotation, wrap`. Colors are ARGB strings, `""` means none.

Not kept on save: pictures and charts of the read file, comments, conditional formatting,
data validation, frozen panes; a gradient fill becomes a solid fill.

## Samples

| File | What it shows |
|---|---|
| `samples/XLSXWriter.prg` | values of all types, styles, header and footer, merged cells, a picture, two sheets |
| `samples/XLSXReport.prg` | printable report: title, table header, number / date / percent formats, totals with formulas, page setup |
| `samples/XLSXChart.prg` | column, line, horizontal bar and pie charts |
| `samples/XLSXReader.prg` | reading `sample.xlsx`: sheet names, values, dates, formulas, ranges, merged cells |

## License

Dual licensed: **GNU GPL version 3 or later** (`COPYING`) **or a commercial license** from the
author, Srđan Dragojlović (S.O.K. doo Kraljevo, https://sokdoo.com, digikv@yahoo.com,
sok@sokdoo.com). Under the GPL a program that
includes the class must be distributed under the GPL with its source; for closed source programs
ask for a commercial license. See `LICENSE`.
