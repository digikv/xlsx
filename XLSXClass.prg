/*
 *
 * XLSX writer/reader Class - Pure Harbour 
 *
 * Copyright 2018 - 2026 Srđan Dragojlović <digikv@yahoo.com>
 * Company  S.O.K. doo Kraljevo
 * www - https://sokdoo.com
 *
 * Dual licensed: GNU GPL version 3 or later (COPYING), or a commercial
 * license from the author - see LICENSE.
*/

#require "hbmzip"
#include "simpleio.ch"
#INCLUDE "hbclass.ch"


REQUEST HB_CODEPAGE_UTF8EX

STATIC aStdNumFormats := { "0" => -1, "0.00" => -2, "#,##0" => -3, "#,##0.00" => -4, "0%" => -9,  "0.00%" => -10, "0.00E+00" => -11, "# ?/?" => -12, "# ??/??" => -13, "m/d/yyyy" => -14, "dd.mm.yyyy" => -14, "dd/mm/yyyy" => -14, "d-mmm-yy" => -15, "d-mmm" => -16, "mmm-yy" => -17, "h:mm AM/PM" => -18, "h:mm:ss AM/PM" => -19, "h:mm" => -20, "h:mm:ss" => -21, "m/d/yyyy h:mm" => -22, "#,##0 ;(#,##0)" => -37, "#,##0 ;Red" => -38, "#,##0.00 ;(#,##0.00)" => -39, "#,##0.00 ;Red" => -40, "#,##0 ;[Red](#,##0)" => -38, "#,##0.00;(#,##0.00)" => -39, "#,##0.00;[Red](#,##0.00)" => -40, "mm:ss" => -45, "[h]:mm:ss" => -46, "mm:ss.0" => -47, "##0.00E+0" => -48, "@"  => -49 }
STATIC aHorizontal := {"left","center","right"}
STATIC aVertical := {"top","center","bottom"}
STATIC aBordersStyle := { "thin", "medium", "thick", "double", "dashed", "dotted", "hair", "mediumDashed", "dashDot", "mediumDashDot", "dashDotDot", "mediumDashDotDot", "slantDashDot" }
// sare za NewFillPattern( n ): ranije imena iz Excel 2003 XML formata (horzStripe...)
// koja xlsx ne poznaje, pa je radila samo 1 (solid); redosled 1-13 je isti kao ranije
STATIC aPattern := { "solid", "darkHorizontal", "darkVertical", "darkDown", "darkUp", "darkGrid", "darkTrellis", "lightHorizontal", "lightVertical", "lightDown", "lightUp", "lightGrid", "lightTrellis", "mediumGray", "darkGray", "lightGray", "gray125", "gray0625" }
STATIC aWorkbookDrawings, aPictureType

CLASS WorkBook
DATA cTempDir PROTECTED
DATA cName
DATA cFilePath
DATA Company INIT "S.O.K. doo Kraljevo, Serbia"
DATA Application INIT "XLSX Class"
DATA DocSecurity INIT 0
DATA ScaleCrop INIT .F.
DATA LinksUpToDate INIT .F.
DATA SharedDoc INIT .F.
DATA HyperlinksChanged INIT .F.
DATA AppVersion INIT "12.1000"
DATA aWorkSheetNames PROTECTED  
DATA aWorkSheetObjects PROTECTED 
DATA aWorkSheetDrawingObjects PROTECTED 
DATA aWorkbookFonts PROTECTED
DATA aWorkbookStyles PROTECTED
DATA aWorkbookFills PROTECTED
DATA aWorkbookBorders PROTECTED
DATA aUsedDrawingType 
DATA aNumFormat PROTECTED
DATA nDrawings PROTECTED
DATA nCharts INIT 0 PROTECTED
DATA aCondStyles INIT {} PROTECTED     // NewCondStyle(): { cFont, cFill, cBorder } (styles.xml <dxfs>)
DATA aSharedStrings INIT {}
DATA _str_total INIT 0 
DATA nTitlesRows INIT 0
DATA hSharedStrings                // tekst -> indeks u aSharedStrings (brza pretraga)
DATA cCodePage                     // kodna strana teksta koji nije UTF-8 (podraz. hb_cdpSelect())
DATA lFormulaPrefix INIT .F.       // .T. = staro ponasanje: Cell() sa "=..." je formula
DATA cError INIT ""                // opis greske posle neuspelog Save()
DATA lReadDates INIT .T.          // Read(): celija sa formatom datuma -> Date (datum+vreme -> DateTime)
DATA lDate1904 INIT .F.           // Read(): fajl koristi 1904 sistem datuma (Excel za Mac)
METHOD NewFont( cFont, nFontSize, lBold, lItalic, lUnderline, lStrike, cRGB )
METHOD NewFillPattern( nFillPattern, cFG, cBG )
METHOD NewBorder( nTL, nTR, nTT, nTB, nTD, cCL, cCR, cCT, cCB, cCD  )
METHOD NewStyle( nFont, nBorder, nFill, nVA, nHA, nFormat, nRotation, lWrap )
METHOD NewFormat( cFormat, lExact )
METHOD NewCondStyle( cFontColor, lBold, cFillColor, lItalic, lUnderline, cBorderColor )
METHOD New(cName)
METHOD WorkSheet(cName)
METHOD AddDrawing()
METHOD AddChartId()
METHOD GetSheet( cName )
METHOD GetTempDir()
METHOD Save()
METHOD Read(cName)
METHOD ToUTF8( c )
METHOD SheetNames()
METHOD StyleInfo( nStyle )
ENDCLASS

// Ucitava postojeci .xlsx. Vraca .T. ako je uspelo, inace .F. i opis greske u ::cError.
// Sheet-ovi: GetSheet( cIme ) ili GetSheet( nRedniBroj ), imena: SheetNames().
// Sheet se trazi preko xl/_rels/workbook.xml.rels (ranije po sheetId, sto ne vazi
// za fajlove u kojima su sheet-ovi brisani ili premestani).
METHOD Read( cName ) CLASS WorkBook
LOCAL cXL, cTok, cA, cTag, nKind, cText, cCur := "", lPh := .F., h, cTarget, cFile
LOCAL aSST := {}, aKinds, aXfMap, hRels := { => }, aSheets := {}
   ::cError := ""
   IF ! HB_ISSTRING( cName ) .OR. ! hb_FileExists( cName )
      ::cError := "Fajl ne postoji: " + hb_CStr( cName )
      RETURN .F.
   ENDIF
   // isto kao New(): putanja i ime posebno (ranije Read() pa Save() je spajao
   // tekuci folder sa punom putanjom)
   ::cFilePath := iif( FilePath(cName)=="", DiskName() + hb_OSDriveSeparator() + hb_PS() + CurDir() + hb_PS() , FilePath(cName) )
   ::cName := iif( AT(::cFilePath,cName)>0, SUBSTR(cName,LEN(::cFilePath)+1), cName )
   IF ! HB_ISHASH( ::hSharedStrings )
      ::hSharedStrings := { => }
   ENDIF
   hb_DirRemoveAll( ::cTempDir )
   MakeDir( ::cTempDir )
   BEGIN SEQUENCE WITH {| e | Break( e ) }
      my_UnzipFile( cName, ::cTempDir + hb_PS() )
   END SEQUENCE
   cXL := ::cTempDir + hb_PS() + "xl" + hb_PS()
   IF ! hb_FileExists( cXL + "workbook.xml" )
      ::cError := "Nije ispravan .xlsx fajl: " + cName
      hb_DirRemoveAll( ::cTempDir )
      RETURN .F.
   ENDIF

   // sheet-ovi redom i sistem datuma
   FOR EACH cTok IN _XTokens( cXL + "workbook.xml" )
      cA := _XTok( cTok, @cTag, @nKind )
      IF cTag == "workbookPr" .AND. nKind != 2
         cText := Lower( _HAttr( _XAttrs( cA ), "date1904" ) )
         ::lDate1904 := ( cText == "1" .OR. cText == "true" )
      ELSEIF cTag == "sheet" .AND. nKind != 2
         AAdd( aSheets, _XAttrs( cA ) )
      ENDIF
   NEXT

   // deljeni stringovi; rich text (vise <r><t>) se spaja, fonetski <rPh> se preskace
   FOR EACH cTok IN _XTokens( cXL + "sharedStrings.xml" )
      _XTok( cTok, @cTag, @nKind, @cText )
      DO CASE
      CASE cTag == "si"
         IF nKind == 1
            cCur := ""
         ELSEIF nKind > 1
            cText := IIF( nKind == 3, "", _XDecode( cCur ) )
            AAdd( aSST, cText )
            IF ! hb_HHasKey( ::hSharedStrings, cText )
               AAdd( ::aSharedStrings, cText )
               ::hSharedStrings[ cText ] := LEN( ::aSharedStrings )
            ENDIF
            ::_str_total++
         ENDIF
      CASE cTag == "rPh"
         lPh := ( nKind == 1 )
      CASE cTag == "t" .AND. nKind == 1 .AND. ! lPh
         cCur += _XEnt( cText )
      ENDCASE
   NEXT

   // stilovi: za svaki stil iz fajla (cellXfs) odgovarajuci stil ovog workbook-a
   // (font, boje, okviri, format, poravnanje) i da li je format datum/vreme
   aKinds := _ReadStyles( Self, cXL + "styles.xml", cXL + "theme" + hb_PS() + "theme1.xml" )
   aXfMap := aKinds[ 2 ]
   aKinds := aKinds[ 1 ]

   // rId -> putanja sheet-a
   FOR EACH cTok IN _XTokens( cXL + "_rels" + hb_PS() + "workbook.xml.rels" )
      cA := _XTok( cTok, @cTag, @nKind )
      IF cTag == "Relationship" .AND. nKind != 2
         h := _XAttrs( cA )
         hRels[ _HAttr( h, "Id" ) ] := _HAttr( h, "Target" )
      ENDIF
   NEXT

   FOR EACH h IN aSheets
      cTarget := hb_HGetDef( hRels, _HAttr( h, "id" ), "" )
      IF Empty( cTarget )
         cTarget := "worksheets/sheet" + _HAttr( h, "sheetId" ) + ".xml"
      ENDIF
      IF Left( cTarget, 1 ) == "/"
         cFile := ::cTempDir + hb_DirSepToOS( cTarget )
      ELSE
         cFile := cXL + hb_DirSepToOS( cTarget )
      ENDIF
      ::WorkSheet( _HAttr( h, "name" ) ):ReadWorkSheet( cFile, aSST, aKinds, aXfMap )
   NEXT
   hb_DirRemoveAll( ::cTempDir )
RETURN .T.

// Tekst koji nije ispravan UTF-8 smatra se tekstom u kodnoj strani ::cCodePage
// (podrazumevano kodna strana aplikacije u trenutku New()) i konvertuje se.
// ASCII i vec UTF-8 tekst (npr. cirilicni literali u izvornom kodu) prolaze
// neizmenjeni, pa pozivalac ne mora sam da konvertuje podatke iz baze.
METHOD ToUTF8( c ) CLASS WorkBook
IF HB_ISSTRING( c ) .AND. HB_ISSTRING( ::cCodePage ) .AND. ! Empty( ::cCodePage ) .AND. ;
   !( "UTF8" $ Upper( ::cCodePage ) ) .AND. ! hb_StrIsUTF8( c )
	c := hb_StrToUTF8( c, ::cCodePage )
ENDIF
RETURN c

METHOD GetTempDir() CLASS WorkBook
RETURN ::cTempDir

METHOD AddDrawing() CLASS WorkBook
::nDrawings := ::nDrawings +1
RETURN ::nDrawings

// redni broj grafikona u workbook-u (xl/charts/chart<N>.xml)
METHOD AddChartId() CLASS WorkBook
::nCharts := ::nCharts + 1
RETURN ::nCharts

// postojeci sheet po imenu ili rednom broju (1, 2, ...), inace NIL
// (za razliku od WorkSheet() ne pravi novi)
METHOD GetSheet( cName ) CLASS WorkBook
LOCAL n
IF HB_ISNUMERIC( cName )
   RETURN IIF( cName >= 1 .AND. cName <= LEN( ::aWorkSheetObjects ), ::aWorkSheetObjects[ cName ], NIL )
ENDIF
n := _SheetPos( ::aWorkSheetNames, _SheetName( cName ) )
RETURN IIF( n > 0, ::aWorkSheetObjects[ n ], NIL )

// imena svih sheet-ova, redom
METHOD SheetNames() CLASS WorkBook
RETURN AClone( ::aWorkSheetNames )

// Opis stila (broj iz NewStyle() ili oSheet:CellStyle()) kao hash:
// font, size, bold, italic, underline, strike, color, fill, pattern, border
// { left, right, top, bottom }, bordercolor { ... }, numformat, halign, valign,
// rotation, wrap. Boje su ARGB ("FFFF0000"), "" = nema.
METHOD StyleInfo( nStyle ) CLASS WorkBook
LOCAL h := { => }, s, f, b, p, nId, aSide := { "left", "right", "top", "bottom" }, i
s := IIF( HB_ISNUMERIC( nStyle ) .AND. nStyle >= 1 .AND. nStyle <= LEN( ::aWorkbookStyles ), ;
          hb_ATokens( ::aWorkbookStyles[ nStyle ], "," ), { "0", "0", "0", "0", "0", "0", "0", "0" } )
f := hb_ATokens( ::aWorkbookFonts[ Val( s[1] ) + 1 ], "," )
h[ "font" ]      := f[1]
h[ "size" ]      := Val( f[2] )
h[ "bold" ]      := ( f[3] == "1" )
h[ "italic" ]    := ( f[4] == "1" )
h[ "underline" ] := ( f[5] == "1" )
h[ "strike" ]    := ( f[6] == "1" )
h[ "color" ]     := _ARGB( f[7] )
p := hb_ATokens( ::aWorkbookFills[ Val( s[3] ) + 1 ], "," )
h[ "pattern" ]   := IIF( Val( s[3] ) >= 2, Val( p[1] ), 0 )
h[ "fill" ]      := IIF( h[ "pattern" ] > 0, _ARGB( p[2] ), "" )
b := hb_ATokens( ::aWorkbookBorders[ Val( s[2] ) + 1 ], "," )
h[ "border" ]      := { => }
h[ "bordercolor" ] := { => }
FOR i := 1 TO 4
   h[ "border" ][ aSide[ i ] ]      := IIF( LEN( b ) >= i * 2 .AND. Val( b[ i * 2 - 1 ] ) > 0, aBordersStyle[ Val( b[ i * 2 - 1 ] ) ], "" )
   h[ "bordercolor" ][ aSide[ i ] ] := IIF( Empty( h[ "border" ][ aSide[ i ] ] ), "", _ARGB( b[ i * 2 ] ) )
NEXT
nId := Val( s[6] )
h[ "numformat" ] := IIF( nId >= 164 .AND. nId - 163 <= LEN( ::aNumFormat ), ::aNumFormat[ nId - 163 ], _BuiltinFmtCode( nId ) )
h[ "halign" ]    := IIF( Val( s[4] ) > 0, aHorizontal[ Val( s[4] ) ], "" )
h[ "valign" ]    := IIF( Val( s[5] ) > 0, aVertical[ Val( s[5] ) ], "" )
h[ "rotation" ]  := Val( s[7] )
h[ "wrap" ]      := ( s[8] == "1" )
RETURN h

METHOD NewFillPattern( nFillPattern, cFG, cBG ) CLASS WorkBook
LOCAL nPos := 0, c
IF HB_ISNUMERIC(nFillPattern) .AND. nFillPattern>0 .AND. nFillPattern<=LEN(aPattern) .AND. HB_ISSTRING( cBG ) .AND. HB_ISSTRING( cFG )
	c := HB_NTOC(nFillPattern,0)+","+cFG+","+cBG
	IF (nPos:=_AScanX( ::aWorkbookFills, c ) )==0
		AADD( ::aWorkbookFills, c )
		nPos := LEN( ::aWorkbookFills )
	ENDIF
ENDIF
Return nPos

METHOD NewBorder( nTL, nTR, nTT, nTB, nTD, cCL, cCR, cCT, cCB, cCD  ) CLASS WorkBook
LOCAL nPos
LOCAL c := ""
c += IIF( (HB_ISNUMERIC(nTL) .AND. nTL>0 .AND. nTL<14), HB_NTOC(nTL,0)+IIF( (!HB_ISNIL(cCL) .AND. HB_ISSTRING( cCL )), ","+cCL+",", ",000000," ),"0,000000," )
c += IIF( (HB_ISNUMERIC(nTR) .AND. nTR>0 .AND. nTR<14), HB_NTOC(nTR,0)+IIF( (!HB_ISNIL(cCR) .AND. HB_ISSTRING( cCR )), ","+cCR+",", ",000000," ),"0,000000," )
c += IIF( (HB_ISNUMERIC(nTT) .AND. nTT>0 .AND. nTT<14), HB_NTOC(nTT,0)+IIF( (!HB_ISNIL(cCT) .AND. HB_ISSTRING( cCT )), ","+cCT+",", ",000000," ),"0,000000," )
c += IIF( (HB_ISNUMERIC(nTB) .AND. nTB>0 .AND. nTB<14), HB_NTOC(nTB,0)+IIF( (!HB_ISNIL(cCB) .AND. HB_ISSTRING( cCB )), ","+cCB+",", ",000000," ),"0,000000," )
c += IIF( (HB_ISNUMERIC(nTD) .AND. nTD>0 .AND. nTD<14), HB_NTOC(nTD,0)+IIF( (!HB_ISNIL(cCD) .AND. HB_ISSTRING( cCD )), ","+cCD+",", ",000000," ),"0,000000" )
IF (nPos:=_AScanX( ::aWorkbookBorders, c ) )==0
	AADD( ::aWorkbookBorders, c )
	nPos := LEN( ::aWorkbookBorders )
ENDIF
Return nPos

// lExact := .T.: format se upisuje tacno kako je zadat (Read() tako cuva npr.
// "dd.mm.yyyy" iz fajla, umesto ugradjenog formata koji zavisi od podesavanja Windows-a)
METHOD NewFormat( cFormat, lExact ) CLASS WorkBook
Local nPos := 0
IF HB_ISSTRING( cFormat )
	IF hb_hHaskey( aStdNumFormats, cFormat ) .AND. !( HB_ISLOGICAL( lExact ) .AND. lExact )
		nPos := aStdNumFormats[cFormat]
	ELSEIF (nPos:=_AScanX( ::aNumFormat, cFormat ) )==0
		AADD( ::aNumFormat, cFormat )
		nPos := LEN( ::aNumFormat )
	ENDIF
ENDIF
Return nPos

// the look of a conditional format (Conditional()): the color of the text, bold, the color of
// the cell, italic, underline, a thin border in a color; each one is optional ("" / NIL: as the cell)
METHOD NewCondStyle( cFontColor, lBold, cFillColor, lItalic, lUnderline, cBorderColor ) CLASS WorkBook
LOCAL cFont := "", cFill := "", cBorder := "", c, nPos, cSide
IF HB_ISLOGICAL( lBold ) .AND. lBold
	cFont += '<b/>'
ENDIF
IF HB_ISLOGICAL( lItalic ) .AND. lItalic
	cFont += '<i/>'
ENDIF
IF HB_ISLOGICAL( lUnderline ) .AND. lUnderline
	cFont += '<u/>'
ENDIF
IF HB_ISSTRING( cFontColor ) .AND. ! Empty( cFontColor )
	cFont += '<color rgb="' + _ARGB( cFontColor ) + '"/>'
ENDIF
IF HB_ISSTRING( cFillColor ) .AND. ! Empty( cFillColor )
	cFill := '<fill><patternFill patternType="solid"><fgColor rgb="' + _ARGB( cFillColor ) + '"/><bgColor rgb="' + _ARGB( cFillColor ) + '"/></patternFill></fill>'
ENDIF
IF HB_ISSTRING( cBorderColor ) .AND. ! Empty( cBorderColor )
	cBorder := '<border>'
	FOR EACH cSide IN { "left", "right", "top", "bottom" }
		cBorder += '<' + cSide + ' style="thin"><color rgb="' + _ARGB( cBorderColor ) + '"/></' + cSide + '>'
	NEXT
	cBorder += '</border>'
ENDIF
c := IIF( Empty( cFont ), "", '<font>' + cFont + '</font>' ) + cFill + cBorder
IF ( nPos := _AScanX( ::aCondStyles, c ) ) == 0
	AAdd( ::aCondStyles, c )
	nPos := LEN( ::aCondStyles )
ENDIF
RETURN nPos

METHOD NewFont( cFont, nFontSize, lBold, lItalic, lUnderline, lStrike, cRGB ) CLASS WorkBook
Local nPos
cFont := iif( HB_ISNIL( cFont ), "Calibri", cFont )
// HB_NTOC( n ) bez ",0": velicina 10.5 se ranije zaokruzivala na 11
cFont += ","+iif( HB_ISNIL( nFontSize ) .OR. !HB_ISNUMERIC( nFontSize ), "11", HB_NTOC(nFontSize) )
cFont += ","+iif( HB_ISNIL( lBold ) .OR. !HB_ISLOGICAL( lBold ), "0", iif(lBold,"1","0") )
cFont += ","+iif( HB_ISNIL( lItalic ) .OR. !HB_ISLOGICAL( lItalic ), "0", iif(lItalic,"1","0") ) 
cFont += ","+iif( HB_ISNIL( lUnderline ) .OR. !HB_ISLOGICAL( lUnderline ), "0", iif(lUnderline,"1","0") )
cFont += ","+iif( HB_ISNIL( lStrike ) .OR. !HB_ISLOGICAL( lStrike ), "0", iif(lStrike,"1","0") )
cFont += ","+iif( HB_ISNIL( cRGB ) .OR. !HB_ISSTRING( cRGB ), "000000", cRGB )
IF (nPos:=_AScanX( ::aWorkbookFonts, cFont ) )==0
	AADD( ::aWorkbookFonts, cFont )
	nPos := LEN( ::aWorkbookFonts )
ENDIF
Return nPos

METHOD NewStyle( nFont, nBorder, nFill, nVA, nHA, nFormat, nRotation, lWrap ) CLASS WorkBook
Local nPos, c
c := iif( HB_ISNIL( nFont ) .OR. !HB_ISNUMERIC( nFont ), "0", HB_NTOC(IIF(nFont==0,1,nFont)-1,0) )
c += ","+iif( HB_ISNIL( nBorder ) .OR. !HB_ISNUMERIC( nBorder ), "0", HB_NTOC(IIF(nBorder==0,1,nBorder)-1,0) )
c += ","+iif( HB_ISNIL( nFill ) .OR. !HB_ISNUMERIC( nFill ), "0", HB_NTOC(IIF(nFill==0,1,nFill)-1,0) )
c += ","+iif( HB_ISNIL( nHA ) .OR. !HB_ISNUMERIC( nHA ) .OR. nHA<1 .OR. nHa>3, "0", HB_NTOC(nHA,0) )
c += ","+iif( HB_ISNIL( nVA ) .OR. !HB_ISNUMERIC( nVA ) .OR. nVA<1 .OR. nVa>3, "0", HB_NTOC(nVA,0) )
c += ","+iif( HB_ISNIL( nFormat ) .OR. !HB_ISNUMERIC( nFormat ), "0", IIF( nFormat>0,HB_NTOS(nFormat+163), HB_NTOS(-nFormat) ) )
c += ","+iif( HB_ISNIL( nRotation ) .OR. !HB_ISNUMERIC( nRotation ), "0", HB_NTOC(nRotation,0) )
c += ","+iif( HB_ISLOGICAL( lWrap ) .AND. lWrap, "1", "0" )     // ranije je i .F. ukljucivao prelom
IF (nPos:=_AScanX( ::aWorkbookStyles, c ) )==0
	AADD( ::aWorkbookStyles, c )
	nPos := LEN( ::aWorkbookStyles )
ENDIF
Return nPos

METHOD New(cName) CLASS WorkBook
// bez imena (npr. samo za Read()) AT() je pucao sa argument error
IF ! HB_ISSTRING( cName )
	cName := ""
ENDIF
::cFilePath := iif( FilePath(cName)=="", DiskName() + hb_OSDriveSeparator() + hb_PS() + CurDir() + hb_PS() , FilePath(cName) )
::cName := iif( AT(::cFilePath,cName)>0, SUBSTR(cName,LEN(::cFilePath)+1), cName )
::aWorkSheetNames := {}  
::aWorkSheetObjects := {}
::aWorkSheetDrawingObjects := {}
::aWorkbookFonts := { "Calibri,11,0,0,0,0,000000" }
::aWorkbookStyles := {}
::aWorkbookFills := { "0,00000000,00000000", "0,00000000,00000000" }
::aWorkbookBorders := { "0,0,0,0,0" }
::aNumFormat := { }
::aUsedDrawingType := {  }
::nDrawings := 0
::aCondStyles := {}
::aSharedStrings := {}
::hSharedStrings := { => }       // podrazumevano binarno i case-sensitive
::_str_total := 0
::cCodePage := hb_cdpSelect()
// svaki workbook ima svoj temp folder (ranije fiksni "ExcelTemp" koji je New()
// brisao, pa su dva workbook-a istovremeno smetala jedan drugom). Folder prave
// tek Save()/Read(), da workbook koji se ne snimi ne ostavi smece u TEMP-u.
DO WHILE .T.
	::cTempDir := hb_DirSepToOS( hb_DirTemp() + "ExcelTemp_" + hb_NToS( hb_RandomInt( 1, 99999999 ) ) )
	IF ! hb_DirExists( ::cTempDir )
		EXIT
	ENDIF
ENDDO
Return Self

// Postojeci sheet ili novi. Ime se prilagodi pravilima Excela (bez : \ / ? * [ ],
// najvise 31 znak), a "Podaci" i "PODACI" su isti sheet - inace Excel odbija fajl.
METHOD WorkSheet( cName ) CLASS WorkBook
LOCAL oWorkSheet, nPos
cName := _SheetName( cName )
IF (nPos:=_SheetPos( ::aWorkSheetNames, cName ))==0
    oWorkSheet:= WorkSheet():New( cName )
	oWorkSheet:oParent := Self
	aADD( ::aWorkSheetNames, cName )
	aADD( ::aWorkSheetObjects, oWorkSheet )
	nPos := LEN( ::aWorkSheetNames )
ENDIF
Return ::aWorkSheetObjects[nPos]

METHOD Save() CLASS WorkBook
LOCAL cSep := hb_ps(), i, j, handle, c
LOCAL cFile := ::cFilePath + ::cName, lOk
LOCAL cDate := DTOS( DATE() ), cTime := TIME(), cDateTime := LEFT( cDate, 4 )+'-'+SUBSTR(cDate,5,2)+'-'+RIGHT(cDate,2)+'T'+LEFT(cTime,8)
   ::cError := ""
   IF Empty( ::cName )
      ::cError := "Nije zadato ime fajla: WorkBook():New( cFajl )"
      RETURN .F.
   ENDIF
   // fajl otvoren u Excelu ne moze da se obrise; ranije je Save() to precutao,
   // a pozivalac je zatim otvarao STARI fajl
   IF hb_FileExists( cFile ) .AND. FErase( cFile ) != 0
      ::cError := "Ne mogu da prepisem fajl " + cFile + ";Verovatno je otvoren u Excelu - zatvorite ga i pokusajte ponovo."
      RETURN .F.
   ENDIF
   MakeDir( ::cTempDir )
   MakeDir( ::cTempDir+cSep+"docProps" )
   MakeDir( ::cTempDir+cSep+"_rels" )   
   MakeDir( ::cTempDir+cSep+"xl" )
   MakeDir( ::cTempDir+cSep+"xl"+cSep+"_rels" )
   MakeDir( ::cTempDir+cSep+"xl"+cSep+"worksheets" )
   IF ::nDrawings>0 .OR. ::nCharts>0
		MakeDir( ::cTempDir+cSep+"xl"+cSep+"worksheets"+cSep+"_rels" )
		MakeDir( ::cTempDir+cSep+"xl"+cSep+"drawings" )
		MakeDir( ::cTempDir+cSep+"xl"+cSep+"drawings"+cSep+"_rels" )
   ENDIF
   IF ::nCharts>0
		MakeDir( ::cTempDir+cSep+"xl"+cSep+"charts" )
   ENDIF
   IF ::nDrawings>0
		MakeDir( ::cTempDir+cSep+"xl"+cSep+"media" )
		aEval( ::aWorkSheetObjects, { |arr| aEval( arr:aDrawingsObjects, { |arr1| _CopyImageToMedia( arr1:cPath, ::cTempDir, cSep, arr1:ID ) } ) } )
			/* __CopyFile( arr1:cPath, ::cTempDir+cSep+"xl"+cSep+"media"+cSep+"image"+arr1:ID+lower(hb_FNameExt( arr1:cPath )) ) } )  */
   ENDIF
// -------------------------------------------------------- [Content_Types].xml ------------------------------------------------------------------------------   
   handle := FCreate( ::cTempDir+cSep+"[Content_Types].xml" )
   FWrite( handle, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+chr(10) )
   FWrite( handle, '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'+chr(10) )
   IF ::nDrawings>0 .OR. ::nCharts>0
        //aEval( ::aUsedDrawingType, { |arr| FWrite( handle, '<Default Extension="'+arr+'" ContentType="image/'+arr+'"/>'+chr(10) ) } )
		aEval( ::aUsedDrawingType, { |ext| FWrite( handle, ;
			'<Default Extension="' + IIF( Lower(ext) == "jpg", "jpeg", Lower(ext) ) + ;
			'" ContentType="' + _MimeForImageExt( Lower(ext) ) + '"/>' + chr(10) ) } )		
		FWrite( handle, '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'+chr(10) )
		FWrite( handle, '<Default Extension="xml" ContentType="application/xml"/>'+chr(10) )
		// drawing<n>.xml za svaki sheet sa slikama ili grafikonima (ranije samo drawing1)
		FOR i := 1 TO LEN( ::aWorkSheetObjects )
			IF LEN( ::aWorkSheetObjects[i]:aDrawingsObjects ) > 0 .OR. LEN( ::aWorkSheetObjects[i]:aCharts ) > 0
				FWrite( handle, '<Override PartName="/xl/drawings/drawing'+HB_NTOS(i)+'.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'+chr(10) )
			ENDIF
			FOR j := 1 TO LEN( ::aWorkSheetObjects[i]:aCharts )
				FWrite( handle, '<Override PartName="/xl/charts/chart'+HB_NTOS(::aWorkSheetObjects[i]:aCharts[j]:ID)+'.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'+chr(10) )
			NEXT
		NEXT
   ENDIF
   FWrite( handle, '<Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'+chr(10) )
   FWrite( handle, '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'+chr(10) )
   FWrite( handle, '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'+chr(10) )   
   FWrite( handle, '<Override PartName="/xl/_rels/workbook.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'+chr(10) )
   FWrite( handle, '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'+chr(10) )
   FWrite( handle, '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'+chr(10) )
   FOR I:=1 TO LEN( ::aWorkSheetNames )
		FWrite( handle, '<Override PartName="/xl/worksheets/sheet'+HB_NTOS(i)+'.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />'+chr(10) )   
   NEXT I
   IF LEN( ::aSharedStrings ) > 0
		FWrite( handle, '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'+chr(10) )      
   ENDIF
   FWrite( handle, '</Types>' )	
   FClose( handle )   
   
// ------------------------------------------------------------ .rels ------------------------------------------------------------------------------   
   handle := FCreate( ::cTempDir+cSep+"_rels"+cSep+".rels" )
   FWrite( handle, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+chr(10) )   
   FWrite( handle, '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+chr(10) )
   FWrite( handle, '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'+chr(10) )
   FWrite( handle, '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'+chr(10) )   
   FWrite( handle, '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'+chr(10) )
   FWrite( handle, '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="xl/styles.xml"/>'+chr(10) )   
   FWrite( handle, '</Relationships>' )   
   FClose( handle )
   
// ------------------------------------------------------------ app.xml ------------------------------------------------------------------------------   
   handle := FCreate( ::cTempDir+cSep+"docProps"+cSep+"app.xml" )
   FWrite( handle, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+chr(10) )   
   FWrite( handle, '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'+chr(10) )
   FWrite( handle, '<Application>'+XmlEsc(::ToUTF8(_S(::Application)))+'</Application>'+chr(10) )
   FWrite( handle, '<Company>'+XmlEsc(::ToUTF8(_S(::Company)))+'</Company>'+chr(10) )
   FWrite( handle, '<Template></Template>'+chr(10) )      
   FWrite( handle, '<TotalTime>0</TotalTime>'+chr(10) )
   FWrite( handle, '</Properties>' )   
   FClose( handle )         

// ------------------------------------------------------------ core.xml ------------------------------------------------------------------------------   
   handle := FCreate( ::cTempDir+cSep+"docProps"+cSep+"core.xml" )
   FWrite( handle, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+chr(10) )   
   FWrite( handle, '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'+chr(10) )
   FWrite( handle, '<dc:creator>Srdan Dragojlovic, Kosovo is Serbia</dc:creator>'+chr(10) )
   FWrite( handle, '<cp:lastModifiedBy></cp:lastModifiedBy>'+chr(10) )
   FWrite( handle, '<dcterms:created xsi:type="dcterms:W3CDTF">'+cDateTime+'Z</dcterms:created>'+chr(10) )
   FWrite( handle, '<dcterms:modified xsi:type="dcterms:W3CDTF">'+cDateTime+'Z</dcterms:modified>'+chr(10) )
   FWrite( handle, '</cp:coreProperties>' )   
   FClose( handle )


// ------------------------------------------------------------ workbook.xml.rels ------------------------------------------------------------------------------   
   handle := FCreate( ::cTempDir+cSep+"xl"+cSep+"_rels"+cSep+"workbook.xml.rels" )
   FWrite( handle, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+chr(10) )   
   FWrite( handle, '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+chr(10) )
   FOR I:=1 TO LEN( ::aWorkSheetNames )
		FWrite( handle, '<Relationship Id="rId'+HB_NTOS(i)+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet'+HB_NTOS(i)+'.xml"/>'+chr(10) )   
   NEXT I
   FWrite( handle, '<Relationship Id="rId'+HB_NTOS(i)+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml" />'+chr(10) )
   IF LEN( ::aSharedStrings ) > 0
		FWrite( handle, '<Relationship Id="rId'+HB_NTOS(++i)+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'+chr(10) )
   ENDIF 	
   FWrite( handle, '</Relationships>' )   
   FClose( handle )         
   
   // ------------------------------------------------------------ workbook.xml ------------------------------------------------------------------------------   
	handle := FCreate( ::cTempDir+cSep+"xl"+cSep+"workbook.xml" )
	FWrite( handle, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+chr(10) )   
	FWrite( handle, '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15 xr xr6 xr10 xr2" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr6="http://schemas.microsoft.com/office/spreadsheetml/2016/revision6" xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2">'+chr(10) )
	FWrite( handle, '<fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="28429"/>'+chr(10) )
	FWrite( handle, '<workbookPr defaultThemeVersion="166925"/>'+chr(10) )
	// (ranije je ovde u svaki fajl upisivana putanja C:\Dropbox\XLSX\ - x15ac:absPath)
	FWrite( handle, '<xr:revisionPtr revIDLastSave="0" documentId="8_{063AD3CC-BC4D-49E0-9FE7-D121456950CA}" xr6:coauthVersionLast="47" xr6:coauthVersionMax="47" xr10:uidLastSave="{00000000-0000-0000-0000-000000000000}"/>'+chr(10) )
	FWrite( handle, '<bookViews>'+chr(10) )
	FWrite( handle, '<workbookView xWindow="-108" yWindow="-108" windowWidth="23256" windowHeight="12456" xr2:uid="{00000000-000D-0000-FFFF-FFFF00000000}"/>'+chr(10) )
	FWrite( handle, '</bookViews>'+chr(10) )
	FWrite( handle, '<sheets>'+chr(10) )
	FOR I:=1 TO LEN( ::aWorkSheetNames )
		FWrite( handle, '<sheet name="'+XmlEsc(::ToUTF8(::aWorkSheetNames[i]))+'" sheetId="'+HB_NTOS(i)+'" r:id="rId'+HB_NTOS(i)+'"/>'+chr(10) )   
	NEXT I
	FWrite( handle, '</sheets>'+chr(10) )
	// redovi koji se ponavljaju na vrhu svake strane: oSheet:nTitlesRows, a
	// oWorkBook:nTitlesRows vazi za prvi sheet (ranije uvek 'Sheet1', pa nije radilo
	// ako se prvi sheet drugacije zove)
	c := ""
	FOR I:=1 TO LEN( ::aWorkSheetObjects )
		j := ::aWorkSheetObjects[i]:nTitlesRows
		IF ! HB_ISNUMERIC( j ) .OR. j <= 0
			j := IIF( i == 1 .AND. HB_ISNUMERIC( ::nTitlesRows ), ::nTitlesRows, 0 )
		ENDIF
		IF j > 0
			c += '	<definedName name="_xlnm.Print_Titles" localSheetId="'+HB_NTOS(i-1)+'">'+ ;
			     XmlEsc( "'" + StrTran( ::ToUTF8( ::aWorkSheetNames[i] ), "'", "''" ) + "'!$1:$" + HB_NTOS( j ) )+'</definedName>'+chr(10)
		ENDIF
	NEXT I
	IF ! Empty( c )
		FWrite( handle, '<definedNames>'+chr(10)+c+'</definedNames>'+chr(10) )
	ENDIF
	FWrite( handle, '<calcPr calcMode="auto" fullCalcOnLoad="1" calcId="0"/>'+chr(10) )
	/*	
	FWrite( handle, '<extLst>'+chr(10) )
	FWrite( handle, '<ext uri="{B58B0392-4F1F-4190-BB64-5DF3571DCE5F}" xmlns:xcalcf="http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures">'+chr(10) )
	FWrite( handle, '<xcalcf:calcFeatures>'+chr(10) )
	FWrite( handle, '<xcalcf:feature name="microsoft.com:RD"/>'+chr(10) )
	FWrite( handle, '<xcalcf:feature name="microsoft.com:Single"/>'+chr(10) )
	FWrite( handle, '<xcalcf:feature name="microsoft.com:FV"/>'+chr(10) )
	FWrite( handle, '<xcalcf:feature name="microsoft.com:CNMTM"/>'+chr(10) )
	FWrite( handle, '<xcalcf:feature name="microsoft.com:LET_WF"/>'+chr(10) )
	FWrite( handle, '</xcalcf:calcFeatures>'+chr(10) )
	FWrite( handle, '</ext>'+chr(10) )
	FWrite( handle, '</extLst>'+chr(10) )
	*/
	FWrite( handle, '</workbook>' )   
	FClose( handle )         

   // ------------------------------------------------------------ sharedStrings.xml ------------------------------------------------------------------------------------------------------------------------   
   IF ::_str_total>0
		handle := FCreate( ::cTempDir+cSep+"xl"+cSep+"sharedStrings.xml" )
		FWrite( handle, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+chr(10) )   
        FWrite( handle, '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="'+HB_NTOS(::_str_total)+'" uniqueCount="'+HB_NTOS(LEN(::aSharedStrings))+'">'+chr(10) )
		FOR I:=1 TO LEN(::aSharedStrings)
			FWrite( handle, '<si><t xml:space="preserve">'+XmlEsc(::aSharedStrings[I])+'</t></si>'+chr(10) )
		NEXT I
		FWrite( handle, '</sst>' )   
		FClose( handle )         
   ENDIF 		

   // ------------------------------------------------------------ styles.xml ------------------------------------------------------------------------------------------------------------------------   
	handle := FCreate( ::cTempDir+cSep+"xl"+cSep+"styles.xml" )
	FWrite( handle, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+chr(10) )   
    FWrite( handle, '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2 xr" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">'+chr(10) )
	IF LEN( ::aNumFormat ) > 0
		FWrite( handle, '<numFmts count="'+HB_NTOS(LEN(::aNumFormat))+'">'+chr(10))
		FOR I:=1 TO LEN( ::aNumFormat )
			// XmlEsc: format kao '#,##0 "din"' ili '[<1000]0' ranije je kvario styles.xml
			FWrite( handle, '<numFmt numFmtId="'+HB_NTOS(I+163)+'" formatCode="'+XmlEsc(::ToUTF8(::aNumFormat[I]))+'"/>'+chr(10) )
		NEXT I
		FWrite( handle, '</numFmts>'+chr(10))
	ENDIF
	FWrite( handle, '<fonts count="'+HB_NTOS(LEN(::aWorkbookFonts))+'">'+chr(10) )	
	//FWrite( handle, '<font/>'+chr(10) )
	FOR I:=1 TO LEN( ::aWorkbookFonts )
	    J:= hb_ATokens( ::aWorkbookFonts[i], "," )	
		FWrite( handle, '<font>' )		
		FWrite( handle, '<name val="'+XmlEsc(::ToUTF8(j[1]))+'"/>' )
		FWrite( handle, '<sz val="'+j[2]+'"/>' )		
		IF j[3]=="1";FWrite( handle, '<b/>' );ENDIF
		IF j[4]=="1";FWrite( handle, '<i/>' );ENDIF
		IF j[5]=="1";FWrite( handle, '<u/>' );ENDIF
		IF j[6]=="1";FWrite( handle, '<strike/>' );ENDIF		
		FWrite( handle, '<color rgb="'+_ARGB(j[7])+'"/>' )
		FWrite( handle, '</font>'+chr(10) )
	NEXT I
	FWrite( handle, '</fonts>'+chr(10) )		
	FWrite( handle, '<fills count="'+HB_NTOS(LEN(::aWorkbookFills))+'">'+chr(10) )		
	FWrite( handle, '<fill><patternFill patternType="none"/></fill>'+chr(10)+'<fill><patternFill patternType="gray125"/></fill>'+chr(10) )	
	FOR I:=3 TO LEN( ::aWorkbookFills )
	    J:= hb_ATokens( ::aWorkbookFills[i], "," )
		FWrite( handle, '<fill><patternFill patternType="'+aPattern[val(j[1])]+'"><fgColor rgb="'+_ARGB(j[2])+'"/><bgColor rgb="'+_ARGB(j[3])+'"/></patternFill></fill>'+chr(10) )
	NEXT I
	FWrite( handle, '</fills>'+chr(10) )		
	FWrite( handle, '<borders count="'+HB_NTOS(LEN(::aWorkbookBorders))+'">'+chr(10) )		
	FWrite( handle, '<border><left/><right/><top/><bottom/><diagonal/></border>'+chr(10) )		
	FOR I:=2 TO LEN( ::aWorkbookBorders )
	    J:= hb_ATokens( ::aWorkbookBorders[i], "," )
		FWrite( handle, '<border>' )
		FWrite( handle, iif(val(j[1])>0, '<left style="'+aBordersStyle[val(j[1])]+'"'+iif(j[2]!="000000", '><color rgb="'+_ARGB(j[2])+'"/></left>', '/>' ), '<left/>' ) )
		FWrite( handle, iif(val(j[3])>0, '<right style="'+aBordersStyle[val(j[3])]+'"'+iif(j[4]!="000000", '><color rgb="'+_ARGB(j[4])+'"/></right>', '/>' ), '<right/>' ) )
		FWrite( handle, iif(val(j[5])>0, '<top style="'+aBordersStyle[val(j[5])]+'"'+iif(j[6]!="000000", '><color rgb="'+_ARGB(j[6])+'"/></top>', '/>' ), '<top/>' ) )
		FWrite( handle, iif(val(j[7])>0, '<bottom style="'+aBordersStyle[val(j[7])]+'"'+iif(j[8]!="000000", '><color rgb="'+_ARGB(j[8])+'"/></bottom>', '/>' ), '<bottom/>' ) )
		FWrite( handle, iif(val(j[9])>0, '<diagonal style="'+aBordersStyle[val(j[9])]+'"'+iif(j[10]!="000000", '><color rgb="'+_ARGB(j[10])+'"/></diagonal>', '/>' ), '<diagonal/>' ) )
		FWrite( handle, '</border>'+chr(10) )
	NEXT I
	FWrite( handle, '</borders>'+chr(10) )		
	FWrite( handle, '<cellStyleXfs count="1">'+chr(10) )		
	FWrite( handle, '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>'+chr(10) )		
	FWrite( handle, '</cellStyleXfs>'+chr(10) )		
	FWrite( handle, '<cellXfs count="'+HB_NTOS(LEN(::aWorkbookStyles)+1)+'">'+chr(10) )		
	FWrite( handle, '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'+chr(10) )		
	FOR I:=1 TO LEN( ::aWorkbookStyles )
	    J:= hb_ATokens( ::aWorkbookStyles[i], "," )
		FWrite( handle, '<xf ' )
		FWrite( handle, iif(j[6]=="0",'numFmtId="0" ','numFmtId="'+j[6]+'" ') )		
		FWrite( handle, iif(j[1]=="0",'fontId="0" ','fontId="'+j[1]+'" ') )
		FWrite( handle, iif(j[3]=="0",'fillId="0" ','fillId="'+j[3]+'" ') )
		FWrite( handle, iif(j[2]=="0",'borderId="0" ','borderId="'+j[2]+'" ') )
		FWrite( handle, 'xfId="0" ')
		FWrite( handle, iif(j[6]=="0", '', 'applyNumberFormat="1" ') )				
		FWrite( handle, iif(j[1]=="0", '', 'applyFont="1" ') )
		FWrite( handle, iif(j[3]=="0", '', 'applyFill="1" ') )		
		FWrite( handle, iif(j[2]=="0", '', 'applyBorder="1" ') )
		FWrite( handle, iif(j[4]=="0" .AND. j[5]=="0" .AND. j[7]=="0" .AND. j[8]=="0", '/>', 'applyAlignment="1">' ) )
		IF !(j[4]=="0" .AND. j[5]=="0" .AND. j[7]=="0" .AND. j[8]=="0")
			// samo zadato poravnanje (ranije uvek horizontal="general" vertical="bottom",
			// sto LibreOffice za rotiran tekst prikazuje drugacije nego podrazumevano)
			c := IIF( j[4]!="0", ' horizontal="'+aHorizontal[val(j[4])]+'"', '' )+IIF( j[5]!="0", ' vertical="'+aVertical[val(j[5])]+'"', '' )
			c += IIF( j[7]!="0", ' textRotation="'+j[7]+'"', '' )+IIF( j[8]=="0", '', ' wrapText="1"' )
			FWrite( handle, '<alignment'+c+'/>' )
			FWrite( handle, '</xf>'+chr(10) )
		ELSE
			FWrite( handle, chr(10) )
		ENDIF
	NEXT I	
	FWrite( handle, '</cellXfs>'+chr(10) )		
	FWrite( handle, '<cellStyles count="1">'+chr(10) )
	FWrite( handle, '<cellStyle name="Normal" xfId="0" builtinId="0"/>'+chr(10) )
	FWrite( handle, '</cellStyles>'+chr(10) )	
	IF LEN( ::aCondStyles ) == 0
		FWrite( handle, '<dxfs count="0"/>'+chr(10) )
	ELSE
		FWrite( handle, '<dxfs count="'+HB_NTOS(LEN(::aCondStyles))+'">'+chr(10) )
		FOR I := 1 TO LEN( ::aCondStyles )
			FWrite( handle, '<dxf>'+::aCondStyles[I]+'</dxf>'+chr(10) )
		NEXT
		FWrite( handle, '</dxfs>'+chr(10) )
	ENDIF
	FWrite( handle, '<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>'+chr(10) )
	FWrite( handle, '<extLst>'+chr(10) )
		FWrite( handle, '<ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/></ext>'+chr(10) )
		FWrite( handle, '<ext uri="{9260A510-F301-46a8-8635-F512D64BE5F5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><x15:timelineStyles defaultTimelineStyle="TimeSlicerStyleLight1"/></ext>'+chr(10) )
	FWrite( handle, '</extLst>'+chr(10) )
	FWrite( handle, '</styleSheet>' )   
	FClose( handle )         
   // -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	FOR I:=1 TO LEN( ::aWorkSheetNames )
		::aWorkSheetObjects[i]:WriteWorksheet(i, ::cTempDir+cSep+"xl"+cSep+"worksheets"+cSep)
	NEXT I
	// bez DirChange(): zip radi sa apsolutnim putanjama, pa greska usred snimanja
	// ne ostavlja program u temp folderu (relativne putanje do DBF-ova)
	lOk := MyZip( cFile, ::cTempDir ) .AND. hb_FileExists( cFile )
	IF ! lOk
		::cError := "Greska pri pravljenju fajla " + cFile
	ENDIF
	hb_DirRemoveAll( ::cTempDir )
Return lOk

CLASS DrawingML
DATA cName
DATA oWorkSheet
DATA oWorkBook
DATA Type
DATA cPath
DATA nCol1, nRow1
DATA nCol2, nRow2
DATA nX, nY
DATA nSizeX INIT "0"
DATA nSizeY INIT "0"
DATA nOffColFrom INIT "0"
DATA nOffColTo INIT "0"
DATA nOffRowFrom INIT "0"
DATA nOffRowTo INIT "0"
DATA ID
DATA editAs INIT "absolute"
METHOD New(cName)
METHOD OneCellAnchor( uAddr1, nSizeX, nSizeY, cPath )
METHOD TwoCellAnchor( uAddr1, uAddr2, cPath )
METHOD AbsoluteAnchor( nX, nY, nSizeX, nSizeY, cPath )
METHOD AnchorXML( nIdx )
ENDCLASS

// XML sidra jedne slike za drawing<n>.xml. Ranije je WriteDrawing() za svaku
// sliku iznova pisao ceo drawing (sa podacima tekuce slike i </xdr:wsDr> u
// petlji), pa vise slika ili slika + grafikon na istom sheet-u nisu radili.
METHOD AnchorXML( nIdx ) CLASS DrawingML
LOCAL c := "", t := Chr(9), nl := Chr(10), cEnd
DO CASE
CASE ::Type == 0
	c += '<xdr:absoluteAnchor>'+nl
	c += t+'<xdr:pos x="'+_S(::nX)+'" y="'+_S(::nY)+'" />'+nl
	c += t+'<xdr:ext cx="'+_S(::nSizeX)+'" cy="'+_S(::nSizeY)+'" />'+nl
	cEnd := '</xdr:absoluteAnchor>'
CASE ::Type == 1
	c += '<xdr:oneCellAnchor>'+nl
	c += _AnchorPt( "from", ::nCol1, ::nOffColFrom, ::nRow1, ::nOffRowFrom )
	c += t+'<xdr:ext cx="'+_S(::nSizeX)+'" cy="'+_S(::nSizeY)+'" />'+nl
	cEnd := '</xdr:oneCellAnchor>'
OTHERWISE
	c += '<xdr:twoCellAnchor editAs="'+::editAs+'">'+nl
	c += _AnchorPt( "from", ::nCol1, ::nOffColFrom, ::nRow1, ::nOffRowFrom )
	c += _AnchorPt( "to", ::nCol2, ::nOffColTo, ::nRow2, ::nOffRowTo )
	cEnd := '</xdr:twoCellAnchor>'
ENDCASE
c += t+'<xdr:pic>'+nl
c += t+'<xdr:nvPicPr>'+nl
c += t+t+'<xdr:cNvPr id="'+HB_NTOS(nIdx-1)+'" name="'+XmlEsc(_S(::cName))+'" descr=""></xdr:cNvPr>'+nl
c += t+t+'<xdr:cNvPicPr/>'+nl
c += t+'</xdr:nvPicPr>'+nl
c += t+t+'<xdr:blipFill><a:blip r:embed="rId'+::ID+'" cstate="print" /><a:stretch/></xdr:blipFill>'+nl
c += t+t+'<xdr:spPr>'+nl
// Ranije je ovde bila zakucana velicina (1971360 x 783720 EMU) za svaku sliku, pa
// LibreOffice nije razvlacio sliku na zadati opseg celija. Za twoCellAnchor
// velicinu odredjuju celije, a za ostale sidra zadata velicina.
IF ::Type != 2
	c += t+t+t+'<a:xfrm>'+nl
	c += t+t+t+t+'<a:off x="'+_S(IIF(::Type == 0, ::nX, "0"))+'" y="'+_S(IIF(::Type == 0, ::nY, "0"))+'"/>'+nl
	c += t+t+t+t+'<a:ext cx="'+_S(::nSizeX)+'" cy="'+_S(::nSizeY)+'"/>'+nl
	c += t+t+t+'</a:xfrm>'+nl
ENDIF
c += t+t+t+'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'+nl
c += t+t+t+'<a:ln><a:noFill/></a:ln>'+nl
c += t+t+'</xdr:spPr>'+nl
c += t+'</xdr:pic>'+nl
c += '<xdr:clientData />'+nl
c += cEnd+nl
RETURN c

METHOD New(cName) CLASS DrawingML
::cName := cName
Return Self

METHOD AbsoluteAnchor( nX, nY, nSizeX, nSizeY, cPath ) CLASS DrawingML
LOCAL cType := lower( SUBSTR( hb_FNameExt( cPath ), 2 ) )
::Type := 0
::nX := HB_NTOS(nX)
::nY := HB_NTOS(nY)
::nSizeX := HB_NTOS(nSizeX)
::nSizeY := HB_NTOS(nSizeY)
::cPath := cPath
IF ASCAN( ::oWorkBook:aUsedDrawingType, {|aVal| aVal == cType } ) == 0
    aADD( ::oWorkBook:aUsedDrawingType, cType )
ENDIF
Return Self

// Slika od gornjeg levog ugla celije uAddr1 ("C3" ili { nRed, nKol }), velicine
// nSizeX x nSizeY EMU (1 cm = 360000, 1 piksel = 9525)
METHOD OneCellAnchor( uAddr1, nSizeX, nSizeY, cPath ) CLASS DrawingML
LOCAL nRow := 0, nCol := 0
LOCAL cType := lower( SUBSTR( hb_FNameExt( cPath ), 2 ) )
::Type := 1
// ranije je niz { nRed, nKol } uziman bez -1, pa je slika bila red/kolonu dalje
_CellRC( uAddr1, @nRow, @nCol )
::nRow1 := HB_NTOS( nRow - 1 )
::nCol1 := HB_NTOS( nCol - 1 )
::cPath := cPath
::nSizeX := HB_NTOS(nSizeX)
::nSizeY := HB_NTOS(nSizeY)
IF ASCAN( ::oWorkBook:aUsedDrawingType, {|aVal| aVal == cType } ) == 0
    aADD( ::oWorkBook:aUsedDrawingType, cType )
ENDIF
Return Self

// Slika razvucena od gornjeg levog ugla celije uAddr1 do donjeg desnog ugla celije uAddr2
METHOD TwoCellAnchor( uAddr1, uAddr2, cPath ) CLASS DrawingML
LOCAL nRow := 0, nCol := 0
LOCAL cType := lower( SUBSTR( hb_FNameExt( cPath ), 2 ) )
::Type := 2
_CellRC( uAddr1, @nRow, @nCol )
::nRow1 := HB_NTOS( nRow - 1 )
::nCol1 := HB_NTOS( nCol - 1 )
nRow := 0
nCol := 0
_CellRC( uAddr2, @nRow, @nCol )
::nRow2 := HB_NTOS( nRow )
::nCol2 := HB_NTOS( nCol )
::cPath := cPath
IF ASCAN( ::oWorkBook:aUsedDrawingType, {|aVal| aVal == cType } ) == 0
    aADD( ::oWorkBook:aUsedDrawingType, cType )
ENDIF
Return Self

CLASS WorkSheet
DATA cName
DATA oParent
DATA paperSize INIT 9
DATA lLandscape INIT .F.
DATA horizontalDpi INIT 300
DATA verticalDpi INIT 300
DATA leftMargin INIT 0.2
DATA rightMargin INIT 0.2
DATA topMargin INIT 0.2
DATA bottomMargin INIT 0.2
DATA headerMargin INIT 0.1
DATA footerMargin INIT 0.1
DATA fitToPage INIT .F.
DATA fitToHeight INIT .F.
DATA fitToWidth INIT .F.
DATA zoom INIT 100
DATA zoom_scale_normal INIT .T.
DATA print_scale INIT 100
DATA right_to_left INIT .F.
DATA show_zeros INIT .T.
DATA leading_zeros INIT .F.
DATA nTitlesRows INIT 0          // redovi 1..n se ponavljaju na vrhu svake odstampane strane
DATA nMaxRow PROTECTED
DATA nMaxCol PROTECTED
DATA aData PROTECTED
DATA lChar PROTECTED
DATA aMergeCells PROTECTED
DATA aStyle PROTECTED
DATA aCols PROTECTED
DATA aRows PROTECTED
DATA aRowsDetail PROTECTED
DATA aCharts 
DATA aDrawingsObjects 
DATA aDrawingsNames 
DATA lHeader PROTECTED
DATA lFooter PROTECTED
DATA aHeader PROTECTED
DATA aFooter PROTECTED
DATA aCond PROTECTED              // conditional formats: { cSqref, cRule (priority "%P%") }
METHOD New(cName)
METHOD WriteWorksheet(n, cPath)
METHOD Cell(uAddr, xValue, nStyle)
METHOD Formula(uAddr, cFormula, nStyle)
METHOD WriteCell(uAddr, xValue, nStyle)
METHOD ReadCell(uAddr)
METHOD CellFormula( uAddr )
METHOD Range( xRange )
METHOD ToArray()
METHOD LastRow()
METHOD LastCol()
METHOD MergedCells()
METHOD MergeCell(uAddr)
METHOD CellStyle( uAddr, nStyle )
METHOD ColumnsWidth( nMin, nMax, nWidth, lHidden )
METHOD ColumnWidths()
METHOD AutoFit( xFrom, xTo, nMaxWidth )
METHOD Conditional( cRange, cOp, xValue1, xValue2, nStyle )
METHOD ColorScale( cRange, cMinColor, cMaxColor, cMidColor )
METHOD DataBar( cRange, cColor )
METHOD AddHeader( cLeft, cCenter, cRight )
METHOD AddFooter( cLeft, cCenter, cRight )
METHOD RowDetail( nRow, nHeight, nStyle, lHide )
METHOD AddChart( cTitle, cType, cCatRef, xValRef, uFrom, uTo, aNames )
METHOD WriteDrawingPart( n )
METHOD Drawing( cName )
METHOD ReadWorkSheet( cFileName, aSST, aKinds, aXfMap )
ENDCLASS

METHOD Drawing( cName ) CLASS WorkSheet
LOCAL oDrawing, nPos 
IF (nPos:=_AScanX( ::aDrawingsNames, cName ))==0
    oDrawing:= DrawingML():New( cName )
	oDrawing:oWorkSheet := Self
	oDrawing:oWorkBook := Self:oParent
	aADD( ::aDrawingsNames, cName )
	aADD( ::aDrawingsObjects, oDrawing )
	nPos := LEN( ::aDrawingsNames )
	oDrawing:ID := HB_NTOS(::oParent:AddDrawing())
ENDIF
Return ::aDrawingsObjects[nPos]

METHOD New(cName) CLASS WorkSheet
::cName := cName
::aData := {}
::aMergeCells := {}
::aCols := {}
::aRows := {}
::aRowsDetail := {}
::aStyle := {}
::aCharts := {}
::aDrawingsNames := {}
::aDrawingsObjects := {}
::nMaxRow := 0
::nMaxCol := 0
::lChar := .F.
::lHeader := .F.
::lFooter := .F.
::aHeader := { "", "", "" }
::aFooter := { "", "", "" }
::aCond := {}
Return Self


METHOD RowDetail( nRow, nHeight, nStyle, lHide ) CLASS WorkSheet
LOCAL i
IF HB_ISNUMERIC(nRow) .AND. nRow >= 1
    IF LEN(::aRowsDetail)<nRow
		FOR I := LEN(::aRowsDetail) + 1 TO nRow
			aAdd( ::aRowsDetail, { 0, 0, .F. } )
		NEXT I
	ENDIF
	::aRowsDetail[nRow] := { IIF(HB_ISNIL(nHeight),0,IIF(HB_ISNUMERIC(nHeight),nHeight,0)), IIF(HB_ISNIL(nStyle),0,IIF(HB_ISNUMERIC(nStyle),nStyle,0)), IIF(HB_ISNIL(lHide),.F.,IIF(HB_ISLOGICAL(lHide),lHide,.F.)) } 
ENDIF
RETURN Self

//METHOD AddChart( uAddr, oChart, nX, nY, nX_Scale, nY_Scale ) CLASS WorkSheet
//RETURN Self

METHOD MergeCell( uAddr ) CLASS WorkSheet
LOCAL nCol1 := 0, nCol2 := 0, nRow1 := 0, nRow2 := 0, adr, I, J, K
IF HB_ISSTRING( uAddr )
	// "a1:c1" i "$A$1:$C$1" -> "A1:C1" (Excel prihvata samo veliko slovo bez $)
	uAddr := Upper( StrTran( AllTrim( uAddr ), "$", "" ) )
ENDIF
adr := IIF( HB_ISSTRING( uAddr ), hb_ATokens( uAddr, ":" ), {} )
IF LEN( adr ) == 2 .AND. _AScanX( ::aMergeCells, uAddr )==0
	WORKSHEET_RC( adr[1], @nRow1, @nCol1 )
	WORKSHEET_RC( adr[2], @nRow2, @nCol2 )
	IF nRow1 < 1 .OR. nCol1 < 1 .OR. nRow2 < nRow1 .OR. nCol2 < nCol1
		RETURN Self
	ENDIF
	AADD(  ::aMergeCells, uAddr )
    FOR I:=nRow1 TO nRow2
		FOR J:=nCol1 TO nCol2
			::nMaxCol := IIF( J > ::nMaxCol, J, ::nMaxCol )
			::nMaxRow := IIF( I > ::nMaxRow, I, ::nMaxRow )
			IF LEN( ::aData ) < I
				FOR k := LEN(::aData)+1 TO I 
					AADD( ::aData, {} )
				NEXT k
			ENDIF
			IF LEN( ::aData[I] ) < J
				FOR k := LEN(::aData[I])+1 TO J
					AADD( ::aData[I], NIL )
				NEXT k	   
			ENDIF
			IF ::aData[I,J]==NIL
				::aData[I,J] := CHR(0)
			ENDIF
		NEXT J
	NEXT I
ENDIF
Return Self

METHOD AddHeader( cLeft, cCenter, cRight )
IF HB_ISSTRING( cLeft )
   ::aHeader[1] := cLeft
ENDIF
IF HB_ISSTRING( cCenter )
   ::aHeader[2] := cCenter
ENDIF
IF HB_ISSTRING( cRight )
   ::aHeader[3] := cRight
ENDIF
::lHeader := .T.
RETURN Self

METHOD AddFooter( cLeft, cCenter, cRight )
IF HB_ISSTRING( cLeft )
   ::aFooter[1] := cLeft
ENDIF
IF HB_ISSTRING( cCenter )
   ::aFooter[2] := cCenter
ENDIF
IF HB_ISSTRING( cRight )
   ::aFooter[3] := cRight
ENDIF
::lFooter := .T.
RETURN Self

// sirina kolona nMin..nMax; lHidden := .T. sakriva kolone
// sirine kolona kako su zadate ili procitane: { { nMin, nMax, nWidth, lHidden }, ... }
METHOD ColumnWidths() CLASS WorkSheet
RETURN AClone( ::aCols )

METHOD ColumnsWidth( nMin, nMax, nWidth, lHidden ) CLASS WorkSheet
IF HB_ISNUMERIC(nMin) .AND. HB_ISNUMERIC(nMax) .AND. HB_ISNUMERIC(nWidth)
	AADD( ::aCols, { nMin, nMax, nWidth, HB_ISLOGICAL( lHidden ) .AND. lHidden } )
ENDIF
Return Self

// Conditional formatting, as Excel's Home - Conditional Formatting. nStyle: oExcel:NewCondStyle().
//   Conditional( "B2:B100", "<", 0, nStyle )          < <= > >= = <> : the value of the cell
//   Conditional( "C2:C100", "between", 10, 20, nStyle )   also "notBetween"
//   Conditional( "A2:A100", "text", "Kraljevo", nStyle )  contains; "notText", "begins", "ends"
//   Conditional( "A2:D100", "formula", "$D2>1000", nStyle )  for the top left cell of the range
//   Conditional( "E2:E100", "duplicate", , nStyle )   also "unique", "blank", "notBlank",
//                                                     "aboveAverage", "belowAverage"
//   Conditional( "F2:F100", "top", 10, nStyle )       the 10 largest; "bottom", "top%", "bottom%"
// A value: a number, a date, .T. / .F., a text, or a formula as "=$H$1". The style may also be
// the argument after the last value. The ranges: "A1:B5", "A1:B5 D1:D5" (or with commas).
METHOD Conditional( cRange, cOp, xValue1, xValue2, nStyle ) CLASS WorkSheet
LOCAL cSq := _CfRange( cRange ), cTL, cRule, cO, hOp, cT, nN
hOp := { "<" => "lessThan", "<=" => "lessThanOrEqual", ">" => "greaterThan", ">=" => "greaterThanOrEqual", ;
         "=" => "equal", "==" => "equal", "<>" => "notEqual", "!=" => "notEqual", ;
         "between" => "between", "notbetween" => "notBetween" }
cO := Lower( AllTrim( hb_defaultValue( cOp, "" ) ) )
IF Empty( cSq ) .OR. Empty( cO )
	RETURN Self
ENDIF
IF nStyle == NIL .AND. !( cO == "between" .OR. cO == "notbetween" )
	nStyle := xValue2                           // one value (or none): the style after it
ENDIF
cTL := hb_ATokens( hb_ATokens( cSq, " " )[ 1 ], ":" )[ 1 ]      // the top left cell
IF HB_ISSTRING( xValue1 )                        // a text of a DBF: UTF-8, as the cells
	xValue1 := ::oParent:ToUTF8( xValue1 )
ENDIF
IF HB_ISSTRING( xValue2 )
	xValue2 := ::oParent:ToUTF8( xValue2 )
ENDIF
cRule := '<cfRule type="%T%"' + IIF( HB_ISNUMERIC( nStyle ) .AND. nStyle >= 1, ' dxfId="' + HB_NTOS( nStyle - 1 ) + '"', '' ) + ' priority="%P%"'
DO CASE
CASE hb_HHasKey( hOp, cO )
	cRule := StrTran( cRule, "%T%", "cellIs" ) + ' operator="' + hOp[ cO ] + '">' + ;
	         '<formula>' + XmlEsc( _CfValue( xValue1 ) ) + '</formula>' + ;
	         IIF( cO == "between" .OR. cO == "notbetween", '<formula>' + XmlEsc( _CfValue( xValue2 ) ) + '</formula>', '' )
CASE cO == "text" .OR. cO == "nottext" .OR. cO == "begins" .OR. cO == "ends"
	cT := ::oParent:ToUTF8( hb_CStr( xValue1 ) )
	DO CASE
	CASE cO == "text"
		cRule := StrTran( cRule, "%T%", "containsText" ) + ' operator="containsText"'
		cO := 'NOT(ISERROR(SEARCH(' + _CfStr( cT ) + ',' + cTL + ')))'
	CASE cO == "nottext"
		cRule := StrTran( cRule, "%T%", "notContainsText" ) + ' operator="notContains"'
		cO := 'ISERROR(SEARCH(' + _CfStr( cT ) + ',' + cTL + '))'
	CASE cO == "begins"
		cRule := StrTran( cRule, "%T%", "beginsWith" ) + ' operator="beginsWith"'
		cO := 'LEFT(' + cTL + ',LEN(' + _CfStr( cT ) + '))=' + _CfStr( cT )
	OTHERWISE
		cRule := StrTran( cRule, "%T%", "endsWith" ) + ' operator="endsWith"'
		cO := 'RIGHT(' + cTL + ',LEN(' + _CfStr( cT ) + '))=' + _CfStr( cT )
	ENDCASE
	cRule += ' text="' + XmlEsc( cT ) + '"><formula>' + XmlEsc( cO ) + '</formula>'
CASE cO == "formula" .OR. cO == "expression"
	cT := ::oParent:ToUTF8( hb_CStr( xValue1 ) )
	cRule := StrTran( cRule, "%T%", "expression" ) + '><formula>' + XmlEsc( IIF( Left( cT, 1 ) == "=", SubStr( cT, 2 ), cT ) ) + '</formula>'
CASE cO == "duplicate"
	cRule := StrTran( cRule, "%T%", "duplicateValues" ) + '>'
CASE cO == "unique"
	cRule := StrTran( cRule, "%T%", "uniqueValues" ) + '>'
CASE cO == "blank"
	cRule := StrTran( cRule, "%T%", "containsBlanks" ) + '><formula>LEN(TRIM(' + cTL + '))=0</formula>'
CASE cO == "notblank"
	cRule := StrTran( cRule, "%T%", "notContainsBlanks" ) + '><formula>LEN(TRIM(' + cTL + '))&gt;0</formula>'
CASE cO == "aboveaverage"
	cRule := StrTran( cRule, "%T%", "aboveAverage" ) + '>'
CASE cO == "belowaverage"
	cRule := StrTran( cRule, "%T%", "aboveAverage" ) + ' aboveAverage="0">'
CASE cO == "top" .OR. cO == "bottom" .OR. cO == "top%" .OR. cO == "bottom%"
	nN := IIF( HB_ISNUMERIC( xValue1 ) .AND. xValue1 >= 1, Int( xValue1 ), 10 )
	cRule := StrTran( cRule, "%T%", "top10" ) + ' rank="' + HB_NTOS( nN ) + '"' + ;
	         IIF( "%" $ cO, ' percent="1"', '' ) + IIF( Left( cO, 1 ) == "b", ' bottom="1"', '' ) + '>'
OTHERWISE
	RETURN Self                                 // an unknown rule: nothing
ENDCASE
AAdd( ::aCond, { cSq, cRule + '</cfRule>' } )
RETURN Self

// a color scale: the smallest value in cMinColor, the largest in cMaxColor, cMidColor at the
// middle (50th percentile); the colors between are blended by Excel
METHOD ColorScale( cRange, cMinColor, cMaxColor, cMidColor ) CLASS WorkSheet
LOCAL cSq := _CfRange( cRange ), lMid := HB_ISSTRING( cMidColor ) .AND. ! Empty( cMidColor )
IF Empty( cSq )
	RETURN Self
ENDIF
AAdd( ::aCond, { cSq, '<cfRule type="colorScale" priority="%P%"><colorScale>' + ;
	'<cfvo type="min"/>' + IIF( lMid, '<cfvo type="percentile" val="50"/>', '' ) + '<cfvo type="max"/>' + ;
	'<color rgb="' + _ARGB( hb_defaultValue( cMinColor, "F8696B" ) ) + '"/>' + ;
	IIF( lMid, '<color rgb="' + _ARGB( cMidColor ) + '"/>', '' ) + ;
	'<color rgb="' + _ARGB( hb_defaultValue( cMaxColor, "63BE7B" ) ) + '"/></colorScale></cfRule>' } )
RETURN Self

// data bars: a bar in each cell, as long as its value among the values of the range
METHOD DataBar( cRange, cColor ) CLASS WorkSheet
LOCAL cSq := _CfRange( cRange )
IF Empty( cSq )
	RETURN Self
ENDIF
AAdd( ::aCond, { cSq, '<cfRule type="dataBar" priority="%P%"><dataBar><cfvo type="min"/><cfvo type="max"/>' + ;
	'<color rgb="' + _ARGB( hb_defaultValue( cColor, "638EC6" ) ) + '"/></dataBar></cfRule>' } )
RETURN Self

// AutoFit( "A:J" ), AutoFit( "C" ), AutoFit( nFrom, nTo ), AutoFit(): the columns get the
// width of their longest contents, as Excel's AutoFit. A file cannot ask Excel to do it, so
// the width is computed here: the widths of the letters of Calibri (the default font: a
// digit is 7 pixels), the number and date formats, the font size and bold of each cell,
// and Excel's formula width = Truncate( ( pixels + 5 ) / 7 * 256 ) / 256. Other fonts are
// scaled by their average width. Left out as in Excel: merged cells over more columns,
// rotated text; a formula counts with the value saved with it (Read()), else not.
// nMaxWidth: at most this wide (default 255). Call it after the cells are filled.
METHOD AutoFit( xFrom, xTo, nMaxWidth ) CLASS WorkSheet
LOCAL nC1, nC2, a, x, nRow, nCol, nR1 := 0, nC1m := 0, nR2 := 0, nC2m := 0
LOCAL hSkip := { => }, hInfo := { => }, nStyle, h, nPx, nMax, cLine
nMaxWidth := IIF( HB_ISNUMERIC( nMaxWidth ) .AND. nMaxWidth > 0, Min( nMaxWidth, 255 ), 255 )
DO CASE
CASE HB_ISSTRING( xFrom ) .AND. ":" $ xFrom
	a := hb_ATokens( xFrom, ":" )
	nC1 := _ColNum( a[ 1 ] )
	nC2 := _ColNum( a[ 2 ] )
CASE HB_ISSTRING( xFrom ) .OR. HB_ISNUMERIC( xFrom )
	nC1 := IIF( HB_ISSTRING( xFrom ), _ColNum( xFrom ), xFrom )
	nC2 := IIF( HB_ISSTRING( xTo ), _ColNum( xTo ), IIF( HB_ISNUMERIC( xTo ), xTo, nC1 ) )
OTHERWISE
	nC1 := 1
	nC2 := ::nMaxCol
ENDCASE
nC2 := Min( nC2, ::nMaxCol )
IF nC1 < 1 .OR. nC2 < nC1
	RETURN Self
ENDIF
FOR EACH x IN ::aMergeCells
	a := hb_ATokens( x, ":" )
	WORKSHEET_RC( a[ 1 ], @nR1, @nC1m )
	WORKSHEET_RC( a[ 2 ], @nR2, @nC2m )
	IF nC2m > nC1m
		FOR nRow := nR1 TO nR2
			FOR nCol := nC1m TO nC2m
				hSkip[ nRow * 16385 + nCol ] := .T.
			NEXT
		NEXT
	ENDIF
NEXT
FOR nCol := nC1 TO nC2
	nMax := 0
	FOR nRow := 1 TO LEN( ::aData )
		IF nCol > LEN( ::aData[ nRow ] ) .OR. ( x := ::aData[ nRow, nCol ] ) == NIL .OR. ;
		   hb_HHasKey( hSkip, nRow * 16385 + nCol ) .OR. ( HB_ISSTRING( x ) .AND. x == Chr( 0 ) )
			LOOP
		ENDIF
		IF HB_ISARRAY( x )                       // a formula: the value saved with it
			IF LEN( x ) < 2 .OR. x[ 2 ] == NIL
				LOOP
			ENDIF
			x := x[ 2 ]
		ENDIF
		nStyle := IIF( nRow <= LEN( ::aStyle ) .AND. HB_ISARRAY( ::aStyle[ nRow ] ) .AND. nCol <= LEN( ::aStyle[ nRow ] ), ;
		               ::aStyle[ nRow, nCol ], 0 )
		nStyle := IIF( HB_ISSTRING( nStyle ), Val( nStyle ), IIF( HB_ISNUMERIC( nStyle ), nStyle, 0 ) )   // kept as text
		IF ! hb_HHasKey( hInfo, nStyle )
			hInfo[ nStyle ] := ::oParent:StyleInfo( nStyle )
		ENDIF
		h := hInfo[ nStyle ]
		IF h[ "rotation" ] != 0
			LOOP
		ENDIF
		FOR EACH cLine IN hb_ATokens( _CellText( x, h[ "numformat" ] ), Chr( 10 ) )
			nPx := _TextPx( cLine, h[ "bold" ] ) * _FontFactor( h[ "font" ] ) * h[ "size" ] / 11
			nMax := Max( nMax, nPx )
		NEXT
	NEXT
	IF nMax > 0
		::ColumnsWidth( nCol, nCol, Min( nMaxWidth, Int( ( Round( nMax, 0 ) + 5 ) / 7 * 256 ) / 256 ) )
	ENDIF
NEXT
Return Self

// CellStyle( "B3" ) -> broj stila celije (0 = bez stila); CellStyle( "B3", nStyle )
// postavlja stil, i na praznu celiju (npr. okvir bez vrednosti)
METHOD CellStyle( uAddr, nStyle ) CLASS WorkSheet
LOCAL nRow := 0, nCol := 0, x
IF ! HB_ISARRAY( uAddr ) .AND. ! HB_ISSTRING( uAddr )
	RETURN 0
ENDIF
_CellRC( uAddr, @nRow, @nCol )
IF ! HB_ISNUMERIC( nRow ) .OR. ! HB_ISNUMERIC( nCol ) .OR. nRow < 1 .OR. nCol < 1
	RETURN 0
ENDIF
IF HB_ISNUMERIC( nStyle )
	DO WHILE LEN( ::aStyle ) < nRow
		AAdd( ::aStyle, {} )
	ENDDO
	IF LEN( ::aStyle[ nRow ] ) < nCol
		ASize( ::aStyle[ nRow ], nCol )
	ENDIF
	::aStyle[ nRow, nCol ] := IIF( nStyle > 0, HB_NTOS( nStyle ), NIL )
	RETURN nStyle
ENDIF
x := IIF( nRow <= LEN( ::aStyle ) .AND. HB_ISARRAY( ::aStyle[ nRow ] ) .AND. nCol <= LEN( ::aStyle[ nRow ] ), ::aStyle[ nRow, nCol ], NIL )
RETURN IIF( HB_ISSTRING( x ), Val( x ), 0 )

METHOD Cell( uAddr, xValue, nStyle ) CLASS WorkSheet
LOCAL l 
    IF xValue == NIL
		l:= ::ReadCell( uAddr )
	ELSE
		::WriteCell( uAddr, xValue, nStyle ) 
	ENDIF	
Return l

// Vrednost celije ili NIL. Formula se interno cuva kao { cFormula [, xVrednost] }:
// ako je u fajlu bila izracunata vrednost vraca se ona, inace "=formula".
METHOD ReadCell( uAddr ) CLASS WorkSheet
LOCAL x := _CellGet( ::aData, uAddr )
DO CASE
CASE HB_ISARRAY( x )
   RETURN IIF( LEN( x ) >= 2 .AND. x[2] != NIL, x[2], "=" + x[1] )
CASE HB_ISSTRING( x ) .AND. x == Chr(0)          // ostatak spojene celije
   RETURN NIL
ENDCASE
RETURN x

// formula celije ("=SUM(A1:A3)") ili NIL ako celija nema formulu
METHOD CellFormula( uAddr ) CLASS WorkSheet
LOCAL x := _CellGet( ::aData, uAddr )
RETURN IIF( HB_ISARRAY( x ), "=" + x[1], NIL )

// poslednji red / kolona u kojima ima podataka
METHOD LastRow() CLASS WorkSheet
RETURN ::nMaxRow

METHOD LastCol() CLASS WorkSheet
RETURN ::nMaxCol

// spojene celije, npr. { "A1:C1", "B5:B8" }
METHOD MergedCells() CLASS WorkSheet
RETURN AClone( ::aMergeCells )

// Vrednosti opsega kao niz redova: Range( "A1:C10" ), Range( "A:C" ) (sve popunjene
// redove), Range( "2:5" ) (sve popunjene kolone) ili Range( { nRed1, nKol1, nRed2, nKol2 } )
METHOD Range( xRange ) CLASS WorkSheet
LOCAL nR1 := 0, nC1 := 0, nR2 := 0, nC2 := 0, aP, aRes := {}, aRow, i, j
IF HB_ISARRAY( xRange )
   IF LEN( xRange ) >= 4
      nR1 := xRange[1] ; nC1 := xRange[2] ; nR2 := xRange[3] ; nC2 := xRange[4]
   ENDIF
ELSEIF HB_ISSTRING( xRange ) .AND. ! Empty( xRange )
   aP := hb_ATokens( StrTran( Upper( AllTrim( xRange ) ), "$", "" ), ":" )
   WORKSHEET_RC( aP[1], @nR1, @nC1 )
   IF LEN( aP ) > 1
      WORKSHEET_RC( aP[2], @nR2, @nC2 )
   ELSE
      nR2 := nR1 ; nC2 := nC1
   ENDIF
   nR1 := IIF( nR1 == 0, 1, nR1 ) ; nC1 := IIF( nC1 == 0, 1, nC1 )
   nR2 := IIF( nR2 == 0, ::nMaxRow, nR2 ) ; nC2 := IIF( nC2 == 0, ::nMaxCol, nC2 )
ENDIF
FOR i := nR1 TO nR2
   aRow := Array( Max( nC2 - nC1 + 1, 0 ) )
   FOR j := nC1 TO nC2
      aRow[ j - nC1 + 1 ] := ::ReadCell( { i, j } )
   NEXT
   AAdd( aRes, aRow )
NEXT
RETURN aRes

// ceo popunjen deo sheet-a (od A1) kao niz redova
METHOD ToArray() CLASS WorkSheet
RETURN IIF( ::nMaxRow > 0 .AND. ::nMaxCol > 0, ::Range( { 1, 1, ::nMaxRow, ::nMaxCol } ), {} )

// Formula se upisuje eksplicitno. Cell() sa tekstom koji pocinje sa "=" upisuje
// TEKST (podatak iz baze ne sme slucajno da postane formula), osim ako je
// oWorkBook:lFormulaPrefix := .T. (staro ponasanje).
METHOD Formula( uAddr, cFormula, nStyle ) CLASS WorkSheet
IF HB_ISSTRING( cFormula )
	cFormula := XmlClean( ::oParent:ToUTF8( cFormula ) )
	IF LEFT( cFormula, 1 ) == "="
		cFormula := SUBSTR( cFormula, 2 )
	ENDIF
	::WriteCell( uAddr, { cFormula }, nStyle )
ENDIF
RETURN Self

METHOD WriteCell( uAddr, xValue, nStyle ) CLASS WorkSheet
LOCAL nCol := 0, nRow := 0, lIsMerge, nRow1 := 0, nCol1 := 0, nRow2 := 0, nCol2 := 0
LOCAL i, j, k, l, adr
IF HB_ISARRAY( uAddr ) .AND. LEN( uAddr ) >= 2
	nRow := uAddr[1]
	nCol := uAddr[2]
ELSEIF HB_ISSTRING( uAddr )
	WORKSHEET_RC( StrTran( uAddr, "$", "" ), @nRow, @nCol )
ENDIF
// neispravna adresa (npr. "", "A0", NIL) ranije je rusila program ili pravila los fajl;
// Excel ima najvise 1048576 redova i 16384 kolone
IF ! HB_ISNUMERIC( nRow ) .OR. ! HB_ISNUMERIC( nCol ) .OR. nRow < 1 .OR. nCol < 1 .OR. nRow > 1048576 .OR. nCol > 16384
	RETURN NIL
ENDIF
::nMaxCol := IIF( nCol > ::nMaxCol, nCol, ::nMaxCol )
::nMaxRow := IIF( nRow > ::nMaxRow, nRow, ::nMaxRow )
IF LEN( ::aData ) < nRow
	FOR i := LEN(::aData)+1 TO nRow 
		AADD( ::aData, {} )
	NEXT I
ENDIF
IF LEN( ::aData[nRow] ) < nCol
	FOR i := LEN(::aData[nRow])+1 TO nCol
		AADD( ::aData[nRow], NIL )
	NEXT I	   
ENDIF
lIsMerge := IIF( HB_ISSTRING(::aData[nRow,nCol]) .AND. ::aData[nRow,nCol]==CHR(0),.T.,.F.)
IF HB_ISSTRING(xValue)
	// tekst iz baze (npr. CP852) -> UTF-8; kontrolni znakovi nisu dozvoljeni u XML-u
	xValue := XmlClean( ::oParent:ToUTF8( xValue ) )
	// Excel dozvoljava najvise 32767 znakova u celiji, inace odbija ceo fajl
	IF hb_UTF8Len( xValue ) > 32767
		xValue := hb_UTF8Left( xValue, 32767 )
	ENDIF
	IF ::oParent:lFormulaPrefix .AND. LEFT( xValue, 1 ) == "="
		xValue := { SUBSTR( xValue, 2 ) }
	ENDIF
ENDIF
IF !HB_ISNIL(xValue)
	::aData[nRow,nCol] := xValue
	IF HB_ISSTRING(xValue)
		::oParent:_str_total++
		// hash umesto ASCAN: ranije O(n^2) i zavisno od SET EXACT
		IF ! hb_HHasKey( ::oParent:hSharedStrings, xValue )
			AADD( ::oParent:aSharedStrings, xValue )
			::oParent:hSharedStrings[ xValue ] := LEN( ::oParent:aSharedStrings )
		ENDIF
	ENDIF
ENDIF
IF HB_ISNIL(nStyle) .AND. ValType(xValue) == "D"
	nStyle := ::oParent:NewStyle( , , , , ,::oParent:NewFormat("dd/mm/yyyy"))
ELSEIF HB_ISNIL(nStyle) .AND. ValType(xValue) == "T"     // DateTime (npr. iz Read())
	nStyle := ::oParent:NewStyle( , , , , ,::oParent:NewFormat("m/d/yyyy h:mm"))
ENDIF
IF !HB_ISNIL(nStyle) .AND. HB_ISNUMERIC(nStyle) 	
	IF LEN( ::aStyle ) < nRow
		FOR i := LEN(::aStyle)+1 TO nRow 
			AADD( ::aStyle, {} )
		NEXT I
	ENDIF
	IF LEN( ::aStyle[nRow] ) < nCol
		FOR i := LEN(::aStyle[nRow])+1 TO nCol
			AADD( ::aStyle[nRow], NIL )
		NEXT I	   
	ENDIF
	::aStyle[nRow,nCol] := HB_NTOS(nStyle)
	IF lIsMerge
		FOR l:=1 TO LEN( ::aMergeCells )
			adr := hb_ATokens( ::aMergeCells[l], ":" )
			WORKSHEET_RC( adr[1], @nRow1, @nCol1 )
		    IF nRow1==nRow .AND. nCol1==nCol
				WORKSHEET_RC( adr[2], @nRow2, @nCol2 )
				FOR I:=nRow1 TO nRow2
					FOR J:=nCol1 TO nCol2
						IF LEN( ::aStyle ) < I
							FOR k := LEN(::aStyle)+1 TO I 
								AADD( ::aStyle, {} )
							NEXT k
						ENDIF
						IF LEN( ::aStyle[I] ) < J
							FOR k := LEN(::aStyle[I])+1 TO J
								AADD( ::aStyle[I], NIL )
							NEXT k	   
						ENDIF
						::aStyle[I,J] := HB_NTOS(nStyle)						
					NEXT J
				NEXT I
				EXIT
			ENDIF
		NEXT l
	ENDIF
ENDIF
Return ::aData[nRow,nCol]

METHOD WriteWorksheet(n, cPath) CLASS WorkSheet
LOCAL i, j, c, x, handle, cS, cSA, cLastStyle, cRowAttr, cCells, nIdx, aRow, aSty, aDet, nRows
LOCAL hW, aK, nFrom, nTo, nW, a
LOCAL nl := chr(10)
handle := FCreate( cPath+"sheet"+HB_NTOS(n)+".xml" )
FWrite( handle, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+nl )
FWrite( handle, '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac xr xr2 xr3" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" xr:uid="{00000000-0001-0000-0000-000000000000}">'+nl )
FWrite( handle, '<sheetPr><pageSetUpPr fitToPage="'+iif(::fitToPage,"1","0")+'"/></sheetPr>'+nl )
// ColumnIndexToColumnLetter: ranije CHR() pa je kolona posle Z davala "["
FWrite( handle, '<dimension ref="A1'+IIF( ::nMaxRow > 0 .AND. ::nMaxCol > 0, ':'+ColumnIndexToColumnLetter(::nMaxCol)+HB_NTOS(::nMaxRow), '' )+'"/>'+nl )
FWrite( handle, '<sheetFormatPr defaultRowHeight="15"/>'+nl )
// Sirine kolona: Excel trazi <col> rastuce i bez preklapanja, a ColumnsWidth() moze
// da se zove bilo kojim redom i za opsege koji se preklapaju (vazi poslednji poziv)
IF LEN( ::aCols ) > 0
	hW := { => }
	FOR EACH a IN ::aCols
		FOR j := Max( a[1], 1 ) TO Min( a[2], 16384 )
			// sirina + skrivena kolona; kljuc za poredjenje susednih kolona
			hW[ j ] := HB_NTOC( a[3], 8 ) + IIF( LEN( a ) >= 4 .AND. a[4], '" hidden="1', '' )
		NEXT
	NEXT
	aK := ASort( hb_HKeys( hW ) )
	IF LEN( aK ) > 0
		FWrite( handle, '<cols>'+nl )
		nFrom := aK[ 1 ]
		nTo := nFrom
		nW := hW[ nFrom ]
		FOR i := 2 TO LEN( aK ) + 1
			IF i <= LEN( aK ) .AND. aK[ i ] == nTo + 1 .AND. hW[ aK[ i ] ] == nW
				nTo := aK[ i ]
			ELSE
				FWrite( handle, '<col min="'+HB_NTOS(nFrom)+'" max="'+HB_NTOS(nTo)+'" width="'+nW+'" bestFit="1" customWidth="1"/>'+nl )
				IF i <= LEN( aK )
					nFrom := aK[ i ]
					nTo := nFrom
					nW := hW[ nFrom ]
				ENDIF
			ENDIF
		NEXT
		FWrite( handle, '</cols>'+nl )
	ENDIF
ENDIF
FWrite( handle, '<sheetData>'+nl )
// Red se pise ako ima podataka ili podesavanja iz RowDetail() (ranije je visina
// praznog reda bila izgubljena, a stil i skrivanje reda se nisu ni upisivali)
nRows := Max( Max( LEN( ::aData ), LEN( ::aRowsDetail ) ), LEN( ::aStyle ) )
FOR i := 1 TO nRows
	aRow := IIF( i <= LEN( ::aData ), ::aData[ i ], {} )
	aSty := IIF( i <= LEN( ::aStyle ) .AND. HB_ISARRAY( ::aStyle[ i ] ), ::aStyle[ i ], {} )
	aDet := IIF( i <= LEN( ::aRowsDetail ), ::aRowsDetail[ i ], NIL )
	cRowAttr := ""
	IF aDet != NIL
		IF aDet[ 1 ] > 0
			cRowAttr += ' ht="'+HB_NTOS( aDet[ 1 ] )+'" customHeight="1"'
		ENDIF
		IF aDet[ 3 ]
			cRowAttr += ' hidden="1"'
		ENDIF
		IF aDet[ 2 ] > 0
			cRowAttr += ' s="'+HB_NTOS( aDet[ 2 ] )+'" customFormat="1"'
		ENDIF
	ENDIF
	cCells := ""
	cLastStyle := ""
	FOR j := 1 TO Max( LEN( aRow ), LEN( aSty ) )
		x := IIF( j <= LEN( aRow ), aRow[ j ], NIL )
		cS := IIF( j <= LEN( aSty ) .AND. HB_ISSTRING( aSty[ j ] ), aSty[ j ], "" )
		IF x == NIL
			// celija bez vrednosti, ali sa stilom (okvir, pozadina - npr. iz Read() ili CellStyle())
			IF ! Empty( cS )
				cCells += '<c r="'+ColumnIndexToColumnLetter( j )+HB_NTOS( i )+'" s="'+cS+'"/>'+nl
			ENDIF
			LOOP
		ENDIF
		IF ! Empty( cS )
			cLastStyle := cS
		ENDIF
		cSA := IIF( Empty( cS ), "", ' s="'+cS+'"' )
		c := ColumnIndexToColumnLetter( j ) + HB_NTOS( i )
		DO CASE
		CASE HB_ISNUMERIC( x )
			cCells += '<c r="'+c+'"'+cSA+'><v>'+HB_NTOC( x )+'</v></c>'+nl
		CASE HB_ISARRAY( x )
			// formula, interno { cFormula [, xVrednost] }
			// tip rezultata: tekst samo kad je poznata tekstualna vrednost, inace broj
			// (t="str" bez vrednosti je Calc-u i zbir prikazivao kao tekst)
			cCells += '<c r="'+c+'"'+IIF( LEN( x ) >= 2 .AND. HB_ISSTRING( x[ 2 ] ), ' t="str"', '' )+cSA+'><f>'+XmlEsc( x[ 1 ] )+'</f>' + ;
			          IIF( LEN( x ) >= 2 .AND. HB_ISNUMERIC( x[ 2 ] ), '<v>'+HB_NTOC( x[ 2 ] )+'</v>', ;
			          IIF( LEN( x ) >= 2 .AND. HB_ISSTRING( x[ 2 ] ), '<v>'+XmlEsc( x[ 2 ] )+'</v>', '' ) )+'</c>'+nl
		CASE HB_ISSTRING( x )
			IF LEFT( x, 1 ) == CHR(0)
				// ostatak spojene celije: stil celije ili poslednji stil u redu
				IF ! Empty( cLastStyle )
					cCells += '<c r="'+c+'" s="'+cLastStyle+'"/>'+nl
				ENDIF
			ELSE
				nIdx := hb_HGetDef( ::oParent:hSharedStrings, x, 0 )
				IF nIdx > 0
					cCells += '<c r="'+c+'" t="s"'+cSA+'><v>'+HB_NTOS(nIdx-1)+'</v></c>'+nl
				ELSE
					// tekst koji nije prosao kroz WriteCell (npr. ucitan inlineStr)
					cCells += '<c r="'+c+'" t="inlineStr"'+cSA+'><is><t xml:space="preserve">'+XmlEsc( x )+'</t></is></c>'+nl
				ENDIF
			ENDIF
		CASE HB_ISLOGICAL( x )
			cCells += '<c r="'+c+'" t="b"'+cSA+'><v>'+IIF( x, "1", "0" )+'</v></c>'+nl
		CASE ValType( x ) == "D" .OR. ValType( x ) == "T"
			// prazan datum = prazna celija (ranije negativan broj, Excel prikaze ####)
			IF EMPTY( x )
				cCells += '<c r="'+c+'"'+cSA+'/>'+nl
			ELSE
				cCells += '<c r="'+c+'"'+cSA+'><v>'+_DateToSerial( x )+'</v></c>'+nl
			ENDIF
		ENDCASE
	NEXT
	IF ! Empty( cCells ) .OR. ! Empty( cRowAttr )
		FWrite( handle, '<row r="'+HB_NTOS(i)+'"'+cRowAttr+'>'+nl+cCells+'</row>'+nl )
	ENDIF
NEXT
FWrite( handle, '</sheetData>'+chr(10) )
IF LEN( ::aMergeCells ) > 0
	FWrite( handle, '<mergeCells>'+chr(10) )
	FOR I:=1 TO LEN( ::aMergeCells )
		FWrite( handle, '<mergeCell ref="'+::aMergeCells[I]+'"/>'+chr(10) )
	NEXT I
	FWrite( handle, '</mergeCells>'+chr(10) )
ENDIF
FOR I := 1 TO LEN( ::aCond )                    // priority: in the order they were added
	FWrite( handle, '<conditionalFormatting sqref="'+::aCond[I][1]+'">'+StrTran( ::aCond[I][2], "%P%", HB_NTOS(I) )+'</conditionalFormatting>'+chr(10) )
NEXT I
FWrite( handle, '<pageMargins left="'+HB_NTOC(::leftMargin)+'" right="'+HB_NTOC(::rightMargin)+'" top="'+HB_NTOC(::topMargin,2)+'" bottom="'+HB_NTOC(::bottomMargin,2)+'" header="'+HB_NTOC(::headerMargin,2)+'" footer="'+HB_NTOC(::footerMargin,2)+'"/>'+chr(10) )
FWrite( handle, '<pageSetup paperSize="'+HB_NTOC(::PaperSize,0)+'" horizontalDpi="'+HB_NTOC(::horizontalDpi,2)+'" verticalDpi="'+HB_NTOC(::verticalDpi,2)+'"'+iif(::lLandscape,' orientation="landscape"',' orientation="portrait"')+' fitToWidth="'+iif(::fitToPage,"1","0")+'" fitToHeight="0"/>'+chr(10) )
IF ::lHeader .OR. ::lFooter
	FWrite( handle, '<headerFooter>' )
	IF ::lHeader
		FWrite( handle, '<oddHeader>' )
		IF LEN(::aHeader[1])>0
			FWrite( handle, '&amp;L'+XmlEsc(::oParent:ToUTF8(_HFEsc(::aHeader[1]))) )
		ENDIF
		IF LEN(::aHeader[2])>0
			FWrite( handle, '&amp;C'+XmlEsc(::oParent:ToUTF8(_HFEsc(::aHeader[2]))) )
		ENDIF
		IF LEN(::aHeader[3])>0
			FWrite( handle, '&amp;R'+XmlEsc(::oParent:ToUTF8(_HFEsc(::aHeader[3]))) )
		ENDIF		
		FWrite( handle, '</oddHeader>' )
	ENDIF
	IF ::lFooter
		FWrite( handle, '<oddFooter>' )
		IF LEN(::aFooter[1])>0
			FWrite( handle, '&amp;L'+XmlEsc(::oParent:ToUTF8(_HFEsc(::aFooter[1]))) )
		ENDIF
		IF LEN(::aFooter[2])>0
			FWrite( handle, '&amp;C'+XmlEsc(::oParent:ToUTF8(_HFEsc(::aFooter[2]))) )
		ENDIF
		IF LEN(::aFooter[3])>0
			FWrite( handle, '&amp;R'+XmlEsc(::oParent:ToUTF8(_HFEsc(::aFooter[3]))) )
		ENDIF		
		FWrite( handle, '</oddFooter>' )
	ENDIF	
	FWrite( handle, '</headerFooter>'+chr(10) )
ENDIF
IF LEN( ::aDrawingsObjects ) > 0 .OR. LEN( ::aCharts ) > 0
	// jedan <drawing> po sheet-u; slike i grafikoni su u istom drawing<n>.xml
	FWrite( handle, '<drawing r:id="rIdDr1"/>'+chr(10) )
	::WriteDrawingPart( n )
ENDIF
FWrite( handle, '</worksheet>' )   
FClose( handle )         
Return Self

// Ucitava jedan sheet (poziva ga WorkBook:Read). aSST su deljeni stringovi fajla,
// aKinds za svaki stil fajla kaze da li je format datum (1), datum+vreme (2) ili
// vreme (3), a aXfMap daje odgovarajuci stil ovog workbook-a (0 = bez stila).
// Pored vrednosti cita stilove celija (i praznih), sirine kolona, visine redova,
// margine, papir, orijentaciju i zaglavlje/podnozje, pa Read() + Save() cuva izgled.
METHOD ReadWorkSheet( cFileName, aSST, aKinds, aXfMap ) CLASS WorkSheet
LOCAL cTok, cA, cTag, nKind, cText, h, a, oWB := ::oParent
LOCAL nRow := 0, nCol := 0, nR, nXf := 1, nSt := 0, nRSt, cR, cType := "", cV, cF, cIs, xVal, i
LOCAL lIs := .F., lPh := .F., nDateSt, nTimeSt
   IF ! HB_ISARRAY( aSST )
      aSST := oWB:aSharedStrings
   ENDIF
   IF ! HB_ISARRAY( aKinds )
      aKinds := {}
   ENDIF
   IF ! HB_ISARRAY( aXfMap )
      aXfMap := {}
   ENDIF
   FOR EACH cTok IN _XTokens( cFileName )
      cA := _XTok( cTok, @cTag, @nKind, @cText )
      DO CASE
      CASE nKind == 0
      CASE cTag == "v"
         IF nKind == 1
            // prazno <v/> ili <v></v> (npr. openpyxl za formulu) = nema vrednosti
            cV := IIF( cText == "", NIL, _XEnt( cText ) )
         ENDIF
      CASE cTag == "c"
         IF nKind != 2                         // <c ...> ili <c .../>
            h := _XAttrs( cA )
            cR := hb_HGetDef( h, "r", "" )
            IF Empty( cR )                     // adresa nije obavezna
               nCol++
            ELSE
               nR := 0
               WORKSHEET_RC( cR, @nR, @nCol )
               IF nR > 0
                  nRow := nR
               ENDIF
            ENDIF
            cType := hb_HGetDef( h, "t", "" )
            nXf := Val( hb_HGetDef( h, "s", "0" ) ) + 1
            nSt := IIF( nXf <= LEN( aXfMap ), aXfMap[ nXf ], 0 )
            cV := NIL
            cF := NIL
            cIs := NIL
         ENDIF
         IF nKind != 1 .AND. nRow >= 1 .AND. nCol >= 1   // </c> ili <c .../>: upis
            xVal := NIL
            DO CASE
            CASE cType == "inlineStr"
               xVal := IIF( cIs == NIL, NIL, _XDecode( cIs ) )
            CASE cV == NIL
               // prazna celija (samo stil) ili formula bez izracunate vrednosti
            CASE cType == "s"
               i := Val( cV ) + 1
               xVal := IIF( i >= 1 .AND. i <= LEN( aSST ), aSST[ i ], "" )
            CASE cType == "b"
               xVal := ( Val( cV ) != 0 )
            CASE cType == "str" .OR. cType == "e"  // tekst formule / greska (#DIV/0!)
               xVal := _XDecode( cV )
            CASE cType == "d"                      // ISO 8601 datum
               xVal := _IsoToDate( cV )
            OTHERWISE
               xVal := _XVal( cV )
               IF oWB:lReadDates .AND. nXf <= LEN( aKinds ) .AND. aKinds[ nXf ] > 0
                  xVal := _SerialToDate( xVal, aKinds[ nXf ], oWB:lDate1904 )
               ENDIF
            ENDCASE
            IF HB_ISSTRING( cF ) .AND. ! Empty( cF )
               xVal := { _XDecode( cF ), xVal }
            ENDIF
            IF xVal != NIL
               DO WHILE LEN( ::aData ) < nRow
                  AAdd( ::aData, {} )
               ENDDO
               IF LEN( ::aData[ nRow ] ) < nCol
                  ASize( ::aData[ nRow ], nCol )
               ENDIF
               ::aData[ nRow, nCol ] := xVal
               ::nMaxRow := Max( ::nMaxRow, nRow )
               ::nMaxCol := Max( ::nMaxCol, nCol )
               // datum bez stila dobije format datuma, da posle Save() ne postane broj
               IF nSt == 0 .AND. ( ValType( xVal ) == "D" .OR. ValType( xVal ) == "T" )
                  IF nDateSt == NIL
                     nDateSt := oWB:NewStyle( , , , , , oWB:NewFormat( "dd/mm/yyyy" ) )
                     nTimeSt := oWB:NewStyle( , , , , , oWB:NewFormat( "m/d/yyyy h:mm" ) )
                  ENDIF
                  nSt := IIF( ValType( xVal ) == "D", nDateSt, nTimeSt )
               ENDIF
            ENDIF
            IF nSt > 0
               ::CellStyle( { nRow, nCol }, nSt )
            ENDIF
         ENDIF
      CASE cTag == "f"
         IF nKind == 1
            cF := _XEnt( cText )
         ENDIF
      CASE cTag == "is"
         lIs := ( nKind == 1 )
         IF lIs
            cIs := ""
         ENDIF
      CASE cTag == "rPh"
         lPh := ( nKind == 1 )
      CASE cTag == "t"
         IF nKind == 1 .AND. lIs .AND. ! lPh
            cIs += _XEnt( cText )
         ENDIF
      CASE cTag == "row"
         IF nKind != 2
            h := _XAttrs( cA )
            cR := hb_HGetDef( h, "r", "" )
            nRow := IIF( Empty( cR ), nRow + 1, Val( cR ) )
            nCol := 0
            // visina, skriven red, stil reda
            IF _XBool( h, "customHeight" ) .OR. _XBool( h, "hidden" ) .OR. _XBool( h, "customFormat" )
               nRSt := 0
               IF _XBool( h, "customFormat" )
                  i := Val( hb_HGetDef( h, "s", "0" ) ) + 1
                  nRSt := IIF( i <= LEN( aXfMap ), aXfMap[ i ], 0 )
               ENDIF
               ::RowDetail( nRow, IIF( _XBool( h, "customHeight" ), Val( hb_HGetDef( h, "ht", "0" ) ), 0 ), ;
                            nRSt, _XBool( h, "hidden" ) )
            ENDIF
         ENDIF
      CASE cTag == "col"
         IF nKind != 2
            h := _XAttrs( cA )
            IF hb_HHasKey( h, "width" ) .OR. _XBool( h, "hidden" )
               ::ColumnsWidth( Val( hb_HGetDef( h, "min", "0" ) ), Val( hb_HGetDef( h, "max", "0" ) ), ;
                               Val( hb_HGetDef( h, "width", "8.43" ) ), _XBool( h, "hidden" ) )
            ENDIF
         ENDIF
      CASE cTag == "mergeCell"
         cR := hb_HGetDef( _XAttrs( cA ), "ref", "" )
         IF ! Empty( cR ) .AND. _AScanX( ::aMergeCells, cR ) == 0
            AAdd( ::aMergeCells, cR )
         ENDIF
      CASE cTag == "pageSetUpPr"
         ::fitToPage := _XBool( _XAttrs( cA ), "fitToPage" )
      CASE cTag == "pageMargins"
         h := _XAttrs( cA )
         ::leftMargin   := Val( hb_HGetDef( h, "left",   _S( ::leftMargin ) ) )
         ::rightMargin  := Val( hb_HGetDef( h, "right",  _S( ::rightMargin ) ) )
         ::topMargin    := Val( hb_HGetDef( h, "top",    _S( ::topMargin ) ) )
         ::bottomMargin := Val( hb_HGetDef( h, "bottom", _S( ::bottomMargin ) ) )
         ::headerMargin := Val( hb_HGetDef( h, "header", _S( ::headerMargin ) ) )
         ::footerMargin := Val( hb_HGetDef( h, "footer", _S( ::footerMargin ) ) )
      CASE cTag == "pageSetup"
         h := _XAttrs( cA )
         IF hb_HHasKey( h, "paperSize" )
            ::paperSize := Val( h[ "paperSize" ] )
         ENDIF
         ::lLandscape := ( hb_HGetDef( h, "orientation", "" ) == "landscape" )
      CASE cTag == "oddHeader" .AND. nKind == 1
         a := _HFParts( _XEnt( cText ) )
         ::AddHeader( a[1], a[2], a[3] )
      CASE cTag == "oddFooter" .AND. nKind == 1
         a := _HFParts( _XEnt( cText ) )
         ::AddFooter( a[1], a[2], a[3] )
      ENDCASE
   NEXT
RETURN Self
	
STATIC Function ColumnIndexToColumnLetter(colIndex)
LOCAL div := colIndex
LOCAL colLetter := ""
LOCAL modnum
While div > 0
    modnum := MOD((div - 1),26)
    colLetter := Chr(65 + modnum) + colLetter
    div := Int((div - modnum) / 26)
End 
Return colLetter

// Zipuje ceo cDir u cZipName; imena u zipu su relativna u odnosu na cDir.
// Radi sa apsolutnim putanjama (bez DirChange) i vraca .T. ako je sve uspelo.
STATIC FUNCTION MyZip( cZipName, cDir )
LOCAL cSep := hb_ps(), hZip, aFile, lOk := .F.
LOCAL cRels := "_rels" + cSep + ".rels"
hZip := hb_zipOpen( cZipName )
IF ! Empty( hZip )
	lOk := .T.
	FOR EACH aFile IN hb_DirScan( cDir, "*.*" )
		IF !( aFile[ 1 ] == cRels )          // .rels ide poslednji, kao i ranije
			lOk := ( hb_zipStoreFile( hZip, cDir + cSep + aFile[ 1 ], aFile[ 1 ],,, .T. ) == 0 ) .AND. lOk
		ENDIF
	NEXT
	lOk := ( hb_zipStoreFile( hZip, cDir + cSep + cRels, cRels,,, .T. ) == 0 ) .AND. lOk
	lOk := ( hb_zipClose( hZip ) == 0 ) .AND. lOk
ENDIF
RETURN lOk

// Escape za XML tekst/atribut (izbacuje i nedozvoljene kontrolne znakove).
// Vrednosti celija se cuvaju neizmenjene, escape se radi tek pri upisu.
// '&' mora prvi, inace bi se dupli-escape-ovali &lt; &gt; &quot;.
STATIC FUNCTION XmlEsc( c )
c := XmlClean( c )
c := StrTran( c, "&", "&amp;" )
c := StrTran( c, "<", "&lt;" )
c := StrTran( c, ">", "&gt;" )
c := StrTran( c, '"', "&quot;" )
RETURN c

// XML 1.0 ne dozvoljava znakove 0-31 osim TAB, LF i CR. DBF polje sa ostatkom
// chr(0)/chr(26) posle uvoza inace pokvari ceo fajl, isto kao CP852 kvacica.
STATIC FUNCTION XmlClean( c )
STATIC s_cCtrl
LOCAL i
IF s_cCtrl == NIL
	s_cCtrl := ""
	FOR i := 0 TO 31
		IF i != 9 .AND. i != 10 .AND. i != 13
			s_cCtrl += Chr( i )
		ENDIF
	NEXT
ENDIF
RETURN hb_StrReplace( c, s_cCtrl, "" )

// Raspakuje ceo zip u cPath. Folderi se prave rekurzivno (hb_DirBuild): ranije
// HB_DirCreate() nije napravio npr. xl\worksheets ako u zipu pre njega nije bio
// fajl iz xl\ (redosled kod openpyxl-a i drugih), pa se sheet-ovi nisu ucitali.
STATIC FUNCTION my_UnZipFile( cFile, cPath )
LOCAL aFiles := hb_GetFilesInZip( cFile ), cName, cDir
IF ! HB_ISARRAY( aFiles ) .OR. Empty( aFiles )
   RETURN .F.
ENDIF
FOR EACH cName IN aFiles
   cDir := HB_FNameDir( HB_DirSepToOS( cName ) )
   IF ! Empty( cDir )
      hb_DirBuild( HB_DirSepAdd( cPath ) + cDir )
   ENDIF
NEXT
RETURN HB_UnZipFile( cFile, NIL, .T., NIL, cPath, aFiles )

STATIC PROCEDURE _CopyImageToMedia( cSrcPath, cTempDir, cSep, cID )
LOCAL cExt := lower( hb_FNameExt( cSrcPath ) )
IF Empty( cExt ); cExt := ""; ENDIF
IF cExt == ".jpg"; cExt := ".jpeg"; ENDIF
__CopyFile( cSrcPath, cTempDir + cSep + "xl" + cSep + "media" + cSep + "image" + cID + lower( cExt ) )
RETURN

STATIC FUNCTION FilePath( cFile )
LOCAL nPos, cFilePath := iif((nPos := RAT(hb_PS(), cFile)) != 0,SUBSTR(cFile, 1, nPos),"")
RETURN cFilePath

STATIC FUNCTION _MimeForImageExt( cExtNoDot )
   LOCAL c := Lower( cExtNoDot )
   DO CASE
   CASE c == "jpeg" .OR. c == "jpg"
      RETURN "image/jpeg"
   CASE c == "png"
      RETURN "image/png"
   CASE c == "gif"
      RETURN "image/gif"
   CASE c == "bmp"
      RETURN "image/bmp"
   CASE c == "emf"
      RETURN "image/x-emf"
   CASE c == "wmf"
      RETURN "image/x-wmf"
   CASE c == "svg"
      RETURN "image/svg+xml"
   CASE c == "tiff" .OR. c == "tif"
      RETURN "image/tiff"
   OTHERWISE
      RETURN "application/octet-stream"
   ENDCASE
RETURN NIL

// ---------------------------------------------------------------------------
// Grafikon (DrawingML chart). Pravi ga WorkSheet:AddChart(), upisuje Save().
// Boje su eksplicitne (aColors): workbook nema temu (theme1.xml), pa bi bez
// njih LibreOffice crtao nevidljive stubice/delove pite, a Excel temine boje.
// ---------------------------------------------------------------------------
CLASS ChartML
   DATA oWorkSheet
   DATA ID                 // redni broj u workbook-u -> xl/charts/chart<ID>.xml
   DATA cTitle
   DATA cType              // "bar" uspravni stubici, "hbar" vodoravni, "line", "pie"
   DATA cCatRef            // kategorije, npr. "A2:A13"
   DATA aValRefs           // serije, npr. { "B2:B13", "C2:C13" }
   DATA aNames             // imena serija (opciono)
   DATA aColors            // RGB boje serija (za pitu: boje delova)
   DATA nCol1, nRow1, nCol2, nRow2   // sidro, 0-based kao kod slika
   DATA lLegend INIT .T.
   DATA lCatReverse        // kategorije odozgo nadole (podrazumevano za "hbar",
                           // da redosled bude kao u tabeli: prvi red na vrhu)

   METHOD New( oWS, nID, cTitle, cType, cCatRef, xValRef, uFrom, uTo, aNames )
   METHOD WritePart( cDir )      // upise xl/charts/chart<ID>.xml
   METHOD AnchorXML()            // xdr:twoCellAnchor za drawing<n>.xml
   METHOD SerXML( n, cCatF, aCat, oCat )
   METHOD Color( n )
ENDCLASS

METHOD New( oWS, nID, cTitle, cType, cCatRef, xValRef, uFrom, uTo, aNames ) CLASS ChartML
LOCAL nRow := 0, nCol := 0
   ::oWorkSheet := oWS
   ::ID         := nID
   ::cTitle     := IIF( HB_ISSTRING( cTitle ), cTitle, "" )
   ::cType      := IIF( HB_ISSTRING( cType ) .AND. ! Empty( cType ), Lower( AllTrim( cType ) ), "bar" )
   IF AScan( { "bar", "hbar", "line", "pie" }, {| x | x == ::cType } ) == 0
      ::cType := "bar"
   ENDIF
   ::cCatRef    := cCatRef
   ::aValRefs   := IIF( HB_ISARRAY( xValRef ), xValRef, { xValRef } )
   ::aNames     := IIF( HB_ISARRAY( aNames ), aNames, {} )
   ::lCatReverse := ( ::cType == "hbar" )
   // Office paleta
   ::aColors    := { "4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47", ;
                     "264478", "9E480E", "636363", "997300", "255E91", "43682B" }
   // od gornjeg levog ugla celije uFrom do donjeg desnog ugla celije uTo
   _CellRC( uFrom, @nRow, @nCol )
   ::nRow1 := nRow - 1
   ::nCol1 := nCol - 1
   _CellRC( uTo, @nRow, @nCol )
   ::nRow2 := nRow
   ::nCol2 := nCol
RETURN Self

METHOD Color( n ) CLASS ChartML
RETURN ::aColors[ ( n - 1 ) % LEN( ::aColors ) + 1 ]

METHOD WritePart( cDir ) CLASS ChartML
LOCAL c := "", nl := Chr(10), i, h, aCat := {}, cCatF := NIL, oCat := NIL
LOCAL ax1 := HB_NTOS( 50000000 + ::ID * 10 + 1 ), ax2 := HB_NTOS( 50000000 + ::ID * 10 + 2 )
LOCAL lH := ( ::cType == "hbar" )
LOCAL cLnAx := '<c:spPr><a:ln w="9525">'+_Fill( "BFBFBF" )+'</a:ln></c:spPr>'
   IF HB_ISSTRING( ::cCatRef ) .AND. ! Empty( ::cCatRef )
      cCatF := _ChartRef( ::cCatRef, ::oWorkSheet, aCat, @oCat )
   ENDIF
   c += '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+nl
   c += '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
   c += 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
   c += 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'+nl
   c += '<c:roundedCorners val="0"/>'+nl
   c += '<c:chart>'+nl
   IF ! Empty( ::cTitle )
      c += '<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" b="1"/></a:pPr><a:r><a:rPr lang="sr-Latn-RS" sz="1400" b="1"/><a:t>'
      c += XmlEsc( ::oWorkSheet:oParent:ToUTF8( ::cTitle ) )
      c += '</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>'+nl
      c += '<c:autoTitleDeleted val="0"/>'+nl
   ELSE
      c += '<c:autoTitleDeleted val="1"/>'+nl
   ENDIF
   c += '<c:plotArea><c:layout/>'+nl
   // redosled elemenata prati OOXML semu (Excel odbija drugaciji redosled)
   DO CASE
   CASE ::cType == "pie"
      c += '<c:pieChart><c:varyColors val="1"/>'+nl
      c += ::SerXML( 1, cCatF, aCat, oCat )    // pita prikazuje samo prvu seriju
      c += '<c:firstSliceAng val="0"/></c:pieChart>'+nl
   CASE ::cType == "line"
      c += '<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>'+nl
      FOR i := 1 TO LEN( ::aValRefs )
         c += ::SerXML( i, cCatF, aCat, oCat )
      NEXT
      c += '<c:marker val="1"/><c:axId val="'+ax1+'"/><c:axId val="'+ax2+'"/></c:lineChart>'+nl
   OTHERWISE
      c += '<c:barChart><c:barDir val="'+IIF( lH, "bar", "col" )+'"/><c:grouping val="clustered"/><c:varyColors val="0"/>'+nl
      FOR i := 1 TO LEN( ::aValRefs )
         c += ::SerXML( i, cCatF, aCat, oCat )
      NEXT
      c += '<c:gapWidth val="150"/><c:axId val="'+ax1+'"/><c:axId val="'+ax2+'"/></c:barChart>'+nl
   ENDCASE
   IF !( ::cType == "pie" )
      c += '<c:catAx><c:axId val="'+ax1+'"/><c:scaling><c:orientation val="'+IIF( ::lCatReverse, "maxMin", "minMax" )+'"/></c:scaling><c:delete val="0"/>'
      c += '<c:axPos val="'+IIF( lH, "l", "b" )+'"/><c:numFmt formatCode="General" sourceLinked="1"/>'
      // "low": nazivi kategorija uz ivicu, ne preko negativnih vrednosti
      c += '<c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="low"/>'+cLnAx
      c += '<c:crossAx val="'+ax2+'"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/>'
      c += '<c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>'+nl
      c += '<c:valAx><c:axId val="'+ax2+'"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>'
      c += '<c:axPos val="'+IIF( lH, "b", "l" )+'"/>'
      c += '<c:majorGridlines><c:spPr><a:ln w="9525">'+_Fill( "D9D9D9" )+'</a:ln></c:spPr></c:majorGridlines>'
      c += '<c:numFmt formatCode="#,##0" sourceLinked="0"/>'
      c += '<c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'+cLnAx
      c += '<c:crossAx val="'+ax1+'"/><c:crosses val="'+IIF( ::lCatReverse, "max", "autoZero" )+'"/><c:crossBetween val="between"/></c:valAx>'+nl
   ENDIF
   c += '</c:plotArea>'+nl
   IF ::lLegend
      c += '<c:legend><c:legendPos val="'+IIF( ::cType == "pie", "r", "b" )+'"/><c:overlay val="0"/></c:legend>'+nl
   ENDIF
   c += '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/>'+nl
   c += '</c:chart>'+nl
   c += '<c:spPr>'+_Fill( "FFFFFF" )+'<a:ln w="9525">'+_Fill( "D9D9D9" )+'</a:ln></c:spPr>'+nl
   c += '</c:chartSpace>'
   h := FCreate( cDir + "chart" + HB_NTOS( ::ID ) + ".xml" )
   FWrite( h, c )
   FClose( h )
RETURN Self

// Jedna serija. Ime: aNames[n], inace tekst u celiji iznad (ili levo od) opsega.
METHOD SerXML( n, cCatF, aCat, oCat ) CLASS ChartML
LOCAL c := "", aVal := {}, oVal := NIL, cValF, cName := NIL, x, i, nl := Chr(10)
LOCAL oWB := ::oWorkSheet:oParent, cCol := ::Color( n ), cWhite := '<a:ln w="12700">'+_Fill( "FFFFFF" )+'</a:ln>'
   cValF := _ChartRef( ::aValRefs[n], ::oWorkSheet, aVal, @oVal )
   IF LEN( ::aNames ) >= n .AND. HB_ISSTRING( ::aNames[n] )
      cName := oWB:ToUTF8( ::aNames[n] )
   ELSEIF LEN( aVal ) > 0 .AND. oVal != NIL
      IF LEN( aVal ) > 1 .AND. aVal[1][1] == aVal[2][1]           // vodoravan opseg: naziv levo
         x := IIF( aVal[1][2] > 1, oVal:ReadCell( { aVal[1][1], aVal[1][2] - 1 } ), NIL )
      ELSE                                                         // uspravan opseg: naziv iznad
         x := IIF( aVal[1][1] > 1, oVal:ReadCell( { aVal[1][1] - 1, aVal[1][2] } ), NIL )
      ENDIF
      IF HB_ISSTRING( x ) .AND. ! Empty( x ) .AND. !( Left( x, 1 ) $ "=" + Chr(0) )
         cName := x
      ENDIF
   ENDIF
   IF cName == NIL
      cName := "Serija " + HB_NTOS( n )
   ENDIF
   c += '<c:ser><c:idx val="'+HB_NTOS( n - 1 )+'"/><c:order val="'+HB_NTOS( n - 1 )+'"/>'
   c += '<c:tx><c:v>'+XmlEsc( cName )+'</c:v></c:tx>'
   DO CASE
   CASE ::cType == "line"
      c += '<c:spPr><a:ln w="28575" cap="rnd">'+_Fill( cCol )+'<a:round/></a:ln></c:spPr>'
      c += '<c:marker><c:symbol val="circle"/><c:size val="5"/><c:spPr>'+_Fill( cCol )+'<a:ln w="9525">'+_Fill( cCol )+'</a:ln></c:spPr></c:marker>'
   CASE ::cType == "pie"
      c += '<c:spPr>'+cWhite+'</c:spPr>'
      FOR i := 1 TO MAX( LEN( aVal ), LEN( aCat ) )                 // boja svakog dela
         c += '<c:dPt><c:idx val="'+HB_NTOS( i - 1 )+'"/><c:bubble3D val="0"/><c:spPr>'+_Fill( ::Color( i ) )+cWhite+'</c:spPr></c:dPt>'
      NEXT
      c += '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/>'
      c += '<c:showPercent val="1"/><c:showBubbleSize val="0"/><c:showLeaderLines val="1"/></c:dLbls>'
   OTHERWISE
      c += '<c:spPr>'+_Fill( cCol )+'</c:spPr><c:invertIfNegative val="0"/>'
   ENDCASE
   IF cCatF != NIL
      c += '<c:cat><c:strRef><c:f>'+XmlEsc( cCatF )+'</c:f>'+_ChartCache( oCat, aCat, .F. )+'</c:strRef></c:cat>'
   ENDIF
   c += '<c:val><c:numRef><c:f>'+XmlEsc( cValF )+'</c:f>'+_ChartCache( oVal, aVal, .T. )+'</c:numRef></c:val>'
   IF ::cType == "line"
      c += '<c:smooth val="0"/>'
   ENDIF
   c += '</c:ser>'+nl
RETURN c

METHOD AnchorXML() CLASS ChartML
LOCAL c := "", nl := Chr(10)
   c += '<xdr:twoCellAnchor>'+nl
   c += _AnchorPt( "from", ::nCol1, 0, ::nRow1, 0 )
   c += _AnchorPt( "to", ::nCol2, 0, ::nRow2, 0 )
   c += '<xdr:graphicFrame macro="">'+nl
   c += '<xdr:nvGraphicFramePr><xdr:cNvPr id="'+HB_NTOS( 1000 + ::ID )+'" name="Chart '+HB_NTOS( ::ID )+'"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>'+nl
   c += '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'+nl
   c += '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
   c += '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdCh'+HB_NTOS( ::ID )+'"/>'
   c += '</a:graphicData></a:graphic>'+nl
   c += '</xdr:graphicFrame>'+nl
   c += '<xdr:clientData/>'+nl
   c += '</xdr:twoCellAnchor>'+nl
RETURN c

// Dodaje grafikon na sheet i vraca ChartML objekat.
//   cTitle  - naslov ("" = bez naslova)
//   cType   - "bar" (uspravni stubici), "hbar" (vodoravni), "line", "pie" (samo prva serija)
//   cCatRef - kategorije, opseg na ovom sheet-u ("A2:A13") ili sa drugog ("'Drugi'!A2:A13")
//   xValRef - opseg vrednosti ili niz opsega za vise serija
//   uFrom, uTo - gornja leva i donja desna celija ("H2", "P20" ili { nRed, nKol })
//   aNames  - imena serija; bez njih se uzima zaglavlje iznad (ili levo od) opsega
// Primer: oSheet:AddChart( "Prihodi 2026", "bar", "A2:A13", { "B2:B13", "C2:C13" }, "E2", "M20" )
// Boje: oChart := oSheet:AddChart(...); oChart:aColors := { "1F77B4", "D62728" }
METHOD AddChart( cTitle, cType, cCatRef, xValRef, uFrom, uTo, aNames ) CLASS WorkSheet
LOCAL oChart := ChartML():New( Self, ::oParent:AddChartId(), cTitle, cType, cCatRef, xValRef, ;
                               IIF( uFrom == NIL, "H2", uFrom ), IIF( uTo == NIL, "P20", uTo ), aNames )
   AAdd( ::aCharts, oChart )
RETURN oChart

// drawing<n>.xml sa svim slikama i grafikonima sheet-a, njegove veze, veza
// sheet -> drawing i xl/charts/chart<ID>.xml za svaki grafikon
METHOD WriteDrawingPart( n ) CLASS WorkSheet
LOCAL cSep := hb_ps(), cDir := ::oParent:GetTempDir() + hb_ps() + "xl" + hb_ps(), h, o, cExt
LOCAL cRelsHdr := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+Chr(10)+ ;
                  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+Chr(10)
   h := FCreate( cDir+"worksheets"+cSep+"_rels"+cSep+"sheet"+HB_NTOS(n)+".xml.rels" )
   FWrite( h, cRelsHdr )
   FWrite( h, '<Relationship Id="rIdDr1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing'+HB_NTOS(n)+'.xml"/>'+Chr(10) )
   FWrite( h, '</Relationships>' )
   FClose( h )

   h := FCreate( cDir+"drawings"+cSep+"_rels"+cSep+"drawing"+HB_NTOS(n)+".xml.rels" )
   FWrite( h, cRelsHdr )
   FOR EACH o IN ::aDrawingsObjects
      cExt := Lower( hb_FNameExt( o:cPath ) )
      FWrite( h, '<Relationship Id="rId'+o:ID+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image'+o:ID+IIF( cExt == ".jpg", ".jpeg", cExt )+'" />'+Chr(10) )
   NEXT
   FOR EACH o IN ::aCharts
      FWrite( h, '<Relationship Id="rIdCh'+HB_NTOS(o:ID)+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart'+HB_NTOS(o:ID)+'.xml"/>'+Chr(10) )
   NEXT
   FWrite( h, '</Relationships>' )
   FClose( h )

   h := FCreate( cDir+"drawings"+cSep+"drawing"+HB_NTOS(n)+".xml" )
   FWrite( h, '<?xml version="1.0" encoding="utf-8"?>'+Chr(10) )
   FWrite( h, '<xdr:wsDr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">'+Chr(10) )
   FOR EACH o IN ::aDrawingsObjects
      FWrite( h, o:AnchorXML( o:__enumIndex() ) )
   NEXT
   FOR EACH o IN ::aCharts
      FWrite( h, o:AnchorXML() )
      o:WritePart( cDir+"charts"+cSep )
   NEXT
   FWrite( h, '</xdr:wsDr>' )
   FClose( h )
RETURN Self

// Opseg -> puna adresa "'Sheet'!$A$2:$A$13"; u aCells dodaje { red, kolona } svake
// celije, a u oData vraca sheet sa podacima (za kes vrednosti; NIL ako ga nema).
// Prihvata "A2:A13" (ovaj sheet), "'Drugi sheet'!A2:A13" i "Drugi!$A$2:$A$13".
STATIC FUNCTION _ChartRef( cRef, oWS, aCells, oData )
LOCAL aP, nR1 := 0, nC1 := 0, nR2 := 0, nC2 := 0, i, j, nPos, cSheet, cRange
   nPos := RAt( "!", cRef )
   IF nPos > 0
      cSheet := AllTrim( Left( cRef, nPos - 1 ) )
      cRange := SubStr( cRef, nPos + 1 )
      IF Left( cSheet, 1 ) == "'" .AND. Right( cSheet, 1 ) == "'"
         cSheet := StrTran( SubStr( cSheet, 2, Len( cSheet ) - 2 ), "''", "'" )
      ENDIF
      oData := oWS:oParent:GetSheet( cSheet )
   ELSE
      cSheet := oWS:cName
      cRange := cRef
      oData := oWS
   ENDIF
   aP := hb_ATokens( StrTran( Upper( AllTrim( cRange ) ), "$", "" ), ":" )
   WORKSHEET_RC( aP[1], @nR1, @nC1 )
   IF LEN( aP ) > 1
      WORKSHEET_RC( aP[2], @nR2, @nC2 )
   ELSE
      nR2 := nR1
      nC2 := nC1
   ENDIF
   FOR i := nR1 TO nR2
      FOR j := nC1 TO nC2
         AAdd( aCells, { i, j } )
      NEXT
   NEXT
RETURN "'" + StrTran( oWS:oParent:ToUTF8( cSheet ), "'", "''" ) + "'!$" + ColumnIndexToColumnLetter( nC1 ) + "$" + HB_NTOS( nR1 ) + ;
       IIF( LEN( aP ) > 1, ":$" + ColumnIndexToColumnLetter( nC2 ) + "$" + HB_NTOS( nR2 ), "" )

// Kes vrednosti iz celija sheet-a (Excel/Calc ga koriste dok ne preracunaju,
// neki pregledaci samo njega). Formule nemaju vrednost u klasi pa se preskacu.
STATIC FUNCTION _ChartCache( oWS, aCells, lNum )
LOCAL c, i, x, v
   IF oWS == NIL .OR. LEN( aCells ) == 0
      RETURN ""
   ENDIF
   c := IIF( lNum, '<c:numCache><c:formatCode>General</c:formatCode>', '<c:strCache>' )
   c += '<c:ptCount val="'+HB_NTOS( LEN( aCells ) )+'"/>'
   FOR i := 1 TO LEN( aCells )
      x := oWS:ReadCell( aCells[i] )
      v := NIL
      DO CASE
      CASE HB_ISNUMERIC( x )
         v := HB_NTOC( x )
      CASE lNum
      CASE HB_ISSTRING( x ) .AND. !( Left( x, 1 ) $ "=" + Chr(0) )
         v := XmlEsc( x )
      CASE HB_ISDATE( x ) .AND. ! Empty( x )
         v := DToC( x )
      ENDCASE
      IF v != NIL
         c += '<c:pt idx="'+HB_NTOS( i - 1 )+'"><c:v>'+v+'</c:v></c:pt>'
      ENDIF
   NEXT
RETURN c + IIF( lNum, '</c:numCache>', '</c:strCache>' )

STATIC FUNCTION _Fill( cRGB )
RETURN '<a:solidFill><a:srgbClr val="'+cRGB+'"/></a:solidFill>'

// tacna pretraga stringa u nizu (ASCAN bez SET EXACT poredi samo pocetak)
STATIC FUNCTION _AScanX( a, x )
RETURN AScan( a, {| e | HB_ISSTRING( e ) .AND. e == x } )

// sheet po imenu bez obzira na velika/mala slova (Excel ih ne razlikuje)
STATIC FUNCTION _SheetPos( a, cName )
LOCAL cU := Upper( cName )
RETURN AScan( a, {| e | Upper( e ) == cU } )

// ime sheet-a po pravilima Excela: bez : \ / ? * [ ], ne pocinje i ne zavrsava
// se apostrofom, najvise 31 znak
STATIC FUNCTION _SheetName( cName )
cName := hb_StrReplace( IIF( HB_ISSTRING( cName ), AllTrim( cName ), "" ), ":\/?*[]", "_______" )
IF Left( cName, 1 ) == "'"
   cName := "_" + SubStr( cName, 2 )
ENDIF
IF Right( cName, 1 ) == "'"
   cName := Left( cName, Len( cName ) - 1 ) + "_"
ENDIF
IF Empty( cName )
   cName := "Sheet"
ENDIF
RETURN IIF( hb_StrIsUTF8( cName ), hb_UTF8Left( cName, 31 ), Left( cName, 31 ) )

// Date / DateTime -> Excel serijski broj (1900 sistem). Excel broji nepostojeci
// 29.02.1900, pa su datumi pre 01.03.1900 za jedan manji.
STATIC FUNCTION _DateToSerial( x )
LOCAL d := IIF( ValType( x ) == "T", hb_TToD( x ), x ), nDays, nFrac := 0
nDays := d - hb_SToD( IIF( d < hb_SToD( "19000301" ), "18991231", "18991230" ) )
IF ValType( x ) == "T"
   nFrac := x - hb_DToT( d )            // deo dana (razlika dva timestamp-a je u danima)
ENDIF
RETURN IIF( nFrac == 0, HB_NTOS( nDays ), HB_NTOC( nDays + nFrac ) )

// boja za styles.xml: ARGB sa 8 heks cifara ("FF0000" -> "FFFF0000")
STATIC FUNCTION _ARGB( c )
c := Upper( AllTrim( _S( c ) ) )
IF Left( c, 1 ) == "#"
   c := SubStr( c, 2 )
ENDIF
RETURN IIF( Len( c ) == 6, "FF" + c, c )

// { nRed, nKol } ili "B3" -> red i kolona (1-based)
STATIC PROCEDURE _CellRC( u, nRow, nCol )
   IF HB_ISARRAY( u )
      nRow := u[1]
      nCol := u[2]
   ELSE
      WORKSHEET_RC( u, @nRow, @nCol )
   ENDIF
RETURN

STATIC FUNCTION _AnchorPt( cTag, xCol, xColOff, xRow, xRowOff )
LOCAL t := Chr(9), nl := Chr(10)
RETURN t+'<xdr:'+cTag+'>'+nl+ ;
       t+t+'<xdr:col>'+_S(xCol)+'</xdr:col>'+nl+ ;
       t+t+'<xdr:colOff>'+_S(xColOff)+'</xdr:colOff>'+nl+ ;
       t+t+'<xdr:row>'+_S(xRow)+'</xdr:row>'+nl+ ;
       t+t+'<xdr:rowOff>'+_S(xRowOff)+'</xdr:rowOff>'+nl+ ;
       t+'</xdr:'+cTag+'>'+nl

STATIC FUNCTION _S( x )
RETURN IIF( HB_ISNUMERIC( x ), HB_NTOS( x ), IIF( HB_ISSTRING( x ), x, "" ) )


// ---------------------------------------------------------------------------
// Pomocne funkcije za citanje (Read)
// ---------------------------------------------------------------------------

// vrednost celije iz niza podataka sheet-a ("B3" ili { nRed, nKol }), bez obrade
STATIC FUNCTION _CellGet( aData, uAddr )
LOCAL nRow := 0, nCol := 0
IF ! HB_ISARRAY( uAddr ) .AND. ! HB_ISSTRING( uAddr )
   RETURN NIL
ENDIF
_CellRC( uAddr, @nRow, @nCol )
IF ! HB_ISNUMERIC( nRow ) .OR. ! HB_ISNUMERIC( nCol ) .OR. ;
   nRow < 1 .OR. nRow > LEN( aData ) .OR. nCol < 1 .OR. nCol > LEN( aData[ nRow ] )
   RETURN NIL
ENDIF
RETURN aData[ nRow, nCol ]

// Jednostavan i brz citac XML-a za Read(): fajl se podeli na delove po "<", a
// svaki deo je oblika 'ime atr="..">tekst'. TXmlDocument (hbxml) je ranije za
// sheet od 20.000 redova trosio ~45 s, a uz to brisao razmake na pocetku i kraju
// teksta i &#NN; reference.
STATIC FUNCTION _XTokens( cFile )
RETURN IIF( hb_FileExists( cFile ), hb_ATokens( MemoRead( cFile ), "<" ), {} )

// Jedan deo XML-a: vraca tekst atributa, a u cName ime elementa (bez prefiksa
// "x:"), u nKind 1 <a>, 2 </a>, 3 <a/>, 0 ostalo (<?xml?>, komentar), u cText
// tekst posle elementa (do sledeceg "<", nedekodiran).
STATIC FUNCTION _XTok( cTok, cName, nKind, cText )
LOCAL n := At( ">", cTok ), cHead, ch
cName := ""
nKind := 0
cText := ""
IF n == 0
   RETURN ""
ENDIF
// ">" sme da stoji u vrednosti atributa: trazi se prvi van navodnika
IF '"' $ cTok
   DO WHILE n > 0 .AND. ( Len( Left( cTok, n ) ) - Len( StrTran( Left( cTok, n ), '"', "" ) ) ) % 2 == 1
      n := hb_At( ">", cTok, n + 1 )
   ENDDO
   IF n == 0
      RETURN ""
   ENDIF
ENDIF
cHead := Left( cTok, n - 1 )
cText := SubStr( cTok, n + 1 )
ch := Left( cHead, 1 )
IF ch == "?" .OR. ch == "!"
   RETURN ""
ENDIF
IF ch == "/"
   nKind := 2
   cHead := SubStr( cHead, 2 )
ELSEIF Right( cHead, 1 ) == "/"
   nKind := 3
   cHead := Left( cHead, Len( cHead ) - 1 )
ELSE
   nKind := 1
ENDIF
cHead := hb_StrReplace( cHead, Chr( 9 ) + Chr( 10 ) + Chr( 13 ), "   " )
n := At( " ", cHead )
cName := IIF( n == 0, cHead, Left( cHead, n - 1 ) )
IF ( ch := At( ":", cName ) ) > 0
   cName := SubStr( cName, ch + 1 )
ENDIF
RETURN IIF( n == 0, "", SubStr( cHead, n + 1 ) )

// 'a="1" b="x &amp; y"' -> { "a" => "1", "b" => "x & y" }
STATIC FUNCTION _XAttrs( cA )
LOCAL h := { => }, n, m, q, cKey
DO WHILE ( n := At( "=", cA ) ) > 0
   cKey := AllTrim( Left( cA, n - 1 ) )
   cA := LTrim( SubStr( cA, n + 1 ) )
   q := Left( cA, 1 )
   IF !( q == '"' .OR. q == "'" ) .OR. ( m := hb_At( q, cA, 2 ) ) == 0
      EXIT
   ENDIF
   h[ cKey ] := _XEnt( SubStr( cA, 2, m - 2 ) )
   cA := SubStr( cA, m + 1 )
ENDDO
RETURN h

// atribut iz hash-a ili ""; "id" nalazi i "r:id"
STATIC FUNCTION _HAttr( h, cName )
LOCAL k
IF hb_HHasKey( h, cName )
   RETURN h[ cName ]
ENDIF
FOR EACH k IN hb_HKeys( h )
   IF Right( k, Len( cName ) + 1 ) == ":" + cName
      RETURN h[ k ]
   ENDIF
NEXT
RETURN ""

// XML entiteti: &amp; &lt; &gt; &quot; &apos; &#NN; &#xHH;
STATIC FUNCTION _XEnt( c )
LOCAL a, cOut, i, n, cE
IF ! "&" $ c
   RETURN c
ENDIF
a := hb_ATokens( c, "&" )
cOut := a[ 1 ]
FOR i := 2 TO LEN( a )
   n := At( ";", a[ i ] )
   cE := IIF( n > 1 .AND. n <= 10, Left( a[ i ], n - 1 ), "" )
   DO CASE
   CASE cE == "amp"  ; cOut += "&"
   CASE cE == "lt"   ; cOut += "<"
   CASE cE == "gt"   ; cOut += ">"
   CASE cE == "quot" ; cOut += '"'
   CASE cE == "apos" ; cOut += "'"
   CASE Lower( Left( cE, 2 ) ) == "#x" ; cOut += hb_UTF8Chr( hb_HexToNum( SubStr( cE, 3 ) ) )
   CASE Left( cE, 1 ) == "#" .AND. Len( cE ) > 1 ; cOut += hb_UTF8Chr( Val( SubStr( cE, 2 ) ) )
   OTHERWISE
      cOut += "&" + a[ i ]
      LOOP
   ENDCASE
   cOut += SubStr( a[ i ], n + 1 )
NEXT
RETURN cOut

// Excelov zapis _xHHHH_ za kontrolne znakove (npr. _x000D_ = CR)
STATIC FUNCTION _XDecode( c )
LOCAL n, cCode
IF ! HB_ISSTRING( c )
   RETURN c
ENDIF
IF "_x" $ c
   n := 1
   DO WHILE ( n := hb_At( "_x", c, n ) ) > 0
      cCode := SubStr( c, n + 2, 4 )
      IF SubStr( c, n + 6, 1 ) == "_" .AND. Len( cCode ) == 4 .AND. ;
         Empty( hb_StrReplace( Upper( cCode ), "0123456789ABCDEF", "" ) )
         c := Left( c, n - 1 ) + hb_UTF8Chr( hb_HexToNum( cCode ) ) + SubStr( c, n + 7 )
      ENDIF
      n++
   ENDDO
ENDIF
RETURN c

// broj iz <v>; Val() ne razume eksponent (1.5E-3, 2E+20), a hb_jsonDecode() ne prima "E+"
STATIC FUNCTION _XVal( c )
LOCAL x
IF "E" $ Upper( c )
   x := hb_jsonDecode( StrTran( Upper( c ), "+", "" ) )
   RETURN IIF( HB_ISNUMERIC( x ), x, Val( c ) )
ENDIF
RETURN Val( c )

// "2026-11-14" -> Date, "2026-11-14T10:20:30" -> DateTime
STATIC FUNCTION _IsoToDate( c )
c := Left( hb_StrReplace( c, "-:T", "" ), 14 )
RETURN IIF( Len( c ) <= 8, hb_SToD( c ), hb_SToT( c ) )

// Excel serijski broj -> Date (nKind 1) ili DateTime (2); samo vreme (3) ostaje broj
STATIC FUNCTION _SerialToDate( n, nKind, l1904 )
LOCAL d, nDays := Int( n )
IF nKind == 3
   RETURN n
ENDIF
IF l1904
   d := hb_SToD( "19040101" ) + nDays
ELSEIF nDays < 61
   d := hb_SToD( "18991231" ) + nDays        // pre nepostojeceg 29.02.1900
ELSE
   d := hb_SToD( "18991230" ) + nDays
ENDIF
IF nKind == 1
   RETURN d
ENDIF
RETURN hb_DToT( d ) + Round( ( n - nDays ) * 86400, 0 ) / 86400

// Cita xl/styles.xml i vraca { aKinds, aXfMap } za svaki stil fajla (cellXfs, redom):
// aKinds - 0 broj, 1 datum, 2 datum+vreme, 3 vreme; aXfMap - stil ovog workbook-a
// napravljen sa NewFont/NewFillPattern/NewBorder/NewFormat/NewStyle (0 = bez stila).
// Boje mogu biti rgb, iz teme (theme1.xml, sa tint) ili indeksirane (paleta).
STATIC FUNCTION _ReadStyles( oWB, cFile, cThemeFile )
LOCAL aTok := _XTokens( cFile ), aTheme := _ThemeColors( cThemeFile ), aIdx := _Palette()
LOCAL aKinds := {}, aMap := {}, hFmt := { => }, aFonts := {}, aFills := {}, aBorders := {}
LOCAL hSec := { "numFmts" => 1, "fonts" => 1, "fills" => 1, "borders" => 1, "cellStyleXfs" => 1, ;
                "cellXfs" => 1, "cellStyles" => 1, "dxfs" => 1, "tableStyles" => 1, "colors" => 1, ;
                "indexedColors" => 1, "mruColors" => 1, "extLst" => 1 }
LOCAL cTok, cA, cTag, nKind, h, cSec := "", x, cSide := "", nId, n := 0
// 1. prolaz: sopstvena paleta indeksiranih boja (u fajlu je posle stilova)
FOR EACH cTok IN aTok
   cA := _XTok( cTok, @cTag, @nKind )
   IF cTag == "indexedColors"
      cSec := IIF( nKind == 1, cTag, "" )
   ELSEIF cTag == "rgbColor" .AND. cSec == "indexedColors" .AND. nKind != 2
      IF ++n <= LEN( aIdx )
         aIdx[ n ] := Right( _ARGB( hb_HGetDef( _XAttrs( cA ), "rgb", "FF000000" ) ), 6 )
      ENDIF
   ENDIF
NEXT
cSec := ""
FOR EACH cTok IN aTok
   cA := _XTok( cTok, @cTag, @nKind )
   IF nKind == 0
      LOOP
   ENDIF
   IF hb_HHasKey( hSec, cTag )
      IF nKind == 1
         cSec := cTag
      ELSEIF nKind == 2
         cSec := ""
      ENDIF
      LOOP
   ENDIF
   DO CASE
   CASE cSec == "numFmts"
      IF cTag == "numFmt" .AND. nKind != 2
         h := _XAttrs( cA )
         hFmt[ Val( hb_HGetDef( h, "numFmtId", "0" ) ) ] := hb_HGetDef( h, "formatCode", "" )
      ENDIF
   CASE cSec == "fonts"
      DO CASE
      CASE cTag == "font"
         IF nKind != 2
            x := { "name" => "Calibri", "sz" => 11, "b" => .F., "i" => .F., "u" => .F., "s" => .F., "color" => "" }
         ENDIF
         IF nKind != 1
            AAdd( aFonts, x )
         ENDIF
      CASE x == NIL .OR. nKind == 2
      CASE cTag == "b" .OR. cTag == "i" .OR. cTag == "strike"
         x[ IIF( cTag == "strike", "s", cTag ) ] := _XBool( _XAttrs( cA ), "val", .T. )
      CASE cTag == "u"
         x[ "u" ] := !( Lower( hb_HGetDef( _XAttrs( cA ), "val", "single" ) ) == "none" )
      CASE cTag == "sz"
         x[ "sz" ] := Val( hb_HGetDef( _XAttrs( cA ), "val", "11" ) )
      CASE cTag == "name"
         x[ "name" ] := hb_HGetDef( _XAttrs( cA ), "val", "Calibri" )
      CASE cTag == "color"
         x[ "color" ] := _XColor( _XAttrs( cA ), aTheme, aIdx )
      ENDCASE
   CASE cSec == "fills"
      DO CASE
      CASE cTag == "fill"
         IF nKind != 2
            x := { "type" => "none", "fg" => "", "bg" => "" }
         ENDIF
         IF nKind != 1
            AAdd( aFills, x )
         ENDIF
      CASE x == NIL .OR. nKind == 2
      CASE cTag == "patternFill"
         x[ "type" ] := hb_HGetDef( _XAttrs( cA ), "patternType", "none" )
      CASE cTag == "gradientFill"            // prelaz: uzima se prva boja kao puna
         x[ "type" ] := "gradient"
      CASE cTag == "fgColor" .OR. ( cTag == "color" .AND. x[ "type" ] == "gradient" .AND. Empty( x[ "fg" ] ) )
         x[ "fg" ] := _XColor( _XAttrs( cA ), aTheme, aIdx )
      CASE cTag == "bgColor"
         x[ "bg" ] := _XColor( _XAttrs( cA ), aTheme, aIdx )
      ENDCASE
   CASE cSec == "borders"
      DO CASE
      CASE cTag == "border"
         IF nKind != 2
            x := { "left" => { 0, "" }, "right" => { 0, "" }, "top" => { 0, "" }, "bottom" => { 0, "" }, "diagonal" => { 0, "" } }
         ENDIF
         IF nKind != 1
            AAdd( aBorders, x )
         ENDIF
      CASE x == NIL
      CASE cTag == "left" .OR. cTag == "right" .OR. cTag == "top" .OR. cTag == "bottom" .OR. ;
           cTag == "diagonal" .OR. cTag == "start" .OR. cTag == "end"
         IF nKind == 2
            cSide := ""
         ELSE
            cSide := IIF( cTag == "start", "left", IIF( cTag == "end", "right", cTag ) )
            x[ cSide ][ 1 ] := _AScanX( aBordersStyle, hb_HGetDef( _XAttrs( cA ), "style", "" ) )
            IF nKind == 3
               cSide := ""
            ENDIF
         ENDIF
      CASE cTag == "color" .AND. ! Empty( cSide ) .AND. nKind != 2
         x[ cSide ][ 2 ] := _XColor( _XAttrs( cA ), aTheme, aIdx )
      ENDCASE
   CASE cSec == "cellXfs"
      DO CASE
      CASE cTag == "xf"
         IF nKind != 2
            h := _XAttrs( cA )
            x := { "numFmtId" => Val( hb_HGetDef( h, "numFmtId", "0" ) ), "fontId" => Val( hb_HGetDef( h, "fontId", "0" ) ), ;
                   "fillId" => Val( hb_HGetDef( h, "fillId", "0" ) ), "borderId" => Val( hb_HGetDef( h, "borderId", "0" ) ), ;
                   "ha" => 0, "va" => 0, "rot" => 0, "wrap" => .F. }
         ENDIF
         IF nKind != 1 .AND. x != NIL
            nId := x[ "numFmtId" ]
            AAdd( aKinds, IIF( hb_HHasKey( hFmt, nId ), _FmtKind( hFmt[ nId ] ), _BuiltinFmtKind( nId ) ) )
            AAdd( aMap, _MapStyle( oWB, x, aFonts, aFills, aBorders, hFmt ) )
         ENDIF
      CASE cTag == "alignment" .AND. x != NIL .AND. nKind != 2
         h := _XAttrs( cA )
         x[ "ha" ]   := hb_HGetDef( { "left" => 1, "center" => 2, "centerContinuous" => 2, "right" => 3 }, hb_HGetDef( h, "horizontal", "" ), 0 )
         x[ "va" ]   := hb_HGetDef( { "top" => 1, "center" => 2 }, hb_HGetDef( h, "vertical", "" ), 0 )
         x[ "rot" ]  := Val( hb_HGetDef( h, "textRotation", "0" ) )
         x[ "wrap" ] := _XBool( h, "wrapText" )
      ENDCASE
   ENDCASE
NEXT
RETURN { aKinds, aMap }

// stil iz fajla (xf) -> broj stila ovog workbook-a (0 = podrazumevani izgled)
STATIC FUNCTION _MapStyle( oWB, x, aFonts, aFills, aBorders, hFmt )
LOCAL f, fi, b, nFont, nFill, nBorder, nFmt, nPat, nId := x[ "numFmtId" ], c, i, aC := {}
LOCAL aSide := { "left", "right", "top", "bottom", "diagonal" }
IF x[ "fontId" ] + 1 <= LEN( aFonts )
   f := aFonts[ x[ "fontId" ] + 1 ]
   c := IIF( Empty( f[ "color" ] ) .OR. f[ "color" ] == "FF000000", "000000", f[ "color" ] )
   nFont := oWB:NewFont( f[ "name" ], f[ "sz" ], f[ "b" ], f[ "i" ], f[ "u" ], f[ "s" ], c )
   IF nFont == 1                                  // isto sto i podrazumevani font
      nFont := NIL
   ENDIF
ENDIF
IF x[ "fillId" ] + 1 <= LEN( aFills )
   fi := aFills[ x[ "fillId" ] + 1 ]
   nPat := IIF( fi[ "type" ] == "gradient", 1, _AScanX( aPattern, fi[ "type" ] ) )
   IF nPat > 0
      nFill := oWB:NewFillPattern( nPat, IIF( Empty( fi[ "fg" ] ), "FF000000", fi[ "fg" ] ), ;
                                         IIF( Empty( fi[ "bg" ] ), "FFFFFFFF", fi[ "bg" ] ) )
   ENDIF
ENDIF
IF x[ "borderId" ] + 1 <= LEN( aBorders )
   b := aBorders[ x[ "borderId" ] + 1 ]
   IF AScan( aSide, {| s | b[ s ][ 1 ] > 0 } ) > 0
      FOR i := 1 TO 5
         c := b[ aSide[ i ] ][ 2 ]
         AAdd( aC, IIF( Empty( c ) .OR. c == "FF000000", NIL, c ) )
      NEXT
      nBorder := oWB:NewBorder( b[ "left" ][ 1 ], b[ "right" ][ 1 ], b[ "top" ][ 1 ], b[ "bottom" ][ 1 ], b[ "diagonal" ][ 1 ], ;
                                aC[ 1 ], aC[ 2 ], aC[ 3 ], aC[ 4 ], aC[ 5 ] )
   ENDIF
ENDIF
IF hb_HHasKey( hFmt, nId )
   nFmt := oWB:NewFormat( hFmt[ nId ], .T. )
ELSEIF nId > 0
   nFmt := -nId                                   // ugradjeni format
ENDIF
IF nFont == NIL .AND. nFill == NIL .AND. nBorder == NIL .AND. nFmt == NIL .AND. ;
   x[ "ha" ] == 0 .AND. x[ "va" ] == 0 .AND. x[ "rot" ] == 0 .AND. ! x[ "wrap" ]
   RETURN 0
ENDIF
RETURN oWB:NewStyle( nFont, nBorder, nFill, IIF( x[ "va" ] > 0, x[ "va" ], NIL ), IIF( x[ "ha" ] > 0, x[ "ha" ], NIL ), ;
                     nFmt, IIF( x[ "rot" ] > 0, x[ "rot" ], NIL ), IIF( x[ "wrap" ], .T., NIL ) )

// boja iz atributa (rgb / theme + tint / indexed / auto) -> ARGB, "" = nije zadata
STATIC FUNCTION _XColor( h, aTheme, aIdx )
LOCAL c := "", n
DO CASE
CASE hb_HHasKey( h, "rgb" )
   c := Right( _ARGB( h[ "rgb" ] ), 6 )
CASE hb_HHasKey( h, "theme" )
   n := Val( h[ "theme" ] ) + 1
   c := IIF( n >= 1 .AND. n <= LEN( aTheme ), aTheme[ n ], "000000" )
CASE hb_HHasKey( h, "indexed" )
   n := Val( h[ "indexed" ] ) + 1
   c := IIF( n >= 1 .AND. n <= LEN( aIdx ), aIdx[ n ], IIF( n == 66, "FFFFFF", "000000" ) )   // 64 tekst, 65 pozadina
CASE hb_HHasKey( h, "auto" )
   c := "000000"
ENDCASE
IF Empty( c )
   RETURN ""
ENDIF
IF hb_HHasKey( h, "tint" ) .AND. Val( h[ "tint" ] ) != 0
   c := _Tint( c, Val( h[ "tint" ] ) )
ENDIF
RETURN "FF" + Upper( c )

// Excel "tint": posvetljavanje (> 0) ili potamnjivanje (< 0) boje preko HSL
STATIC FUNCTION _Tint( c, nTint )
LOCAL r := hb_HexToNum( SubStr( c, 1, 2 ) ) / 255, g := hb_HexToNum( SubStr( c, 3, 2 ) ) / 255
LOCAL b := hb_HexToNum( SubStr( c, 5, 2 ) ) / 255, nMax, nMin, h := 0, s := 0, l, d, p, q
nMax := Max( r, Max( g, b ) )
nMin := Min( r, Min( g, b ) )
l := ( nMax + nMin ) / 2
IF nMax != nMin
   d := nMax - nMin
   s := IIF( l > 0.5, d / ( 2 - nMax - nMin ), d / ( nMax + nMin ) )
   DO CASE
   CASE nMax == r ; h := ( g - b ) / d + IIF( g < b, 6, 0 )
   CASE nMax == g ; h := ( b - r ) / d + 2
   OTHERWISE      ; h := ( r - g ) / d + 4
   ENDCASE
   h := h / 6
ENDIF
l := IIF( nTint < 0, l * ( 1 + nTint ), l * ( 1 - nTint ) + nTint )
IF s == 0
   r := g := b := l
ELSE
   q := IIF( l < 0.5, l * ( 1 + s ), l + s - l * s )
   p := 2 * l - q
   r := _Hue( p, q, h + 1 / 3 )
   g := _Hue( p, q, h )
   b := _Hue( p, q, h - 1 / 3 )
ENDIF
RETURN _Hex2( r ) + _Hex2( g ) + _Hex2( b )

STATIC FUNCTION _Hue( p, q, t )
IF t < 0
   t += 1
ENDIF
IF t > 1
   t -= 1
ENDIF
DO CASE
CASE t < 1 / 6 ; RETURN p + ( q - p ) * 6 * t
CASE t < 1 / 2 ; RETURN q
CASE t < 2 / 3 ; RETURN p + ( q - p ) * ( 2 / 3 - t ) * 6
ENDCASE
RETURN p

STATIC FUNCTION _Hex2( x )
RETURN hb_NumToHex( Int( Max( 0, Min( 255, Round( x * 255, 0 ) ) ) ), 2 )

// boje teme po indeksu iz styles.xml: lt1, dk1, lt2, dk2, accent1..6, hlink, folHlink
// (bez theme1.xml vazi Office tema)
STATIC FUNCTION _ThemeColors( cFile )
LOCAL a := { "FFFFFF", "000000", "E7E6E6", "44546A", "4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47", "0563C1", "954F72" }
LOCAL hPos := { "lt1" => 1, "dk1" => 2, "lt2" => 3, "dk2" => 4, "accent1" => 5, "accent2" => 6, "accent3" => 7, ;
                "accent4" => 8, "accent5" => 9, "accent6" => 10, "hlink" => 11, "folHlink" => 12 }
LOCAL cTok, cA, cTag, nKind, h, n := 0, lIn := .F., lDone := .F.
FOR EACH cTok IN _XTokens( cFile )
   cA := _XTok( cTok, @cTag, @nKind )
   DO CASE
   CASE cTag == "clrScheme"
      IF nKind == 1 .AND. ! lDone
         lIn := .T.
      ELSEIF nKind == 2 .AND. lIn
         lIn := .F.
         lDone := .T.
      ENDIF
   CASE ! lIn
   CASE hb_HHasKey( hPos, cTag ) .AND. nKind == 1
      n := hPos[ cTag ]
   CASE n > 0 .AND. nKind != 2 .AND. ( cTag == "srgbClr" .OR. cTag == "sysClr" )
      h := _XAttrs( cA )
      a[ n ] := Upper( hb_HGetDef( h, IIF( cTag == "srgbClr", "val", "lastClr" ), a[ n ] ) )
      n := 0
   ENDCASE
NEXT
RETURN a

// podrazumevana Excel paleta indeksiranih boja (0-63)
STATIC FUNCTION _Palette()
RETURN { "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF", ;
         "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF", ;
         "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080", ;
         "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF", ;
         "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF", ;
         "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99", ;
         "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696", ;
         "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333" }

// atribut kao logicka vrednost ("1"/"true"); lDef ako atributa nema
STATIC FUNCTION _XBool( h, cKey, lDef )
LOCAL c
IF ! hb_HHasKey( h, cKey )
   RETURN IIF( HB_ISLOGICAL( lDef ), lDef, .F. )
ENDIF
c := Lower( h[ cKey ] )
RETURN c == "1" .OR. c == "true"

// zaglavlje/podnozje: usamljen "&" (npr. "Procter & Gamble") Excel tumaci kao kod i
// brise ga, pa se udvaja; kodovi &P &N &D &T &F &A &B &I &U &"Font" &12 ... ostaju
STATIC FUNCTION _HFEsc( c )
LOCAL cOut := "", i := 1, ch, ch2
DO WHILE i <= Len( c )
   ch := SubStr( c, i, 1 )
   IF ch == "&"
      ch2 := SubStr( c, i + 1, 1 )
      IF ! ch2 == "" .AND. ch2 $ 'PNDTFAZGLCRBIUSXYEOHK&"+-0123456789'
         cOut += ch + ch2
         i += 2
      ELSE
         cOut += "&&"
         i++
      ENDIF
      LOOP
   ENDIF
   cOut += ch
   i++
ENDDO
RETURN cOut

// zaglavlje/podnozje "&LLevo&CSredina&RDesno" -> { levo, sredina, desno };
// ostali kodovi (&P, &N, &D, &&, &"Arial,Bold"...) ostaju u tekstu
STATIC FUNCTION _HFParts( c )
LOCAL a := { "", "", "" }, n := 2, i := 1, ch, ch2
DO WHILE i <= Len( c )
   ch := SubStr( c, i, 1 )
   IF ch == "&" .AND. i < Len( c )
      ch2 := SubStr( c, i + 1, 1 )
      DO CASE
      CASE ch2 == "L" ; n := 1
      CASE ch2 == "C" ; n := 2
      CASE ch2 == "R" ; n := 3
      OTHERWISE       ; a[ n ] += ch + ch2
      ENDCASE
      i += 2
      LOOP
   ENDIF
   a[ n ] += ch
   i++
ENDDO
RETURN a

// kod ugradjenog formata broja (ECMA-376 18.8.30)
STATIC FUNCTION _BuiltinFmtCode( n )
STATIC s_h
IF s_h == NIL
   s_h := { 0 => "General", 1 => "0", 2 => "0.00", 3 => "#,##0", 4 => "#,##0.00", 9 => "0%", 10 => "0.00%", ;
            11 => "0.00E+00", 12 => "# ?/?", 13 => "# ??/??", 14 => "mm-dd-yy", 15 => "d-mmm-yy", 16 => "d-mmm", ;
            17 => "mmm-yy", 18 => "h:mm AM/PM", 19 => "h:mm:ss AM/PM", 20 => "h:mm", 21 => "h:mm:ss", ;
            22 => "m/d/yy h:mm", 37 => "#,##0 ;(#,##0)", 38 => "#,##0 ;[Red](#,##0)", 39 => "#,##0.00;(#,##0.00)", ;
            40 => "#,##0.00;[Red](#,##0.00)", 45 => "mm:ss", 46 => "[h]:mm:ss", 47 => "mmss.0", 48 => "##0.0E+0", 49 => "@" }
ENDIF
RETURN hb_HGetDef( s_h, n, "General" )

// ugradjeni formati (ECMA-376 18.8.30); 27-36 i 50-58 su istocnoazijski datumi
STATIC FUNCTION _BuiltinFmtKind( n )
DO CASE
CASE ( n >= 14 .AND. n <= 17 ) .OR. ( n >= 27 .AND. n <= 31 ) .OR. ;
     ( n >= 34 .AND. n <= 36 ) .OR. ( n >= 50 .AND. n <= 58 )
   RETURN 1
CASE n == 22
   RETURN 2
CASE ( n >= 18 .AND. n <= 21 ) .OR. ( n >= 45 .AND. n <= 47 ) .OR. n == 32 .OR. n == 33
   RETURN 3
ENDCASE
RETURN 0

// korisnicki format: posle uklanjanja "teksta", \x, [Red], [$-409]... trazi d/m/y i h/s
STATIC FUNCTION _FmtKind( cFmt )
LOCAL c := "", i := 1, n, ch, cIn, lD, lT
DO WHILE i <= Len( cFmt )
   ch := SubStr( cFmt, i, 1 )
   DO CASE
   CASE ch == '"'
      n := hb_At( '"', cFmt, i + 1 )
      i := IIF( n == 0, Len( cFmt ), n )
   CASE ch $ "\_*"
      i++
   CASE ch == "["
      n := hb_At( "]", cFmt, i + 1 )
      cIn := Lower( SubStr( cFmt, i + 1, IIF( n == 0, Len( cFmt ) + 1, n ) - i - 1 ) )
      IF ! Empty( cIn ) .AND. Empty( hb_StrReplace( cIn, "hms", "" ) )   // [h] [mm] [ss]
         c += "h"
      ENDIF
      i := IIF( n == 0, Len( cFmt ), n )
   CASE ch == ";"                       // samo prvi deo (pozitivni brojevi)
      EXIT
   OTHERWISE
      c += Lower( ch )
   ENDCASE
   i++
ENDDO
c := StrTran( c, "general", "" )
lT := "h" $ c .OR. "s" $ c
lD := "d" $ c .OR. "y" $ c .OR. ( "m" $ c .AND. ! lT )
RETURN IIF( lD, IIF( lT, 2, 1 ), IIF( lT, 3, 0 ) )

//------------------------------------------------------------- Conditional() helpers

// "a1:b5, $D$1:$D$5" -> "A1:B5 D1:D5" (sqref)
STATIC FUNCTION _CfRange( c )
LOCAL a := {}, x
IF ! HB_ISSTRING( c )
	RETURN ""
ENDIF
FOR EACH x IN hb_ATokens( StrTran( Upper( StrTran( c, "$", "" ) ), ",", " " ), " " )
	IF ! Empty( x )
		AAdd( a, AllTrim( x ) )
	ENDIF
NEXT
RETURN _Join( a, " " )

STATIC FUNCTION _Join( a, cSep )
LOCAL c := "", x
FOR EACH x IN a
	c += IIF( x:__enumIndex() == 1, "", cSep ) + x
NEXT
RETURN c

// a value of a rule as a formula: a number, a date (its serial number), TRUE / FALSE, a text in
// quotes, "=..." a formula
STATIC FUNCTION _CfValue( x )
DO CASE
CASE HB_ISNUMERIC( x )
	RETURN hb_ntos( x )
CASE HB_ISDATE( x )
	RETURN hb_ntos( x - hb_SToD( "18991230" ) )
CASE HB_ISDATETIME( x )
	RETURN hb_ntos( Round( x - hb_SToT( "18991230000000" ), 10 ) )
CASE HB_ISLOGICAL( x )
	RETURN IIF( x, "TRUE", "FALSE" )
CASE HB_ISSTRING( x ) .AND. Left( x, 1 ) == "="
	RETURN SubStr( x, 2 )
CASE HB_ISSTRING( x )
	RETURN _CfStr( x )
ENDCASE
RETURN "0"

// a text for a formula: in quotes, a quote doubled
STATIC FUNCTION _CfStr( c )
RETURN '"' + StrTran( c, '"', '""' ) + '"'

//------------------------------------------------------------- AutoFit() helpers

// "C", "$AB", "AB12" -> the number of the column
STATIC FUNCTION _ColNum( c )
LOCAL nRow := 0, nCol := 0, i, cL := ""
c := Upper( AllTrim( c ) )
FOR i := 1 TO LEN( c )
	IF SubStr( c, i, 1 ) >= "A" .AND. SubStr( c, i, 1 ) <= "Z"
		cL += SubStr( c, i, 1 )
	ENDIF
NEXT
IF Empty( cL )
	RETURN 0
ENDIF
WORKSHEET_RC( cL + "1", @nRow, @nCol )
RETURN nCol

// the text a cell shows: numbers and dates in their format
STATIC FUNCTION _CellText( x, cFmt )
DO CASE                                          // built-in 14 / 22: the short date of Windows
CASE cFmt == "mm-dd-yy"
	cFmt := "dd/mm/yyyy"
CASE cFmt == "m/d/yy h:mm"
	cFmt := "dd/mm/yyyy hh:mm"
ENDCASE
DO CASE
CASE HB_ISSTRING( x )
	RETURN x
CASE HB_ISLOGICAL( x )
	RETURN IIF( x, "TRUE", "FALSE" )
CASE HB_ISDATE( x ) .OR. HB_ISDATETIME( x )
	RETURN _DateText( IIF( Empty( cFmt ) .OR. Upper( cFmt ) == "GENERAL", "dd/mm/yyyy", cFmt ) )
CASE HB_ISNUMERIC( x )
	IF _IsDateFmt( cFmt )
		RETURN _DateText( cFmt )
	ENDIF
	RETURN _NumText( x, cFmt )
ENDCASE
RETURN ""

// a format of dates / times: d m y h s outside quotes, no digit placeholders
STATIC FUNCTION _IsDateFmt( cFmt )
LOCAL c := Lower( _FmtPlain( hb_defaultValue( cFmt, "" ) ) )
RETURN ! Empty( c ) .AND. ! "general" $ c .AND. ! "0" $ c .AND. ! "#" $ c .AND. ;
       ( "d" $ c .OR. "m" $ c .OR. "y" $ c .OR. "h" $ c .OR. "s" $ c )

// the first section of a format without [colors], quoted texts, \x, _x and *x
STATIC FUNCTION _FmtPlain( cFmt, cLit )
LOCAL c := "", i := 1, ch, n
cLit := ""
DO WHILE i <= LEN( cFmt )
	ch := SubStr( cFmt, i, 1 )
	DO CASE
	CASE ch == ";"
		EXIT
	CASE ch == "["
		n := hb_At( "]", cFmt, i )
		i := IIF( n == 0, LEN( cFmt ), n )
	CASE ch == '"'
		n := hb_At( '"', cFmt, i + 1 )
		cLit += SubStr( cFmt, i + 1, IIF( n == 0, LEN( cFmt ), n ) - i - 1 )
		i := IIF( n == 0, LEN( cFmt ), n )
	CASE ch == "\"
		cLit += SubStr( cFmt, i + 1, 1 )
		i++
	CASE ch == "_"
		cLit += " "
		i++
	CASE ch == "*"
		i++
	OTHERWISE
		c += ch
	ENDCASE
	i++
ENDDO
RETURN c

// a number in a number format ("#,##0.00", "0%", "0.00 €" ...), General: at most 11 characters
STATIC FUNCTION _NumText( n, cFmt )
LOCAL aSec, cF, cLit := "", cPlain, nDot, nDec := 0, i, cInt, cDec, lNeg := n < 0, cOther := "", nZeros := 0, c, cNum
IF Empty( cFmt ) .OR. Upper( cFmt ) == "GENERAL"
	c := hb_ntos( n )
	RETURN IIF( LEN( c ) > 11, Left( c, 11 ), c )
ENDIF
aSec := hb_ATokens( cFmt, ";" )
cF := aSec[ 1 ]
IF lNeg .AND. LEN( aSec ) >= 2
	cF := aSec[ 2 ]                             // its own sign: (1,234.00)
	lNeg := .F.
ENDIF
cPlain := _FmtPlain( cF, @cLit )
IF "E+" $ Upper( cPlain ) .OR. "E-" $ Upper( cPlain )
	RETURN Replicate( "0", LEN( cPlain ) ) + cLit
ENDIF
FOR i := 1 TO LEN( cPlain )
	IF SubStr( cPlain, i, 1 ) == "%"
		n *= 100
	ENDIF
NEXT
nDot := At( ".", cPlain )
IF nDot > 0
	FOR i := nDot + 1 TO LEN( cPlain )
		IF SubStr( cPlain, i, 1 ) $ "0#?"
			nDec++
		ELSE
			EXIT
		ENDIF
	NEXT
ENDIF
FOR i := 1 TO IIF( nDot > 0, nDot - 1, LEN( cPlain ) )
	IF SubStr( cPlain, i, 1 ) == "0"
		nZeros++
	ENDIF
NEXT
cNum := LTrim( Str( Abs( n ), 40, nDec ) )
cInt := IIF( nDec > 0, Left( cNum, At( ".", cNum ) - 1 ), cNum )
cDec := IIF( nDec > 0, "." + SubStr( cNum, At( ".", cNum ) + 1 ), "" )
IF cInt == "0" .AND. nZeros == 0
	cInt := ""
ENDIF
IF LEN( cInt ) < nZeros                         // "00000": leading zeros (PadL() would also cut)
	cInt := Replicate( "0", nZeros - LEN( cInt ) ) + cInt
ENDIF
IF "," $ Left( cPlain, IIF( nDot > 0, nDot - 1, LEN( cPlain ) ) )
	c := ""
	FOR i := LEN( cInt ) TO 1 STEP -1
		c := SubStr( cInt, i, 1 ) + IIF( ( LEN( cInt ) - i ) % 3 == 0 .AND. i < LEN( cInt ), ",", "" ) + c
	NEXT
	cInt := c
ENDIF
FOR i := 1 TO LEN( cPlain )                    // the signs and texts of the format
	IF !( SubStr( cPlain, i, 1 ) $ "0#?.," )
		cOther += SubStr( cPlain, i, 1 )
	ENDIF
NEXT
RETURN IIF( lNeg, "-", "" ) + cInt + cDec + cOther + cLit

// a date / time format with its widest values (English month and day names)
STATIC FUNCTION _DateText( cFmt )
LOCAL cLit := "", c := _FmtPlain( cFmt, @cLit ), cL := Lower( c ), cOut := "", i := 1, n, ch, aTok, t
aTok := { { "yyyy", "2026" }, { "yy", "26" }, { "mmmmm", "S" }, { "mmmm", "September" }, { "mmm", "Sep" }, ;
          { "mm", "12" }, { "m", "12" }, { "dddd", "Wednesday" }, { "ddd", "Wed" }, { "dd", "28" }, { "d", "28" }, ;
          { "hh", "23" }, { "h", "23" }, { "ss", "59" }, { "s", "59" }, { "am/pm", "PM" }, { "a/p", "P" } }
DO WHILE i <= LEN( c )
	n := 0
	FOR EACH t IN aTok
		IF SubStr( cL, i, LEN( t[ 1 ] ) ) == t[ 1 ]
			cOut += t[ 2 ]
			n := LEN( t[ 1 ] )
			EXIT
		ENDIF
	NEXT
	IF n == 0
		ch := SubStr( c, i, 1 )
		cOut += ch
		n := 1
	ENDIF
	i += n
ENDDO
RETURN cOut + cLit

// the width of a text in pixels in Calibri 11 (bold: Calibri Bold) - widths of the font files
STATIC FUNCTION _TextPx( c, lBold )
STATIC s_cAsc := "35677::3447744467777777777447777=888977994586=9:8:87798=877464774786874783473<888856587:6765757", s_cAscB := "35677;:3557744467777777777447777=988977994586=::8:877:9=887565774786875784474<888856587;7765757"
STATIC s_cLat := "357777776<6874765755489454689::7888888;87777444499:::::7:9999788777777;677773333888888878888878787878786868686989877777777779797979798:8434343434387548776363648564989898898:8:8:8=<85858576767676757575989898989898=:7777676764", s_cLatB := "357777776<687476575548944468:::7999999;8777744449::::::7:::::888777777;677774444888888878888878797979786868686999877777777779797979798:8444444444497548776464648664:8:8:8998:8:8:8=<85858576767676757575:8:8:8:8:8:8=;8787676764"
STATIC s_cCyr := "779687445==98989888697<79989=9:98878:898==9;88=8787587:68877:88886679687;;8:76;7778576334;;87878", s_cCyrB := "779687445==98:899886:7<7::89=9:98878;8:8=>9<88=8787597;68878:8888667:787;<8:76;7778566444;<87878"
STATIC s_hPun, s_hPunB
LOCAL nPx := 0, i, n, nLen := hb_UTF8Len( c )
IF s_hPun == NIL
	s_hPun := { 0x20AC => 7, 0x2013 => 7, 0x2014 => 13, 0x2018 => 4, 0x2019 => 4, 0x201C => 6, 0x201D => 6, 0x2026 => 10 }
	s_hPunB := { 0x20AC => 7, 0x2013 => 7, 0x2014 => 13, 0x2018 => 4, 0x2019 => 4, 0x201C => 6, 0x201D => 6, 0x2026 => 10 }
ENDIF
FOR i := 1 TO nLen
	n := hb_UTF8Peek( c, i )
	DO CASE
	CASE n >= 32 .AND. n <= 126
		nPx += hb_BCode( hb_BSubStr( IIF( lBold, s_cAscB, s_cAsc ), n - 31, 1 ) ) - 48
	CASE n >= 0xA0 .AND. n <= 0x17F
		nPx += hb_BCode( hb_BSubStr( IIF( lBold, s_cLatB, s_cLat ), n - 0x9F, 1 ) ) - 48
	CASE n >= 0x400 .AND. n <= 0x45F
		nPx += hb_BCode( hb_BSubStr( IIF( lBold, s_cCyrB, s_cCyr ), n - 0x3FF, 1 ) ) - 48
	CASE hb_HHasKey( s_hPun, n )
		nPx += IIF( lBold, s_hPunB[ n ], s_hPun[ n ] )
	CASE ( n >= 0x1100 .AND. n <= 0x115F ) .OR. ( n >= 0x2E80 .AND. n <= 0xA4CF ) .OR. ;
	     ( n >= 0xAC00 .AND. n <= 0xD7A3 ) .OR. ( n >= 0xF900 .AND. n <= 0xFAFF ) .OR. ;
	     ( n >= 0xFE30 .AND. n <= 0xFE4F ) .OR. ( n >= 0xFF00 .AND. n <= 0xFF60 ) .OR. ;
	     ( n >= 0xFFE0 .AND. n <= 0xFFE6 ) .OR. n >= 0x20000
		nPx += 15                                // CJK: a wide character
	OTHERWISE
		nPx += 7
	ENDCASE
NEXT
RETURN nPx

// how much wider than Calibri a font is on average (from the font files)
STATIC FUNCTION _FontFactor( cFont )
LOCAL c := Lower( hb_defaultValue( cFont, "" ) )
DO CASE
CASE "calibri" $ c .OR. Empty( c )
	RETURN 1
CASE "arial" $ c .OR. "helvetica" $ c
	RETURN 1.128
CASE "times" $ c
	RETURN 1.065
CASE "verdana" $ c
	RETURN 1.268
CASE "tahoma" $ c
	RETURN 1.13
CASE "courier" $ c
	RETURN 1.413
CASE "consolas" $ c
	RETURN 1.256
CASE "segoe" $ c
	RETURN 1.111
CASE "georgia" $ c
	RETURN 1.14
ENDCASE
RETURN 1.1                                       // another font: a little wider than Calibri


#pragma BEGINDUMP
#include "hbapi.h"

HB_FUNC( WORKSHEET_RC )
{ 
  const char *cellAddr = hb_parc(1);
  int ii=0, jj, colVal=0;
  if( ! cellAddr )   /* NIL ili ne-string adresa: ranije pad programa */
  {
    hb_storni( 0, 2 );
    hb_storni( 0, 3 );
    return;
  }
  while(cellAddr[ii++] >= 'A') {};
  ii--;
  for(jj=0;jj<ii;jj++) colVal = 26*colVal + HB_TOUPPER(cellAddr[jj]) -'A' + 1;
  hb_storni( atoi(cellAddr+ii), 2 );
  hb_storni( colVal, 3 );
}	

#pragma ENDDUMP
