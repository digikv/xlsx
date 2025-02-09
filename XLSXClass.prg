/*
 *
 * XLSX writer/reader Class - Pure Harbour 
 *
 * Copyright 2018 - 2025 Srđan Dragojlović <digikv@yahoo.com>
 * Company  S.O.K. doo Kraljevo
 * www - http://www.sokdoo.com
*/

#require "hbmzip"
#include "simpleio.ch"
#INCLUDE "hbclass.ch"
#include "hbxml.ch"

REQUEST HB_CODEPAGE_UTF8EX

STATIC aStdNumFormats := { "0" => -1, "0.00" => -2, "#,##0" => -3, "#,##0.00" => -4, "0%" => -9,  "0.00%" => -10, "0.00E+00" => -11, "# ?/?" => -12, "# ??/??" => -13, "m/d/yyyy" => -14, "dd.mm.yyyy" => -14, "dd/mm/yyyy" => -14, "d-mmm-yy" => -15, "d-mmm" => -16, "mmm-yy" => -17, "h:mm AM/PM" => -18, "h:mm:ss AM/PM" => -19, "h:mm" => -20, "h:mm:ss" => -21, "m/d/yyyy h:mm" => -22, "#,##0 ;(#,##0)" => -37, "#,##0 ;Red" => -38, "#,##0.00 ;(#,##0.00)" => -39, "#,##0.00 ;Red" => -40, "mm:ss" => -45, "[h]:mm:ss" => -46, "mm:ss.0" => -47, "##0.00E+0" => -48, "@"  => -49 }
STATIC aHorizontal := {"left","center","right"}
STATIC aVertical := {"top","center","bottom"}
STATIC aBordersStyle := { "thin", "medium", "thick", "double", "dashed", "dotted", "hair", "mediumDashed", "dashDot", "mediumDashDot", "dashDotDot", "mediumDashDotDot", "slantDashDot" }
STATIC aPattern := {"solid", "horzStripe", "vertStripe", "reverseDiagStripe", "diagStripe", "horzCross", "diagCross", "thinHorzStripe", "thinVertStripe", "thinReverseDiagStripe", "thinDiagStripe", "thinHorzCross", "thinDiagCross" }
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
DATA aSharedStrings INIT {}
DATA _str_total INIT 0
METHOD NewFont( cFont, nFontSize, lBold, lItalic, lUnderline, lStrike, cRGB )
METHOD NewFillPattern( nFillPattern, cFG, cBG )
METHOD NewBorder( nTL, nTR, nTT, nTB, nTD, cCL, cCR, cCT, cCB, cCD  )
METHOD NewStyle( nFont, nBorder, nFill, nVA, nHA, nRotation, lWrap )
METHOD NewFormat( cFormat )
METHOD New(cName)
METHOD WorkSheet(cName)
METHOD AddDrawing()
METHOD GetTempDir()
METHOD Save()
METHOD Read(cName)
ENDCLASS

METHOD Read(cName) CLASS WorkBook
LOCAL l := .F., cString, cFileName, oWorkSheet
LOCAL oXmlFieldScan, oXmlDoc, oXmlNode0, oXmlNode, oXmlRecScan, oXmlNodeSheets, oXmlNodeSheet
::cName := cName
my_UnzipFile( cName, ::cTempDir+hb_PS())

cFileName := ::cTempDir+hb_PS()+"xl"+hb_PS()+"sharedStrings.xml"
cString := Memoread( cFileName )
oXmlDoc := TXmlDocument():new(,5)
oXMlDoc:read( cString )
oXmlNode0 := oXmlDoc:findFirst( "sst" )
oXmlRecScan := TXmlIteratorScan():new( oXmlNode0 )
oXmlNode := oXmlRecScan:find( "si" )
WHILE oXmlNode <> NIL
	oXmlFieldScan := TXmlIteratorScan():new( oXmlNode )
	oXmlNode := oXmlFieldScan:find( "t" )
	if oXmlNode != NIL
		if oXmlNode:cData != NIL
			aADD( ::aSharedStrings, oXmlNode:cData )
		endif
	endif
	::_str_total++
	oXmlNode := oXmlRecScan:next("si")
ENDDO

cFileName := ::cTempDir+hb_PS()+"xl"+hb_PS()+"workbook.xml"
cString := Memoread( cFileName )
oXmlDoc := TXmlDocument():new(,5) 
oXMlDoc:read( cString )
oXmlNode0 := oXmlDoc:findFirst( "workbook" )
oXmlRecScan := TXmlIteratorScan():new( oXmlNode0 )
oXmlNodeSheets := oXmlRecScan:find( "sheets" )
oXmlRecScan := TXmlIteratorScan():new( oXmlNodeSheets )
oXmlNodeSheet := oXmlRecScan:find( "sheet" )
WHILE oXmlNodeSheet <> NIL
	oWorkSheet := ::Worksheet(oXmlNodeSheet:aattributes["name"])
	oWorkSheet:ReadWorkSheet( ::cTempDir+hb_PS()+"xl"+hb_PS()+"worksheets"+hb_ps()+"sheet"+oXmlNodeSheet:aattributes["sheetId"]+".xml" )
	oXmlNodeSheet := oXmlRecScan:next( "sheet" )
ENDDO

RETURN l

METHOD GetTempDir() CLASS WorkBook
RETURN ::cTempDir

METHOD AddDrawing() CLASS WorkBook
::nDrawings := ::nDrawings +1
RETURN ::nDrawings

METHOD NewFillPattern( nFillPattern, cFG, cBG ) CLASS WorkBook
LOCAL nPos := 0, c:= ""
IF HB_ISNUMERIC(nFillPattern) .AND. nFillPattern>0 .AND. nFillPattern<14 .AND. HB_ISSTRING( cBG ) .AND. HB_ISSTRING( cFG )
	c := HB_NTOC(nFillPattern,0)+","+cFG+","+cBG
	IF (nPos:=ASCAN( ::aWorkbookFills, c ) )==0
		AADD( ::aWorkbookFills, c )
		nPos := LEN( ::aWorkbookFills )
	ENDIF
ENDIF
Return nPos

METHOD NewBorder( nTL, nTR, nTT, nTB, nTD, cCL, cCR, cCT, cCB, cCD  ) CLASS WorkBook
LOCAL nPos := 0
LOCAL c := ""
c += IIF( (HB_ISNUMERIC(nTL) .AND. nTL>0 .AND. nTL<14), HB_NTOC(nTL,0)+IIF( (!HB_ISNIL(cCL) .AND. HB_ISSTRING( cCL )), ","+cCL+",", ",000000," ),"0,000000," )
c += IIF( (HB_ISNUMERIC(nTR) .AND. nTR>0 .AND. nTR<14), HB_NTOC(nTR,0)+IIF( (!HB_ISNIL(cCR) .AND. HB_ISSTRING( cCR )), ","+cCR+",", ",000000," ),"0,000000," )
c += IIF( (HB_ISNUMERIC(nTT) .AND. nTT>0 .AND. nTT<14), HB_NTOC(nTT,0)+IIF( (!HB_ISNIL(cCT) .AND. HB_ISSTRING( cCT )), ","+cCT+",", ",000000," ),"0,000000," )
c += IIF( (HB_ISNUMERIC(nTB) .AND. nTB>0 .AND. nTB<14), HB_NTOC(nTB,0)+IIF( (!HB_ISNIL(cCB) .AND. HB_ISSTRING( cCB )), ","+cCB+",", ",000000," ),"0,000000," )
c += IIF( (HB_ISNUMERIC(nTD) .AND. nTD>0 .AND. nTD<14), HB_NTOC(nTD,0)+IIF( (!HB_ISNIL(cCD) .AND. HB_ISSTRING( cCD )), ","+cCD+",", ",000000," ),"0,000000" )
IF (nPos:=ASCAN( ::aWorkbookBorders, c ) )==0
	AADD( ::aWorkbookBorders, c )
	nPos := LEN( ::aWorkbookBorders )
ENDIF
Return nPos

METHOD NewFormat( cFormat ) CLASS WorkBook
Local nPos := 0
IF HB_ISSTRING( cFormat )
	IF hb_hHaskey( aStdNumFormats, cFormat )
		nPos := aStdNumFormats[cFormat]
	ELSEIF (nPos:=ASCAN( ::aNumFormat, cFormat ) )==0
		AADD( ::aNumFormat, cFormat )
		nPos := LEN( ::aNumFormat )
	ENDIF
ENDIF
Return nPos

METHOD NewFont( cFont, nFontSize, lBold, lItalic, lUnderline, lStrike, cRGB ) CLASS WorkBook
Local nPos := 0
cFont := iif( HB_ISNIL( cFont ), "Calibri", cFont )
cFont += ","+iif( HB_ISNIL( nFontSize ) .OR. !HB_ISNUMERIC( nFontSize ), "11", HB_NTOC(nFontSize,0) )
cFont += ","+iif( HB_ISNIL( lBold ) .OR. !HB_ISLOGICAL( lBold ), "0", iif(lBold,"1","0") )
cFont += ","+iif( HB_ISNIL( lItalic ) .OR. !HB_ISLOGICAL( lItalic ), "0", iif(lItalic,"1","0") ) 
cFont += ","+iif( HB_ISNIL( lUnderline ) .OR. !HB_ISLOGICAL( lUnderline ), "0", iif(lUnderline,"1","0") )
cFont += ","+iif( HB_ISNIL( lStrike ) .OR. !HB_ISLOGICAL( lStrike ), "0", iif(lStrike,"1","0") )
cFont += ","+iif( HB_ISNIL( cRGB ) .OR. !HB_ISSTRING( cRGB ), "000000", cRGB )
IF (nPos:=ASCAN( ::aWorkbookFonts, cFont ) )==0
	AADD( ::aWorkbookFonts, cFont )
	nPos := LEN( ::aWorkbookFonts )
ENDIF
Return nPos

METHOD NewStyle( nFont, nBorder, nFill, nVA, nHA, nFormat, nRotation, lWrap ) CLASS WorkBook
Local nPos := 0, c := ""
c := iif( HB_ISNIL( nFont ) .OR. !HB_ISNUMERIC( nFont ), "0", HB_NTOC(IIF(nFont==0,1,nFont)-1,0) )
c += ","+iif( HB_ISNIL( nBorder ) .OR. !HB_ISNUMERIC( nBorder ), "0", HB_NTOC(IIF(nBorder==0,1,nBorder)-1,0) )
c += ","+iif( HB_ISNIL( nFill ) .OR. !HB_ISNUMERIC( nFill ), "0", HB_NTOC(IIF(nFill==0,1,nFill)-1,0) )
c += ","+iif( HB_ISNIL( nHA ) .OR. !HB_ISNUMERIC( nHA ) .OR. nHA<1 .OR. nHa>3, "0", HB_NTOC(nHA,0) )
c += ","+iif( HB_ISNIL( nVA ) .OR. !HB_ISNUMERIC( nVA ) .OR. nVA<1 .OR. nVa>3, "0", HB_NTOC(nVA,0) )
c += ","+iif( HB_ISNIL( nFormat ) .OR. !HB_ISNUMERIC( nFormat ), "0", IIF( nFormat>0,HB_NTOS(nFormat+163), HB_NTOS(-nFormat) ) )
c += ","+iif( HB_ISNIL( nRotation ) .OR. !HB_ISNUMERIC( nRotation ), "0", HB_NTOC(nRotation,0) )
c += ","+iif( HB_ISNIL( lWrap ) .OR. !HB_ISLOGICAL( lWrap ), "0", "1" )
IF (nPos:=ASCAN( ::aWorkbookStyles, c ) )==0
	AADD( ::aWorkbookStyles, c )
	nPos := LEN( ::aWorkbookStyles )
ENDIF
Return nPos

METHOD New(cName) CLASS WorkBook
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
::cTempDir := hb_DirSepToOS( hb_DirTemp()+"ExcelTemp" )
hb_DirRemoveAll( ::cTempDir )
MakeDir( ::cTempDir )
Return Self

METHOD WorkSheet( cName ) CLASS WorkBook
LOCAL oWorkSheet, nPos 
IF (nPos:=ASCAN(::aWorkSheetNames, cName))==0
    oWorkSheet:= WorkSheet():New( cName )
	oWorkSheet:oParent := Self
	aADD( ::aWorkSheetNames, cName )
	aADD( ::aWorkSheetObjects, oWorkSheet )
	nPos := LEN( ::aWorkSheetNames )
ENDIF
Return ::aWorkSheetObjects[nPos]

METHOD Save() CLASS WorkBook
LOCAL n := hb_RandomInt( 1, 10000 )
LOCAL cSep := hb_ps(), i, j, handle, c
LOCAL cDate := DTOS( DATE() ), cTime := TIME(), cDateTime := LEFT( cDate, 4 )+'-'+SUBSTR(cDate,5,2)+'-'+RIGHT(cDate,2)+'T'+LEFT(cTime,8)
   MakeDir( ::cTempDir+cSep+"docProps" )
   MakeDir( ::cTempDir+cSep+"_rels" )   
   MakeDir( ::cTempDir+cSep+"xl" )
   MakeDir( ::cTempDir+cSep+"xl"+cSep+"_rels" )
   MakeDir( ::cTempDir+cSep+"xl"+cSep+"worksheets" )
   IF ::nDrawings>0
		MakeDir( ::cTempDir+cSep+"xl"+cSep+"worksheets"+cSep+"_rels" )
		MakeDir( ::cTempDir+cSep+"xl"+cSep+"drawings" )
		MakeDir( ::cTempDir+cSep+"xl"+cSep+"drawings"+cSep+"_rels" )
		MakeDir( ::cTempDir+cSep+"xl"+cSep+"media" )
		aEval( ::aWorkSheetObjects, { |arr| aEval( arr:aDrawingsObjects, { |arr1| ;
			__CopyFile( arr1:cPath, ::cTempDir+cSep+"xl"+cSep+"media"+cSep+"image"+arr1:ID+lower(right(arr1:cPath,4)) ) } ) } )		
   ENDIF
// -------------------------------------------------------- [Content_Types].xml ------------------------------------------------------------------------------   
   handle := FCreate( ::cTempDir+cSep+"[Content_Types].xml" )
   FWrite( handle, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+chr(10) )
   FWrite( handle, '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'+chr(10) )
   IF ::nDrawings>0
        aEval( ::aUsedDrawingType, { |arr| FWrite( handle, '<Default Extension="'+arr+'" ContentType="image/'+arr+'"/>'+chr(10) ) } )
		FWrite( handle, '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'+chr(10) )
		FWrite( handle, '<Default Extension="xml" ContentType="application/xml"/>'+chr(10) )
		FWrite( handle, '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'+chr(10) )		
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
   FWrite( handle, '<Application>'+::Application+'</Application>'+chr(10) )
   FWrite( handle, '<Company>'+::Company+'</Company>'+chr(10) )   
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
	FWrite( handle, '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'+chr(10) )
	FWrite( handle, '<mc:Choice Requires="x15"><x15ac:absPath url="C:\Dropbox\XLSX\" xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac"/></mc:Choice>'+chr(10) )
	FWrite( handle, '</mc:AlternateContent>'+chr(10) )
	FWrite( handle, '<xr:revisionPtr revIDLastSave="0" documentId="8_{063AD3CC-BC4D-49E0-9FE7-D121456950CA}" xr6:coauthVersionLast="47" xr6:coauthVersionMax="47" xr10:uidLastSave="{00000000-0000-0000-0000-000000000000}"/>'+chr(10) )
	FWrite( handle, '<bookViews>'+chr(10) )
	FWrite( handle, '<workbookView xWindow="-108" yWindow="-108" windowWidth="23256" windowHeight="12456" xr2:uid="{00000000-000D-0000-FFFF-FFFF00000000}"/>'+chr(10) )
	FWrite( handle, '</bookViews>'+chr(10) )
	FWrite( handle, '<sheets>'+chr(10) )
	FOR I:=1 TO LEN( ::aWorkSheetNames )
		FWrite( handle, '<sheet name="'+::aWorkSheetNames[i]+'" sheetId="'+HB_NTOS(i)+'" r:id="rId'+HB_NTOS(i)+'"/>'+chr(10) )   
	NEXT I
	FWrite( handle, '</sheets>'+chr(10) )
	FWrite( handle, '<calcPr calcId="191029"/>'+chr(10) )
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
	FWrite( handle, '</workbook>' )   
	FClose( handle )         

   // ------------------------------------------------------------ sharedStrings.xml ------------------------------------------------------------------------------------------------------------------------   
   IF ::_str_total>0
		handle := FCreate( ::cTempDir+cSep+"xl"+cSep+"sharedStrings.xml" )
		FWrite( handle, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+chr(10) )   
        FWrite( handle, '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="'+HB_NTOS(::_str_total)+'" uniqueCount="'+HB_NTOS(LEN(::aSharedStrings))+'">'+chr(10) )
		FOR I:=1 TO LEN(::aSharedStrings)
			FWrite( handle, '<si><t xml:space="preserve">'+::aSharedStrings[I]+'</t></si>'+chr(10) )
		NEXT I
		FWrite( handle, '</sst>' )   
		FClose( handle )         
   ENDIF 		

   // ------------------------------------------------------------ styles.xml ------------------------------------------------------------------------------------------------------------------------   
	handle := FCreate( ::cTempDir+cSep+"xl"+cSep+"styles.xml" )
	FWrite( handle, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+chr(10) )   
    FWrite( handle, '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2 xr" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">'+chr(10) )
	FWrite( handle, '<numFmts count="'+HB_NTOS(LEN(::aNumFormat))+'">'+chr(10))
	FOR I:=1 TO LEN( ::aNumFormat )
		FWrite( handle, '<numFmt numFmtId="'+HB_NTOS(I+163)+'" formatCode="'+::aNumFormat[I]+'"/>'+chr(10) )
	NEXT I
	FWrite( handle, '</numFmts>'+chr(10))
	FWrite( handle, '<fonts count="'+HB_NTOS(LEN(::aWorkbookFonts))+'">'+chr(10) )	
	//FWrite( handle, '<font/>'+chr(10) )
	FOR I:=1 TO LEN( ::aWorkbookFonts )
	    J:= hb_ATokens( ::aWorkbookFonts[i], "," )	
		FWrite( handle, '<font>' )		
		FWrite( handle, '<name val="'+j[1]+'"/>' )
		FWrite( handle, '<sz val="'+j[2]+'"/>' )		
		IF j[3]=="1";FWrite( handle, '<b/>' );ENDIF
		IF j[4]=="1";FWrite( handle, '<i/>' );ENDIF
		IF j[5]=="1";FWrite( handle, '<u/>' );ENDIF
		IF j[6]=="1";FWrite( handle, '<strike/>' );ENDIF		
		FWrite( handle, '<color rgb="'+j[7]+'"/>' )
		FWrite( handle, '</font>'+chr(10) )
	NEXT I
	FWrite( handle, '</fonts>'+chr(10) )		
	FWrite( handle, '<fills count="'+HB_NTOS(LEN(::aWorkbookFills))+'">'+chr(10) )		
	FWrite( handle, '<fill><patternFill patternType="none"/></fill>'+chr(10)+'<fill><patternFill patternType="gray125"/></fill>'+chr(10) )	
	FOR I:=3 TO LEN( ::aWorkbookFills )
	    J:= hb_ATokens( ::aWorkbookFills[i], "," )
		FWrite( handle, '<fill><patternFill patternType="'+aPattern[val(j[1])]+'"><fgColor rgb="'+j[2]+'"/><bgColor rgb="'+j[3]+'"/></patternFill></fill>'+chr(10) )
	NEXT I
	FWrite( handle, '</fills>'+chr(10) )		
	FWrite( handle, '<borders count="'+HB_NTOS(LEN(::aWorkbookBorders))+'">'+chr(10) )		
	FWrite( handle, '<border><left/><right/><top/><bottom/><diagonal/></border>'+chr(10) )		
	FOR I:=2 TO LEN( ::aWorkbookBorders )
	    J:= hb_ATokens( ::aWorkbookBorders[i], "," )
		FWrite( handle, '<border>' )
		FWrite( handle, iif(val(j[1])>0, '<left style="'+aBordersStyle[val(j[1])]+'"'+iif(j[2]!="000000", '><color rgb="'+j[2]+'"/></left>', '/>' ), '<left/>' ) )
		FWrite( handle, iif(val(j[3])>0, '<right style="'+aBordersStyle[val(j[3])]+'"'+iif(j[4]!="000000", '><color rgb="'+j[4]+'"/></right>', '/>' ), '<right/>' ) )
		FWrite( handle, iif(val(j[5])>0, '<top style="'+aBordersStyle[val(j[5])]+'"'+iif(j[6]!="000000", '><color rgb="'+j[6]+'"/></top>', '/>' ), '<top/>' ) )
		FWrite( handle, iif(val(j[7])>0, '<bottom style="'+aBordersStyle[val(j[7])]+'"'+iif(j[8]!="000000", '><color rgb="'+j[8]+'"/></bottom>', '/>' ), '<bottom/>' ) )		
		FWrite( handle, iif(val(j[9])>0, '<diagonal style="'+aBordersStyle[val(j[9])]+'"'+iif(j[10]!="000000", '><color rgb="'+j[10]+'"/></diagonal>', '/>' ), '<diagonal/>' ) )				
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
		FWrite( handle, iif(j[1]=="0",'fontID="0" ','fontId="'+j[1]+'" ') )
		FWrite( handle, iif(j[3]=="0",'fillId="0" ','fillId="'+j[3]+'" ') )
		FWrite( handle, iif(j[2]=="0",'borderId="0" ','borderId="'+j[2]+'" ') )
		FWrite( handle, 'xfId="0" ')
		FWrite( handle, iif(j[6]=="0", '', 'applyNumberFormat="1" ') )				
		FWrite( handle, iif(j[1]=="0", '', 'applyFont="1" ') )
		FWrite( handle, iif(j[3]=="0", '', 'applyFill="1" ') )		
		FWrite( handle, iif(j[2]=="0", '', 'applyBorder="1" ') )
		FWrite( handle, iif(j[4]=="0" .AND. j[5]=="0" .AND. j[7]=="0" .AND. j[8]=="0", '/>', 'applyAlignment="1">' ) )
		IF !(j[4]=="0" .AND. j[5]=="0" .AND. j[7]=="0" .AND. j[8]=="0")
			c := IIF( j[7]!="0", 'textRotation="'+j[7]+'" ', '' )+IIF( j[8]=="0", '', ' wrapText="1"' )
			FWrite( handle, '<alignment horizontal="'+iif(j[4]=="0", "general", aHorizontal[val(j[4])])+'" vertical="'+iif( j[5]=="0", "bottom", aVertical[val(j[5])])+'"'+c+'/>' )
			FWrite( handle, '</xf>'+chr(10) )
		ELSE
			FWrite( handle, chr(10) )
		ENDIF
	NEXT I	
	FWrite( handle, '</cellXfs>'+chr(10) )		
	FWrite( handle, '<cellStyles count="1">'+chr(10) )
	FWrite( handle, '<cellStyle name="Normal" xfId="0" builtinId="0"/>'+chr(10) )
	FWrite( handle, '</cellStyles>'+chr(10) )	
	FWrite( handle, '<dxfs count="0"/>'+chr(10) )
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
	c := DiskName() + hb_OSDriveSeparator() + hb_PS() + CurDir()
	DirChange( ::cTempDir )
	FERASE( ::cFilePath + ::cName )
	myzip( ::cFilePath + ::cName, "*.*", ::cTempDir)
	DirChange( c )
	hb_DirRemoveAll( ::cTempDir )
Return Self

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
METHOD TwoCellAnchor( uAddr1, uAddr2, nSizeX, nSizeY, cPath )
METHOD AbsoluteAnchor( uAddr1, uAddr2, nSizeX, nSizeY, cPath )
METHOD WriteDrawing(nWorkSheet)
ENDCLASS

METHOD WriteDrawing(nWorkSheet) CLASS DrawingML
LOCAL cTempDir := ::oWorkBook:GetTempDir()
LOCAL cSep := hb_ps(), i, handle
	IF LEN(::oWorkSheet:aDrawingsObjects) > 0
		handle := FCreate( cTempDir+cSep+"xl"+cSep+"worksheets"+cSep+"_rels"+cSep+"sheet"+HB_NTOS(nWorkSheet)+".xml.rels" )
		FWrite( handle, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+chr(10) )
		FWrite( handle, '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+chr(10) )
		FOR I=1 TO LEN(::oWorkSheet:aDrawingsObjects)
			FWrite( handle, '<Relationship Id="rId'+::oWorkSheet:aDrawingsObjects[I]:ID+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing'+HB_NTOS(nWorkSheet)+'.xml"/>'+CHR(10))
		NEXT I
		FWrite( handle, '</Relationships>' )	
		FClose( handle )   				

		handle := FCreate( cTempDir+cSep+"xl"+cSep+"drawings"+cSep+"_rels"+cSep+"drawing"+HB_NTOS(nWorkSheet)+".xml.rels" )
		FWrite( handle, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+chr(10) )   
		FWrite( handle, '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+chr(10) )			
		aEval( ::oWorkSheet:aDrawingsObjects, { |arr1| FWrite( handle, '<Relationship Id="rId'+arr1:ID+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image'+arr1:ID+lower(right(arr1:cPath,4))+'" />'+chr(10)  ) } ) 
		FWrite( handle, '</Relationships>' )   
		FClose( handle )         		
		
		handle := FCreate( cTempDir+cSep+"xl"+cSep+"drawings"+cSep+"drawing"+HB_NTOS(nWorkSheet)+".xml" )		
		FWrite( handle, '<?xml version="1.0" encoding="utf-8"?>'+chr(10) )   
		FWrite( handle, '<xdr:wsDr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">'+chr(10) )			
		FOR I:=1 TO LEN(::oWorkSheet:aDrawingsObjects)
			IF ::Type == 0
				FWrite( handle, '<xdr:absoluteAnchor>'+chr(10) )
				FWrite( handle, chr(9)+'<xdr:pos x="'+::nX+'" y="'+::nY+'" />'+chr(10) )
				FWrite( handle, chr(9)+'<xdr:ext cx="'+::nSizeX+'" cy="'+::nSizeY+'" />'+chr(10) )
				FWrite( handle, chr(9)+'<xdr:pic>'+chr(10) )
				FWrite( handle, chr(9)+'<xdr:nvPicPr>'+chr(10) )
				FWrite( handle, chr(9)+chr(9)+'<xdr:cNvPr id="'+HB_NTOS(I-1)+'" name="'+::oWorkSheet:aDrawingsObjects[i]:cName+'" descr=""></xdr:cNvPr>'+chr(10) )
				FWrite( handle, chr(9)+chr(9)+'<xdr:cNvPicPr/>'+chr(10) )
				FWrite( handle, chr(9)+'</xdr:nvPicPr>'+chr(10) )				
				FWrite( handle, chr(9)+chr(9)+'<xdr:blipFill><a:blip r:embed="rId'+::oWorkSheet:aDrawingsObjects[i]:ID+'" cstate="print" /><a:stretch/></xdr:blipFill>'+chr(10) )
				FWrite( handle, chr(9)+chr(9)+'<xdr:spPr>'+chr(10) )
				FWrite( handle, chr(9)+chr(9)+chr(9)+'<a:xfrm>'+chr(10) )
                FWrite( handle, chr(9)+chr(9)+chr(9)+chr(9)+'<a:off x="0" y="0"/>'+chr(10) )
                FWrite( handle, chr(9)+chr(9)+chr(9)+chr(9)+'<a:ext cx="1971360" cy="783720"/>'+chr(10) )
				FWrite( handle, chr(9)+chr(9)+chr(9)+'</a:xfrm>'+chr(10) )
				FWrite( handle, chr(9)+chr(9)+chr(9)+'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'+chr(10) )
				FWrite( handle, chr(9)+chr(9)+chr(9)+'<a:ln><a:noFill/></a:ln>'+chr(10) )
				FWrite( handle, chr(9)+chr(9)+'</xdr:spPr>'+chr(10) )				
				FWrite( handle, chr(9)+'</xdr:pic>'+chr(10) )
				FWrite( handle, '<xdr:clientData />'+chr(10) )
				FWrite( handle, '</xdr:absoluteAnchor>'+chr(10) )
				FWrite( handle, '</xdr:wsDr>')
			ELSEIF ::Type == 1
				FWrite( handle, '<xdr:oneCellAnchor>'+chr(10) )
				FWrite( handle, chr(9)+'<xdr:from>'+chr(10) )
				FWrite( handle, chr(9)+chr(9)+'<xdr:col>'+::nCol1+'</xdr:col>'+chr(10) )
				FWrite( handle, chr(9)+chr(9)+'<xdr:colOff>'+::nOffColFrom+'</xdr:colOff>'+chr(10) )
				FWrite( handle, chr(9)+chr(9)+'<xdr:row>'+::nRow1+'</xdr:row>'+chr(10) )
				FWrite( handle, chr(9)+chr(9)+'<xdr:rowOff>'+::nOffRowFrom+'</xdr:rowOff>'+chr(10) )
				FWrite( handle, chr(9)+'</xdr:from>'+chr(10) )
				FWrite( handle, chr(9)+'<xdr:ext cx="'+::nSizeX+'" cy="'+::nSizeY+'" />'+chr(10) )
				FWrite( handle, chr(9)+'<xdr:pic>'+chr(10) )
				FWrite( handle, chr(9)+'<xdr:nvPicPr>'+chr(10) )
				FWrite( handle, chr(9)+chr(9)+'<xdr:cNvPr id="'+HB_NTOS(I-1)+'" name="'+::oWorkSheet:aDrawingsObjects[i]:cName+'" descr=""></xdr:cNvPr>'+chr(10) )
				FWrite( handle, chr(9)+chr(9)+'<xdr:cNvPicPr/>'+chr(10) )
				FWrite( handle, chr(9)+'</xdr:nvPicPr>'+chr(10) )				
				FWrite( handle, chr(9)+chr(9)+'<xdr:blipFill><a:blip r:embed="rId'+::oWorkSheet:aDrawingsObjects[i]:ID+'" cstate="print" /><a:stretch/></xdr:blipFill>'+chr(10) )
				FWrite( handle, chr(9)+chr(9)+'<xdr:spPr>'+chr(10) )
				FWrite( handle, chr(9)+chr(9)+chr(9)+'<a:xfrm>'+chr(10) )
                FWrite( handle, chr(9)+chr(9)+chr(9)+chr(9)+'<a:off x="0" y="0"/>'+chr(10) )
                FWrite( handle, chr(9)+chr(9)+chr(9)+chr(9)+'<a:ext cx="1971360" cy="783720"/>'+chr(10) )
				FWrite( handle, chr(9)+chr(9)+chr(9)+'</a:xfrm>'+chr(10) )
				FWrite( handle, chr(9)+chr(9)+chr(9)+'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'+chr(10) )
				FWrite( handle, chr(9)+chr(9)+chr(9)+'<a:ln><a:noFill/></a:ln>'+chr(10) )
				FWrite( handle, chr(9)+chr(9)+'</xdr:spPr>'+chr(10) )				
				FWrite( handle, chr(9)+'</xdr:pic>'+chr(10) )
				FWrite( handle, '<xdr:clientData />'+chr(10) )
				FWrite( handle, '</xdr:oneCellAnchor>'+chr(10) )
				FWrite( handle, '</xdr:wsDr>')								
			ELSEIF ::Type == 2
				FWrite( handle, '<xdr:twoCellAnchor editAs="'+::editAs+'">'+chr(10) )
				FWrite( handle, chr(9)+'<xdr:from>'+chr(10) )
				FWrite( handle, chr(9)+chr(9)+'<xdr:col>'+::nCol1+'</xdr:col>'+chr(10) )
				FWrite( handle, chr(9)+chr(9)+'<xdr:colOff>'+::nOffColFrom+'</xdr:colOff>'+chr(10) )
				FWrite( handle, chr(9)+chr(9)+'<xdr:row>'+::nRow1+'</xdr:row>'+chr(10) )
				FWrite( handle, chr(9)+chr(9)+'<xdr:rowOff>'+::nOffRowFrom+'</xdr:rowOff>'+chr(10) )
				FWrite( handle, chr(9)+'</xdr:from>'+chr(10) )
				FWrite( handle, chr(9)+'<xdr:to>'+chr(10) )
				FWrite( handle, chr(9)+chr(9)+'<xdr:col>'+::nCol2+'</xdr:col>'+chr(10) )
				FWrite( handle, chr(9)+chr(9)+'<xdr:colOff>'+::nOffColTo+'</xdr:colOff>'+chr(10) )
				FWrite( handle, chr(9)+chr(9)+'<xdr:row>'+::nRow2+'</xdr:row>'+chr(10) )
				FWrite( handle, chr(9)+chr(9)+'<xdr:rowOff>'+::nOffRowTo+'</xdr:rowOff>'+chr(10) )
				FWrite( handle, chr(9)+'</xdr:to>'+chr(10) )
				FWrite( handle, chr(9)+'<xdr:pic>'+chr(10) )
				FWrite( handle, chr(9)+'<xdr:nvPicPr>'+chr(10) )
				FWrite( handle, chr(9)+chr(9)+'<xdr:cNvPr id="'+HB_NTOS(I-1)+'" name="'+::oWorkSheet:aDrawingsObjects[i]:cName+'" descr=""></xdr:cNvPr>'+chr(10) )
				FWrite( handle, chr(9)+chr(9)+'<xdr:cNvPicPr/>'+chr(10) )
				FWrite( handle, chr(9)+'</xdr:nvPicPr>'+chr(10) )				
				FWrite( handle, chr(9)+chr(9)+'<xdr:blipFill><a:blip r:embed="rId'+::oWorkSheet:aDrawingsObjects[i]:ID+'" cstate="print" /><a:stretch/></xdr:blipFill>'+chr(10) )
				FWrite( handle, chr(9)+chr(9)+'<xdr:spPr>'+chr(10) )
				FWrite( handle, chr(9)+chr(9)+chr(9)+'<a:xfrm>'+chr(10) )
                FWrite( handle, chr(9)+chr(9)+chr(9)+chr(9)+'<a:off x="0" y="0"/>'+chr(10) )
                FWrite( handle, chr(9)+chr(9)+chr(9)+chr(9)+'<a:ext cx="1971360" cy="783720"/>'+chr(10) )
				FWrite( handle, chr(9)+chr(9)+chr(9)+'</a:xfrm>'+chr(10) )
				FWrite( handle, chr(9)+chr(9)+chr(9)+'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'+chr(10) )
				FWrite( handle, chr(9)+chr(9)+chr(9)+'<a:ln><a:noFill/></a:ln>'+chr(10) )
				FWrite( handle, chr(9)+chr(9)+'</xdr:spPr>'+chr(10) )				
				FWrite( handle, chr(9)+'</xdr:pic>'+chr(10) )
				FWrite( handle, '<xdr:clientData />'+chr(10) )
				FWrite( handle, '</xdr:twoCellAnchor>'+chr(10) )
				FWrite( handle, '</xdr:wsDr>')			
			ENDIF
		NEXT I
		FClose( handle )         				
	ENDIF

Return self

METHOD New(cName) CLASS DrawingML
::cName := cName
Return Self

METHOD AbsoluteAnchor( nX, nY, nSizeX, nSizeY, cPath ) CLASS DrawingML
LOCAL cType := lower( RIGHT(ALLTRIM(cPath), 3 ) )
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

METHOD OneCellAnchor( uAddr1, nSizeX, nSizeY, cPath ) CLASS DrawingML
Local nPos := 0, nRow, nCol
LOCAL cType := lower( RIGHT(ALLTRIM(cPath), 3 ) )
::Type := 1
IF HB_ISARRAY( uAddr1 )
	::nRow1 := HB_NTOS(uAddr1[1])
	::nCol1 := HB_NTOS(uAddr1[2])
ELSE
    nRow := 0
	nCol := 0
	WORKSHEET_RC( uAddr1, @nRow, @nCol )
	::nRow1 := HB_NTOS(nRow-1)
	::nCol1 := HB_NTOS(nCol-1)	
ENDIF
::cPath := cPath
::nSizeX := HB_NTOS(nSizeX)
::nSizeY := HB_NTOS(nSizeY)
IF ASCAN( ::oWorkBook:aUsedDrawingType, {|aVal| aVal == cType } ) == 0
    aADD( ::oWorkBook:aUsedDrawingType, cType )
ENDIF
Return Self


METHOD TwoCellAnchor( uAddr1, uAddr2, cPath ) CLASS DrawingML
Local nPos := 0, nRow, nCol
LOCAL cType := lower( RIGHT(ALLTRIM(cPath), 3 ) )
::Type := 2
IF HB_ISARRAY( uAddr1 )
	::nRow1 := uAddr1[1]
	::nCol1 := uAddr1[2]
ELSE
    nRow := 0
	nCol := 0
	WORKSHEET_RC( uAddr1, @nRow, @nCol )
	::nRow1 := HB_NTOS(nRow-1)
	::nCol1 := HB_NTOS(nCol-1)	
ENDIF
IF HB_ISARRAY( uAddr2 )
	::nRow2 := uAddr2[1]
	::nCol2 := uAddr2[2]
ELSE
    nRow := 0
	nCol := 0
	WORKSHEET_RC( uAddr2, @nRow, @nCol )
	::nRow2 := HB_NTOS(nRow)
	::nCol2 := HB_NTOS(nCol)	
ENDIF
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
METHOD New(cName)
METHOD WriteWorksheet(n, cPath)
METHOD Cell(uAddr, nValue, nStyle)
METHOD WriteCell(uAddr, nValue, nStyle)
METHOD ReadCell(uAddr)
METHOD MergeCell(uAddr)
METHOD ColumnsWidth( nMin, nLast, nWidth )
METHOD AddHeader( cLeft, cCenter, cRight )
METHOD AddFooter( cLeft, cCenter, cRight )
METHOD RowDetail( nRow, nHeight, nStyle, lHide )
//METHOD AddChart( cName )
METHOD Drawing( cName )
METHOD ReadWorkSheet( cXML ) 
ENDCLASS

METHOD Drawing( cName ) CLASS WorkSheet
LOCAL oDrawing, nPos 
IF (nPos:=ASCAN(::aDrawingsNames, cName))==0
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
Return Self


METHOD RowDetail( nRow, nHeight, nStyle, lHide ) CLASS WorkSheet
LOCAL i
IF HB_ISNUMERIC(nRow)
    IF LEN(::aRowsDetail)<nRow
		FOR I=LEN(::aRowsDetail) TO nRow
			aAdd( ::aRowsDetail, { 0, 0, .F. } )
		NEXT I
	ENDIF
	::aRowsDetail[nRow] := { IIF(HB_ISNIL(nHeight),0,IIF(HB_ISNUMERIC(nHeight),nHeight,0)), IIF(HB_ISNIL(nStyle),0,IIF(HB_ISNUMERIC(nStyle),nStyle,0)), IIF(HB_ISNIL(lHide),.F.,IIF(HB_ISLOGICAL(lHide),lHide,.F.)) } 
ENDIF
RETURN Self

METHOD AddChart( uAddr, oChart, nX, nY, nX_Scale, nY_Scale ) CLASS WorkSheet
RETURN Self

METHOD MergeCell( uAddr ) CLASS WorkSheet
LOCAL nCol1, nCol2, nRow1, nRow2, adr, I, J, K
IF HB_ISSTRING(uAddr) .AND. ASCAN( ::aMergeCells, uAddr )==0
	AADD(  ::aMergeCells, uAddr )
	adr := hb_ATokens( uAddr, ":" )
	WORKSHEET_RC( adr[1], @nRow1, @nCol1 )
	WORKSHEET_RC( adr[2], @nRow2, @nCol2 )
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
   ::aHeader[1] := ReplaceAmp(cLeft)
ENDIF
IF HB_ISSTRING( cCenter )
   ::aHeader[2] := ReplaceAmp(cCenter)
ENDIF
IF HB_ISSTRING( cRight )
   ::aHeader[3] := ReplaceAmp(cRight)
ENDIF
::lHeader := .T.
RETURN Self

METHOD AddFooter( cLeft, cCenter, cRight )
IF HB_ISSTRING( cLeft )
   ::aFooter[1] := ReplaceAmp(cLeft)
ENDIF
IF HB_ISSTRING( cCenter )
   ::aFooter[2] := ReplaceAmp(cCenter)
ENDIF
IF HB_ISSTRING( cRight )
   ::aFooter[3] := ReplaceAmp(cRight)
ENDIF
::lFooter := .T.
RETURN Self

METHOD ColumnsWidth( nMin, nMax, nWidth ) CLASS WorkSheet
IF HB_ISNUMERIC(nMin) .AND. HB_ISNUMERIC(nMax) .AND. HB_ISNUMERIC(nWidth)
	AADD( ::aCols, { nMin, nMax, nWidth } )
ENDIF
Return Self

METHOD Cell( uAddr, xValue, nStyle ) CLASS WorkSheet
LOCAL l 
    IF xValue == NIL
		l:= ::ReadCell( uAddr )
	ELSE
		::WriteCell( uAddr, xValue, nStyle ) 
	ENDIF	
Return l

METHOD ReadCell( uAddr ) CLASS WorkSheet
LOCAL nRow := 0, nCol := 0
IF HB_ISARRAY( uAddr )
	nRow := uAddr[1]
	nCol := uAddr[2]
ELSE
	WORKSHEET_RC( uAddr, @nRow, @nCol )
ENDIF
Return ::aData[nRow,nCol]

METHOD WriteCell( uAddr, xValue, nStyle ) CLASS WorkSheet
LOCAL nCol := 0, nRow := 0, lIsMerge := .F., nRow1 := 0, nCol1 := 0, nRow2 := 0, nCol2 := 0
LOCAL nPos, i, j, k, l
IF HB_ISARRAY( uAddr )
	nRow := uAddr[1]
	nCol := uAddr[2]
ELSE
	WORKSHEET_RC( uAddr, @nRow, @nCol )
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
	nPos := AT( "&", xValue )
	xValue := IIF( nPos>0, LEFT(xValue,nPos)+"amp;"+SUBSTR(xValue,nPos+1), xValue )
ENDIF
IF !HB_ISNIL(xValue)
	::aData[nRow,nCol] := xValue
	IF HB_ISSTRING(xValue) .AND. LEFT(xValue,1) != "="
	    ::oParent:_str_total++
		i := ASCAN( ::oParent:aSharedStrings, xValue )
		IF i==0
			AADD( ::oParent:aSharedStrings, xValue )
		Else 
			If !(::oParent:aSharedStrings[i] == xValue )
				AADD( ::oParent:aSharedStrings, xValue )
			EndIf		
		ENDIF
	ENDIF
ENDIF
IF HB_ISNIL(nStyle) .AND. HB_ISDATE(xValue)
	nStyle := ::oParent:NewStyle( , , , , ,::oParent:NewFormat("dd/mm/yyyy"))
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
LOCAL v, i, j, c, x, handle, ht, cLastStyle := ''
handle := FCreate( cPath+"sheet"+HB_NTOS(n)+".xml" )
FWrite( handle, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+chr(10) )   
FWrite( handle, '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac xr xr2 xr3" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" xr:uid="{00000000-0001-0000-0000-000000000000}">'+chr(10) )
FWrite( handle, '<sheetPr><pageSetUpPr fitToPage="'+iif(::fitToPage,"1","0")+'"/></sheetPr>'+chr(10) )
FWrite( handle, '<sheetFormatPr defaultRowHeight="15"/>'+chr(10) )
IF LEN( ::aCols ) > 0
	FWrite( handle, '<cols>'+chr(10) )
	FOR I=1 TO LEN( ::aCols )
		FWrite( handle, '<col min="'+HB_NTOS(::aCols[I,1])+'" max="'+HB_NTOS(::aCols[I,2])+'" width="'+HB_NTOC(::aCols[I,3],8)+'" bestFit="1" customWidth="1"/>'+chr(10) ) 
	NEXT I
	FWrite( handle, '</cols>'+chr(10) )
ENDIF
FWrite( handle, '<sheetData>'+chr(10) )
FOR I:=1 TO LEN( ::aData )
    cLastStyle := ''
    ht:=iif(LEN(::aRowsDetail)>=I,iif(::aRowsDetail[i,1]>0,' customHeight="1" ht="'+HB_NTOS(::aRowsDetail[i,1])+'"',""),"")
	IF LEN(::aData[I])>0 
	    FOR J:=1 TO LEN(::aData[I])
		    IF !HB_ISNIL(::aData[I,J])
				c := J
				J := LEN(::aData[I]) + 1
            ENDIF
        NEXT J 		
		FWrite( handle, '<row r="'+HB_NTOS(I)+'" spans="'+HB_NTOS(c)+':'+HB_NTOS(LEN(::aData[I]))+'"'+ht+'>'+chr(10) )
		FOR J:=1 TO LEN(::aData[I])
			c := ColumnIndexToColumnLetter(J)+HB_NTOS(I)
			IF !HB_ISNIL(::aData[I,J])
				IF !(LEN( ::aStyle ) < I) .AND. !(LEN(::aStyle[I])<J) .AND. (HB_ISSTRING(::aStyle[I,J]))
					cLastStyle := ::aStyle[I,J]
				ENDIF
				IF HB_ISNUMERIC(::aData[I,J])
                    v := HB_NTOC(::aData[I,J]) 
					IF !(LEN( ::aStyle ) < I) .AND. !(LEN(::aStyle[I])<J) .AND. (HB_ISSTRING(::aStyle[I,J]))
						FWrite( handle, '<c r="'+c+'" s="'+::aStyle[I,J]+'"><v>'+v+'</v></c>'+chr(10) )
					ELSE
						FWrite( handle, '<c r="'+c+'"><v>'+v+'</v></c>'+chr(10) )
					ENDIF
				ELSEIF HB_ISSTRING(::aData[I,J])
					IF LEFT( ::aData[I,J], 1 ) == "="
						v := SUBSTR(::aData[I,J],2)
						IF !(LEN( ::aStyle ) < I) .AND. !(LEN(::aStyle[I])<J) .AND. (HB_ISSTRING(::aStyle[I,J]))
							FWrite( handle, '<c r="'+c+'" t="str" s="'+::aStyle[I,J]+'"><f>'+v+'</f></c>'+chr(10) )
						ELSE
							FWrite( handle, '<c r="'+c+'" t="str"><f>'+v+'</f></c>'+chr(10) )
						ENDIF
					ELSEIF LEFT( ::aData[I,J], 1 ) == CHR(0)
						FWrite( handle, iif( HB_ISNIL(cLastStyle) .OR. cLastStyle=='','','<c r="'+c+ '" s="'+cLastStyle+'"></c>'+chr(10) ) )
					ELSE
						Set exact on 
						v := ALLTRIM( STR( ASCAN( ::oParent:aSharedStrings, ::aData[I,J] )-1, 10, 0 ) )
						IF !(LEN( ::aStyle ) < I) .AND. !(LEN(::aStyle[I])<J) .AND. (HB_ISSTRING(::aStyle[I,J]))
							FWrite( handle, '<c r="'+c+'" t="s" s="'+::aStyle[I,J]+'"><v>'+v+'</v></c>'+chr(10) )
						ELSE
							FWrite( handle, '<c r="'+c+'" t="s"><v>'+v+'</v></c>'+chr(10) )
						ENDIF
						Set exact off 
					ENDIF
				ELSEIF HB_ISLOGICAL(::aData[I,J])						
                                        v := IIF( ::aData[I,J], "1", "0" )
					IF !(LEN( ::aStyle ) < I) .AND. !(LEN(::aStyle[I])<J) .AND. (HB_ISSTRING(::aStyle[I,J]))
						FWrite( handle, '<c r="' +c+ '" t="b" s="'+::aStyle[I,J]+'"><v>' +v+ '</v></c>'+chr(10) )
					ELSE
						FWrite( handle, '<c r="' +c+ '" t="b"><v>' +v+ '</v></c>'+chr(10) )							
					ENDIF
				ELSEIF HB_ISDATE(::aData[I,J])	
                    v := ALLTRIM( STR(::aData[I,J]-HB_STOD( "19000101" )+2,10,0) )
					FWrite( handle, '<c r="' +c+ '" s="'+::aStyle[I,J]+'"><v>' +v+ '</v></c>'+chr(10) )
				ENDIF
			ENDIF
		NEXT J
		FWrite( handle, '</row>'+chr(10) )
	ENDIF
NEXT I
FWrite( handle, '</sheetData>'+chr(10) )
IF LEN( ::aMergeCells ) > 0
	FWrite( handle, '<mergeCells>'+chr(10) )
	FOR I:=1 TO LEN( ::aMergeCells )
		FWrite( handle, '<mergeCell ref="'+::aMergeCells[I]+'"/>'+chr(10) )
	NEXT I
	FWrite( handle, '</mergeCells>'+chr(10) )
ENDIF
FWrite( handle, '<pageMargins left="'+HB_NTOC(::leftMargin)+'" right="'+HB_NTOC(::rightMargin)+'" top="'+HB_NTOC(::topMargin,2)+'" bottom="'+HB_NTOC(::bottomMargin,2)+'" header="'+HB_NTOC(::headerMargin,2)+'" footer="'+HB_NTOC(::footerMargin,2)+'"/>'+chr(10) )
FWrite( handle, '<pageSetup paperSize="'+HB_NTOC(::PaperSize,0)+'" horizontalDpi="'+HB_NTOC(::horizontalDpi,2)+'" verticalDpi="'+HB_NTOC(::verticalDpi,2)+'"'+iif(::lLandscape,' orientation="landscape"',' orientation="portrait"')+'/>'+chr(10) )
IF ::lHeader .OR. ::lFooter
	FWrite( handle, '<headerFooter>' )
	IF ::lHeader
		FWrite( handle, '<oddHeader>' )
		IF LEN(::aHeader[1])>0
			FWrite( handle, '&amp;L'+::aHeader[1] )
		ENDIF
		IF LEN(::aHeader[2])>0
			FWrite( handle, '&amp;C'+::aHeader[2] )
		ENDIF
		IF LEN(::aHeader[3])>0
			FWrite( handle, '&amp;R'+::aHeader[3] )
		ENDIF		
		FWrite( handle, '</oddHeader>' )
	ENDIF
	IF ::lFooter
		FWrite( handle, '<oddFooter>' )
		IF LEN(::aFooter[1])>0
			FWrite( handle, '&amp;L'+::aFooter[1] )
		ENDIF
		IF LEN(::aFooter[2])>0
			FWrite( handle, '&amp;C'+::aFooter[2] )
		ENDIF
		IF LEN(::aFooter[3])>0
			FWrite( handle, '&amp;R'+::aFooter[3] )
		ENDIF		
		FWrite( handle, '</oddFooter>' )
	ENDIF	
	FWrite( handle, '</headerFooter>'+chr(10) )
ENDIF
FOR I:=1 TO LEN( ::aDrawingsObjects )
FWrite( handle, '<drawing r:id="rId'+::aDrawingsObjects[I]:ID+'"/>'+chr(10))
::aDrawingsObjects[I]:WriteDrawing(n, cPath )
NEXT I
FWrite( handle, '</worksheet>' )   
FClose( handle )         
Return Self

METHOD ReadWorkSheet( cFileName ) CLASS WorkSheet
	LOCAL i
	LOCAL cString := Memoread( cFileName )
	LOCAL oXmlDoc := TXmlDocument():new(,5)
	LOCAL oXmlNode0, oXmlRecScan, oXmlNode, oRow, oCell, oXmlRecScan1  
	LOCAL nRow, nCol, uAddr, cType, cStyle, xData
	oXMlDoc:read( cString )
	oXmlNode0 := oXmlDoc:findFirst( "worksheet" )
	oXmlRecScan := TXmlIteratorScan():new( oXmlNode0 )
	oXmlNode := oXmlRecScan:find( "sheetData" )
	oXmlRecScan := TXmlIteratorScan():new( oXmlNode )
	oRow := oXmlRecScan:find( "row" )
	WHILE oRow != NIL
		oXmlRecScan1 := TXmlIteratorScan():new( oRow )
		nRow := oRow:aattributes["r"]
		oCell := oXmlRecScan1:find( "c" )
		cStyle := ""
		WHILE oCell != NIL
			uAddr := oCell:aattributes["r"]
			WORKSHEET_RC( uAddr, @nRow, @nCol )
			::nMaxCol := IIF( nCol > ::nMaxCol, nCol, ::nMaxCol )
			::nMaxRow := IIF( nRow > ::nMaxRow, nRow, ::nMaxRow )
			IF LEN( ::aData ) < nRow
				FOR i := LEN(::aData)+1 TO nRow 
					AADD( ::aData, {} )
					AADD( ::aStyle, {} )
				NEXT I
			ENDIF
			IF LEN( ::aData[nRow] ) < nCol
				FOR i := LEN(::aData[nRow])+1 TO nCol
					AADD( ::aData[nRow], NIL )
					AADD( ::aStyle[nRow], NIL )
				NEXT I	   
			ENDIF		
			cType := iif( hb_hHaskey( oCell:aattributes, "t" ), oCell:aattributes["t"], "n" )
			if hb_hHaskey( oCell:aattributes, "s" )
				::aStyle[nRow,nCol] := VAL(oCell:aattributes["s"])
			ENDIF		
			if oCell:oChild != NIL
				xData := oCell:oChild:cData 
				::aData[nRow,nCol] := iif(cType=="n", VAL(xData), iif( cType=="b", iif(VAL(xData)==0,.F.,.T.), iif(cType=="str","=","")+iif(cType=="s", ::oParent:aSharedStrings[VAL(xData)+1], xData ) ) )
			ENDIF	
			oCell := oXmlRecScan1:next( "c" )
		END
		oRow := oXmlRecScan:next( "row" )
	END
RETURN Self
	
STATIC Function ColumnIndexToColumnLetter(colIndex)
LOCAL div := colIndex
LOCAL colLetter := ""
LOCAL modnum := 0
While div > 0
    modnum := MOD((div - 1),26)
    colLetter := Chr(65 + modnum) + colLetter
    div := Int((div - modnum) / 26)
End 
Return colLetter

STATIC PROCEDURE MyZip( cName, cWild, cTempPath )
   LOCAL cSep := hb_ps()
   LOCAL hZip, aDir, aFile, cZipName, cPath, cFileName, cExt, lUnicode := .T.
   hb_FNameSplit( cName, @cPath, @cFileName, @cExt )
   cZipName := hb_FNameMerge( cPath, cFileName, cExt )
   
	hZip := hb_zipOpen( cZipName )
	IF ! Empty( hZip )
        IF ! Empty( cWild )
            hb_FNameSplit( cWild, @cPath, @cFileName, @cExt )
            aDir := hb_DirScan( cPath, cFileName + cExt )
            FOR EACH aFile IN aDir
               IF ! cPath + aFile[ 1 ] == cZipName
                  hb_zipStoreFile( hZip, cPath + aFile[ 1 ], cPath + aFile[ 1 ],,, lUnicode )
               ENDIF
            NEXT
        ENDIF
        hb_zipStoreFile( hZip, cTempPath+cSep+"_rels"+cSep+".rels", "_rels"+cSep+".rels",,, lUnicode )
	hb_zipClose( hZip )
	ENDIF

RETURN

STATIC FUNCTION ReplaceAmp(xValue)
LOCAL nPos, c:= ""

WHILE (nPos := AT( "&", xValue ))>0
	c += LEFT(xValue,nPos)+"amp;"
	xValue := SUBSTR(xValue,nPos+1)	
END

c+=xValue
RETURN c

STATIC FUNCTION my_UnZipFile( cFile, cPath )
  LOCAL nI
  LOCAL aFiles   := { }
  LOCAL aFolders := { }
  aFiles := hb_GetFilesInZip( cFile )
  FOR nI := 1 TO Len( aFiles )
    IF ! HB_DirExists( HB_DirSepToOS( HB_FNameDir( aFiles[nI] ) ) )
      HB_DirCreate( HB_DirSepAdd( cPath ) + HB_DirSepToOS( HB_FNameDir( aFiles[nI] ) ) )
    ENDIF
  NEXT
  HB_UnZipFile( cFile, NIL, NIL, NIL, cPath )
  HB_UnZipFile( cFile, ;
                NIL, ;
                .T., ;
                NIL, ;
                cPath, ;
                aFiles )
RETURN NIL

STATIC FUNCTION FilePath( cFile )
LOCAL nPos, cFilePath := iif((nPos := RAT(hb_PS(), cFile)) != 0,SUBSTR(cFile, 1, nPos),"")
RETURN cFilePath

#pragma BEGINDUMP
#include "hbapi.h"

HB_FUNC( WORKSHEET_RC )
{ 
  const char *cellAddr = hb_parc(1);
  int ii=0, jj, colVal=0;
  while(cellAddr[ii++] >= 'A') {};
  ii--;
  for(jj=0;jj<ii;jj++) colVal = 26*colVal + HB_TOUPPER(cellAddr[jj]) -'A' + 1;
  hb_storni( atoi(cellAddr+ii), 2 );
  hb_storni( colVal, 3 );
}	

#pragma ENDDUMP
