'option explicit
USES PICT
dim aImages

'Printer Color Mode
'Constant Value Description 
Const vbPRCMMonochrome = 1 'Monochrome output 
Const vbPRCMColor = 2 'Color output 


'Duplex Printing
'Constant Value Description 
Const vbPRDPSimplex = 1 'Single-sided printing 
Const vbPRDPHorizontal = 2 'Double-sided horizontal printing 
Const vbPRDPVertical = 3 'Double-sided vertical printing 


'Printer Orientation
'Constant Value Description 
Const vbPRORPortrait = 1 'Documents print with the top at the narrow side of the paper 
Const vbPRORLandscape = 2 'Documents print with the top at the wide side of the paper 


'Print Quality
'Constant Value Description 
Const vbPRPQDraft = -1 'Draft print quality 
Const vbPRPQLow = -2 'Low print quality 
Const vbPRPQMedium = -3 'Medium print quality 
Const vbPRPQHigh = -4 'High print quality 


'PaperBin Property
'Constant Value Description 
Const vbPRBNUpper = 1 'Use paper from the upper bin 
Const vbPRBNLower = 2 'Use paper from the lower bin 
Const vbPRBNMiddle = 3 'Use paper from the middle bin 
Const vbPRBNManual = 4 'Wait for manual insertion of each sheet of paper 
Const vbPRBNEnvelope = 5 'Use envelopes from the envelope feeder 
Const vbPRBNEnvManual = 6 'Use envelopes from the envelope feeder, but wait for manual insertion 
Const vbPRBNAuto = 7 '(Default) Use paper from the current default bin 
Const vbPRBNTractor = 8 'Use paper fed from the tractor feeder 
Const vbPRBNSmallFmt = 9 'Use paper from the small paper feeder 
Const vbPRBNLargeFmt = 10 'Use paper from the large paper bin 
Const vbPRBNLargeCapacity = 11 'Use paper from the large capacity feeder 
Const vbPRBNCassette = 14 'Use paper from the attached cassette cartridge 


'PaperSize Property
'Constant Value Description 
Const vbPRPSLetter = 1 'Letter, 8 1/2 x 11 in 
Const vbPRPSLetterSmall = 2 '+A611Letter Small, 8 1/2 x 11 in 
Const vbPRPSTabloid = 3 'Tabloid, 11 x 17 in 
Const vbPRPSLedger = 4 'Ledger, 17 x 11 in 
Const vbPRPSLegal = 5 'Legal, 8 1/2 x 14 in 
Const vbPRPSStatement = 6 'Statement, 5 1/2 x 8 1/2 in 
Const vbPRPSExecutive = 7 'Executive, 7 1/2 x 10 1/2 in 
Const vbPRPSA3 = 8 'A3, 297 x 420 mm 
Const vbPRPSA4 = 9 'A4, 210 x 297 mm 
Const vbPRPSA4Small = 10 'A4 Small, 210 x 297 mm 
Const vbPRPSA5 = 11 'A5, 148 x 210 mm 
Const vbPRPSB4 = 12 'B4, 250 x 354 mm 
Const vbPRPSB5 = 13 'B5, 182 x 257 mm 
Const vbPRPSFolio = 14 'Folio, 8 1/2 x 13 in 
Const vbPRPSQuarto = 15 'Quarto, 215 x 275 mm 
Const vbPRPS10x14 = 16 '10 x 14 in 
Const vbPRPS11x17 = 17 '11 x 17 in 
Const vbPRPSNote = 18 'Note, 8 1/2 x 11 in 
Const vbPRPSEnv9 = 19 'Envelope #9, 3 7/8 x 8 7/8 in 
Const vbPRPSEnv10 = 20 'Envelope #10, 4 1/8 x 9 1/2 in 
Const vbPRPSEnv11 = 21 'Envelope #11, 4 1/2 x 10 3/8 in 
Const vbPRPSEnv12 = 22 'Envelope #12, 4 1/2 x 11 in 
Const vbPRPSEnv14 = 23 'Envelope #14, 5 x 11 1/2 in 
Const vbPRPSCSheet = 24 'C size sheet 
Const vbPRPSDSheet = 25 'D size sheet 
Const vbPRPSESheet = 26 'E size sheet 
Const vbPRPSEnvDL = 27 'Envelope DL, 110 x 220 mm 
Const vbPRPSEnvC3 = 29 'Envelope C3, 324 x 458 mm 
Const vbPRPSEnvC4 = 30 'Envelope C4, 229 x 324 mm 
Const vbPRPSEnvC5 = 28 'Envelope C5, 162 x 229 mm 
Const vbPRPSEnvC6 = 31 'Envelope C6, 114 x 162 mm 
Const vbPRPSEnvC65 = 32 'Envelope C65, 114 x 229 mm 
Const vbPRPSEnvB4 = 33 'Envelope B4, 250 x 353 mm 
Const vbPRPSEnvB5 = 34 'Envelope B5, 176 x 250 mm 
Const vbPRPSEnvB6 = 35 'Envelope B6, 176 x 125 mm 
Const vbPRPSEnvItaly = 36 'Envelope, 110 x 230 mm 
Const vbPRPSEnvMonarch = 37 'Envelope Monarch, 3 7/8 x 7 1/2 in 
Const vbPRPSEnvPersonal = 38 'Envelope, 3 5/8 x 6 1/2 in 
Const vbPRPSFanfoldUS = 39 'U.S. Standard Fanfold, 14 7/8 x 11 in 
Const vbPRPSFanfoldStdGerman = 40 'German Standard Fanfold, 8 1/2 x 12 in 
Const vbPRPSFanfoldLglGerman = 41 'German Legal Fanfold, 8 1/2 x 13 in 
Const vbPRPSUser = 256 'User-defined 

dim nLeftMargin
dim nTopMargin

'~ Trace.tracing = true
'Trace.tracefile = "ext.log"
'Trace.cleartrace

function InArrayAt( sStr, aArr )
	dim i
	dim j
	j = -1
	for i = LBound( aArr ) To UBound( aArr )
		if aArr( i ) = sStr then
			j = i
			exit for
		end if
	next
	InArrayAt = j
end function

function cm(n)
  cm = int( n * 567 )
end function

sub SETLEFTMARGIN( nTwips )
  nLeftMargin = nTwips
end sub

sub SETTOPMARGIN( nTwips )
  nTopMargin = nTwips
end sub

sub ATSAY( nX, nY, sText )
  dim nSaveX, nSaveY
  nSaveX = Printer.currentx
  nSaveY = Printer.currenty
  Printer.currentx = nx
  Printer.currenty = ny
  text.writetext sText
  Printer.currentx = nsavex
  Printer.currenty = nsavey
end sub

sub ATPICTURE( nX, nY, sFilename )
  Picture.show sFilename, nx, ny
end sub

sub ATPICTURELIMITED( nX1, nY1, nx2, ny2, sFilename )
  Pict.constrained sFilename, nx1, ny1, nx2, ny2
end sub

sub ATPICTUREWH( nX1, nY1, nW, nH, sFilename )
	'~ dim currH, currW
	'~ Picture.loadfile sFilename 
	'~ currH = Picture.Highness
	'~ currW = Picture.Wideness
	'~ Trace.trace "currH=" & currH
	'~ Trace.trace "currW=" & currW
	'~ Trace.trace "sheight=" & sheight
	'~ Trace.trace "swidth=" & swidth
	'~ Trace.trace "pheight=" & pheight
	'~ Trace.trace "pwidth=" & pwidth
	'~ do while currH > nHeight
		'~ currH = currH - Picture.sheight
		'~ currW = currW - Picture.swidth
	'~ loop
	'~ Trace.trace "currH=" & currH
	'~ Trace.trace "currW=" & currW
	'~ Trace.trace "Picture.constrained " & sFilename & ", " & nx1 & ", " & ny1 & ", " & currH & ", " & currW
	'Picture.constrained sFilename, nx1, ny1, nx1 + currH, ny1 + currW
	trace "ATPICTUREWH(" & sFilename & ", " & nx1 & ", " & ny1 & "," & nw & ", " & nh & ")"
	Pict.scaled sFilename, nx1, ny1, nW, nH
end sub

sub ATLINE( x1, y1, h1, w1 )
	Text.WriteLine x1, y1, x1 + h1, y1 + w1
end sub

sub ATSAYCENTRED( nLine, sText )
  dim nSaveY
  dim nSaveX
  nSaveY = Printer.currenty
  nSaveX = Printer.currentX
  Text.CentreText ntopmargin + nLine, sText
  Printer.currenty = nsavey
  Printer.currentX = nsaveX
end sub

sub ATSAYLEFTALIGNED( nX, nLine, sText )
  dim nSaveY
  dim nSaveX
  nSaveY = Printer.currenty
  nSaveX = Printer.currentX
  Printer.currentx = nx
  Printer.currenty = ntopmargin + nLine
  Text.WriteText sText
  Printer.currenty = nsavey
  Printer.currentX = nsaveX
end sub

sub ATSAYRIGHTALIGNED( nLine, sText )
  dim nSaveY
  nSaveY = Printer.currenty
  Printer.currenty = ny
  Text.RightAlignText ntopmargin + nLine,sText
  Printer.currenty = nsavey
end sub

sub ABSOLUTEBOX( x1, y1, x2, y2 )
	dim nSaveX
	dim nSaveY
	nSaveX = Printer.currentx
	nSaveY = Printer.currenty
	Text.WriteLine x1, y1, x2, y2
	Printer.currentx = nsaveX
	Printer.currenty = nsaveY
end sub

sub BOX( x1, y1, h1, w1 )
	dim nSaveX
	dim nSaveY
	nSaveX = Printer.currentx
	nSaveY = Printer.currenty
	Text.WriteLine x1, y1, x1 + h1, y1 + w1
	Printer.currentx = nsaveX
	Printer.currenty = nsaveY
end sub

sub BOXNamed( x1, y1, h1, w1, s1 )
	dim nSaveX
	dim nSaveY
	nSaveX = Printer.currentx
	nSaveY = Printer.currenty
	atline x1, y1, h1, w1 'Text.WriteLine x1, y1, x1 + h1, y1 + w1
	dim thisFont
	set thisFont = Printer.font
	setfont "Arial", 11, array("-bold")
	AtSay x1 + cm(0.75), y1 + cm(0.25), s1
	set Printer.font = thisFont
	Printer.currentx = nsaveX
	Printer.currenty = nsaveY
end sub

sub BOXNamed2( x1, y1, h1, w1, s1 )
	dim nSaveX
	dim nSaveY
	nSaveX = Printer.currentx
	nSaveY = Printer.currenty
	atline x1, y1, h1, w1 'Text.WriteLine x1, y1, x1 + h1, y1 + w1
	dim thisFont
	set thisFont = Printer.font
	setfont "Arial", 11, array("-bold")
	AtSay x1 + cm(0.75), y1 + cm(0.05), s1
	set Printer.font = thisFont
	Printer.currentx = nsaveX
	Printer.currenty = nsaveY
end sub

sub verdana14Bold()
	Font.Name = "Verdana": Font.Size = 14: Font.Bold = True
	Set Printer.Font = Font
end sub

sub verdana12()
	Font.Name = "Verdana": Font.Size = 12: Font.Bold = False
	Set Printer.Font = Font
end sub

sub setfont( sName, nSize, aStyles )
	dim vStyle
	font.name = sName
	font.size = nSize
	for each vStyle in aStyles
		select case lcase( vStyle )
		case "+bold"
			font.bold = true
		case "-bold"
			font.bold = false
		case "+italic"
			font.italic = true
		case "-italic"
			font.italic = false
		end select
	next
	set Printer.font = font
end sub

sub BOX4( x1, y1  )
	dim nSaveX
	dim nSaveY
	nSaveX = Printer.currentx
	nSaveY = Printer.currenty
	
	'Dayname box: 2cm wide x .3cm high
	Box x1, y1, cm(2), cm(0.3)
	
	'Timeslices top row: 2 of (1cm wide x .2 high)
	Box x1, y1 + cm(0.3), cm(1), cm(0.2)
 	Box x1 + cm(1), y1 + cm(0.3), cm(1), cm(0.2)
	
	'Signoffs: top row: 2 of (1cm wide x 1 cm high)
	Box x1, y1 + cm(0.5), cm(1), cm(1)
	Box x1 + cm(1), y1 + cm(0.5), cm(1), cm(1)
	
	'Timeslices bottom row: 2 of (1cm wide x .2 high)
	Box x1, y1 + cm(1.5), cm(1), cm(0.2)
 	Box x1 + cm(1), y1 + cm(1.5), cm(1), cm(0.2)
	
	'Signoffs: bottom row: 2 of (1cm wide x 1 cm high)
	Box x1, y1 + cm(1.7), cm(1), cm(1)
	Box x1 + cm(1), y1 + cm(1.7), cm(1), cm(1)
	
	Printer.currentx = nsaveX
	Printer.currenty = nsaveY
end sub

sub BOX4T( x1, y1, s1 )
	dim nSaveX
	dim nSaveY
	nSaveX = Printer.currentx
	nSaveY = Printer.currenty
	
	'Dayname box: 2cm wide x .3cm high
	Box x1, y1, cm(2), cm(0.3)
	setfont "Arial", 6, array("+bold")
	AtSay x1 + cm(0.75), y1 + cm(0.05), s1
	
	'Timeslices top row: 2 of (1cm wide x .2 high)
	Box x1, y1 + cm(0.3), cm(1), cm(0.2)
 	Box x1 + cm(1), y1 + cm(0.3), cm(1), cm(0.2)
	
	'Signoffs: top row: 2 of (1cm wide x 1 cm high)
	Box x1, y1 + cm(0.5), cm(1), cm(1)
	Box x1 + cm(1), y1 + cm(0.5), cm(1), cm(1)
	
	'Timeslices bottom row: 2 of (1cm wide x .2 high)
	Box x1, y1 + cm(1.5), cm(1), cm(0.2)
 	Box x1 + cm(1), y1 + cm(1.5), cm(1), cm(0.2)
	
	'Signoffs: bottom row: 2 of (1cm wide x 1 cm high)
	Box x1, y1 + cm(1.7), cm(1), cm(1)
	Box x1 + cm(1), y1 + cm(1.7), cm(1), cm(1)
	
	Printer.currentx = nsaveX
	Printer.currenty = nsaveY
end sub

sub BOX4TT( x1, y1, s1, a1 )
	dim nSaveX
	dim nSaveY
	nSaveX = Printer.currentx
	nSaveY = Printer.currenty
	
	'Dayname box: 2cm wide x .3cm high
	Box x1, y1, cm(2), cm(0.3)
	setfont "Arial", 6, array("+bold")
	AtSay x1 + cm(0.75), y1 + cm(0.05), s1
	
	'Timeslices top row: 2 of (1cm wide x .2 high)
	Box x1, y1 + cm(0.3), cm(1), cm(0.2)
		setfont "arial", 4, array("-bold")
		atsay x1 + cm(.03), y1 + cm(0.35), a1(0)
 	Box x1 + cm(1), y1 + cm(0.3), cm(1), cm(0.2)
		setfont "arial", 4, array("-bold")
		atsay x1 + cm(1.03), y1 + cm(0.35), a1(1)
 	
	
	'Signoffs: top row: 2 of (1cm wide x 1 cm high)
	Box x1, y1 + cm(0.5), cm(1), cm(1)
	Box x1 + cm(1), y1 + cm(0.5), cm(1), cm(1)
	
	'Timeslices bottom row: 2 of (1cm wide x .2 high)
	Box x1, y1 + cm(1.5), cm(1), cm(0.2)
		setfont "arial", 4, array("-bold")
		atsay x1 + cm(.03), y1 + cm(1.55), a1(2)
 	Box x1 + cm(1), y1 + cm(1.5), cm(1), cm(0.2)
		setfont "arial", 4, array("-bold")
		atsay x1 + cm(1.03), y1 + cm(1.55), a1(3)
	
	'Signoffs: bottom row: 2 of (1cm wide x 1 cm high)
	Box x1, y1 + cm(1.7), cm(1), cm(1)
	Box x1 + cm(1), y1 + cm(1.7), cm(1), cm(1)
	
	Printer.currentx = nsaveX
	Printer.currenty = nsaveY
end sub

'~ '----------------------------- other functions
'~ Function FirstWord( sText )
    '~ Dim nSpace 
    '~ Dim sTemp
    '~ sTemp = LTrim( sText )
    '~ nSpace = Instr( sTemp, " " )
    '~ IF nSpace > 0 Then
        '~ sTemp = Left( sTemp, nSpace - 1 )
    '~ End If
    '~ FirstWord = sTemp
'~ End Function

'~ Function LastWord( sText )
    '~ Dim nSpace 
    '~ Dim sTemp
    '~ sTemp = LTrim( sText )
    '~ nSpace = InstrRev( sTemp, " " )
    '~ IF nSpace > 0 Then
        '~ sTemp = Mid( sTemp, nSpace + 1 )
    '~ End If
    '~ LastWord = sTemp
'~ End Function

'~ Function BeginsWith( sText, sBeginning, bCaseInsensitive )
    '~ Dim bResult

    '~ bResult = False
    '~ If sBeginning = vbNullString Then
        '~ bResult = True
    '~ Else
        '~ If sText <> vbNullString Then
            '~ If bCaseInsensitive = True Then
                '~ bResult = ( Left( UCase( sText ), Len( sBeginning ) ) = UCase( sBeginning ) )
            '~ Else
                '~ bResult = ( Left( sText, Len( sBeginning ) ) = sBeginning )
            '~ End If
        '~ End If
    '~ End If

    '~ BeginsWith = bResult
'~ End Function

'~ Function EndsWith( sText, sEnding, bCaseInsensitive )
    '~ Dim bResult

    '~ bResult = False
    '~ If sEnding = vbNullString Then
        '~ bResult = True
    '~ Else
        '~ If sText <> vbNullString Then
            '~ If bCaseInsensitive = True Then
                '~ bResult = ( Right( UCase( sText ), Len( sEnding ) ) = UCase( sEnding ) )
            '~ Else
                '~ bResult = ( Right( sText, Len( sEnding ) ) = sEnding )
            '~ End If
        '~ End If
    '~ End If

    '~ EndsWith = bResult
'~ End Function

'~ Function ReadFile( FileName)
    '~ Dim oStream
    '~ Dim sData
    '~ Dim oFSO
    '~ set oFSO = createobject( "Scripting.FileSystemObject" )
    '~ sData = vbNullString

    '~ Set oStream = oFSO.OpenTextFile(FileName )
	'~ On Error Resume Next
        '~ sData = oStream.ReadAll
	'~ If Err.Number <> 0 Then
		'~ sData = vbNullString
	'~ End If
	'~ On Error GoTo 0
    '~ oStream.Close

    '~ ReadFile = sData
'~ End Function

Function LoadImages()
	dim sIDs 
	dim aLine
	dim n
	n = 0
	sIDs = files.ReadFilee( ".\wpid.cfg" )
	aImages = Split( sIDs, vbNewLine )
	for i = 0 to ubound( aImages )
		if aImages(i) <> vbNullString then
			aLine = Split( aImages( i ), "=" )
			aImages( i ) = aLine
			n = n + 1
		end if
	next
	LoadImages = n
end Function

Function GetImage( sLeft, sRight )
	dim i
	dim sResult
	dim sName
	dim sImage
	sResult = vbNullString
	for i = 0 to UBound( aImages ) - 1
		sName = aImages(i)(0)
		If strings.BeginsWith( sName, sLeft, True ) and strings.EndsWith( sName, sRight, True ) then
			sResult = aImages(i)(1)
			exit for
		end if
	next
	GetImage = sResult
End Function

sub ShowBox( sTitle, sMsg )
	dim pHgt, pWid
	pHgt = 1000
	pWid = 3000 + panel.textwidth(sMsg)
	panel.top = screen.height / 2 - (pHgt / 2) 
	panel.left = screen.width / 2 - (pWid / 2)
	panel.height = pHgt
	panel.width = pWid
	panel.caption = sTitle
	panel.borderstyle = 0
	panel.windowstate = 0
	panel.visible = true
	text.pointtopanel
	panel.font.name = "Arial"
	panel.font.size= 14
	text.writetext sMsg
	dim i
	for i = 1 to 2
		system.cooperate
		if panel.visible = false then exit for
		system.sleep 1000
	next
	panel.visible = false
	panel.enabled = false
	text.pointtoprinter
end sub

setleftmargin 0
settopmargin 0
