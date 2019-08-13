Attribute VB_Name = "UScript"
Dim Scr As MSScriptControl.ScriptControl
Dim ARR As Arrays
Dim oArrays As Arrays
Dim BIG As BigString
Dim CLP As Clipboards
Dim CMD As CmdLine
Dim COL As Collections
Dim CPX As Complexes
Dim CMP As Computers
Dim CFG As Config
Dim DBF As Database
Dim DTM As DateTime
Dim Dec As Decimals
Dim DOM As Domains
Dim DRV As Drives
Dim EML As Emails
Dim ENV As Environments
Dim XCL As Excel
Dim Fil As Files
Dim oFiles As Files
Dim GRP As Groups
Dim HLP As Help
'Dim IED As IEDisplay
Dim INI As INI1
Dim IN2 As INI2
Dim oINI2 As INI2
Dim oINI3 As INI3
Dim IN3 As INI3
Dim LGC As Logic
Dim MTH As MathParser
Dim MSC As Misc
Dim MSE As Mouse
Dim NUM As Number
Dim PCT As Pict
Dim pic As Picture
'Dim PRO As PROTIUM
Dim reg As Registry
Dim SCH As Schedule
Dim SRT As Sort
Dim SWT As StopWatch
Dim str As Strings
Dim oStrings As Strings
Dim SYM As SymbolTable
Dim SYS As System
Dim oSystem As System
Dim Tmp As temp
Dim TST As Testing
Dim TXT As Text
Dim TRC As traces
Dim oTrace As traces
Dim UNI As Unicode
Dim USR As Users
Dim VAR As Variable
Dim WAL As Wallpaper

Dim DLG As CommonDialog
Dim FSO As Scripting.FileSystemObject
Dim oFSO As Scripting.FileSystemObject
Dim PRN As Printer
Dim FNT As StdFont
Dim RGX As RegExp

Sub Main()
    
    Dim sData As String
    Dim aData() As String
    Dim sLine As Variant
    
    Dim aCmdLine As Variant

    Dim sFilename As String
    Dim sCodeBlob As String
           
    Set oTrace = New traces
    oTrace.Tracing = True
    oTrace.ClearTrace
    
    If App.StartMode = vbSModeAutomation Then
        
    Else
        Set oStrings = New Strings
        Set oFiles = New Files
        Set oSystem = New System
        Set oFSO = New Scripting.FileSystemObject
        Set oINI2 = New INI2
        Set oINI3 = New INI3
        Set oArrays = New Arrays
        sCodeBlob = vbNullString
        aCmdLine = oStrings.SplitString(Trim(Command()))
        'oTrace.Trace oArrays.ShowStructure(aCmdLine)
        sFilename = aCmdLine(0)
        'sFilename = oStrings.FirstWord(Command())
        
        'sFilename = cmd.Argument(1) 'Command()
        If sFilename = vbNullString Then
            oSystem.ConsoleWriteLine "U298 <script file>"
            End
        End If
        
        sFilename = oStrings.ForceExtension(sFilename, ".298")
        
        If Dir(sFilename) = vbNullString Then
            oSystem.ConsoleWriteLine sFilename & " not found."
            End
        End If
        
        
        'Load Preview
        
        Set Scr = New MSScriptControl.ScriptControl
        Scr.Language = "VBScript"
        Scr.AllowUI = True
        Scr.Timeout = 60000
        
        'Set PRN = Printer
        'Set TXT = New Text
        'Set PCT = New Pict
        'TXT.Init PRN
        'PCT.Init PRN
        
'        Dim sIni As String
'        sIni = ".\u298.ini"
'        If oFSO.FileExists(sIni) Then
'            oINI2.Load sIni
'            If Not oINI2.HasSection(LCase$(sFilename)) Then
'                GoTo everything
'            End If
'            oINI3.Filename = sIni
'            oINI3.Section = LCase$(sFilename)
'            oINI3.Key = "Application"
'            'Debug.Print INI3.Content
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Scr.AddObject "Application", App, True
'            End If
'
'            oINI3.Key = "Arrays"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set ARR = New Arrays
'                Scr.AddObject "Arrays", ARR, True
'            End If
'
'            oINI3.Key = "BigString"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set BIG = New BigString
'                Scr.AddObject "BigString", BIG, True
'            End If
'
'            oINI3.Key = "Clipboard"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set CLP = New Clipboards
'                Scr.AddObject "Clipboard", CLP, True
'            End If
'
'            oINI3.Key = "CmdLine"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set CMD = New CmdLine
'                Scr.AddObject "CmdLine", CMD, True
'            End If
'
'            oINI3.Key = "Complex"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set CPX = New Complexes
'                Scr.AddObject "Complex", CPX, True
'            End If
'
'            oINI3.Key = "Computers"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set CMP = New Computers
'                Scr.AddObject "Computer", CMP, True
'            End If
'
'            oINI3.Key = "Config"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set CFG = New Config
'                Scr.AddObject "Config", CFG, True
'            End If
'
'            oINI3.Key = "Database"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set DBF = New Database
'                Scr.AddObject "Database", DBF, True
'            End If
'
'            oINI3.Key = "DateTime"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set DTM = New DateTime
'                Scr.AddObject "DateTime", DTM, True
'            End If
'
'            oINI3.Key = "Decimal"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set Dec = New Decimals
'                Scr.AddObject "Decimal", Dec, True
'            End If
'
'            oINI3.Key = "Dialogs"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set DLG = New CommonDialog
'                Scr.AddObject "Dialogs", DLG, True
'            End If
'
'            oINI3.Key = "Domains"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set DOM = New Domains
'                Scr.AddObject "Domain", DOM, True
'            End If
'
'            oINI3.Key = "Drives"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set DRV = New Drives
'                Scr.AddObject "Drive", DRV, True
'            End If
'
'            oINI3.Key = "Email"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set EML = New Emails
'                Scr.AddObject "Email", EML, True
'            End If
'
'            oINI3.Key = "Environment"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set ENV = New Environments
'                Scr.AddObject "Environment", ENV, True
'            End If
'
'            oINI3.Key = "Excel"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set XCL = New Excel
'                Scr.AddObject "Excel", XCL, True
'            End If
'
'            oINI3.Key = "Files"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set Fil = New Files
'                Scr.AddObject "Files", Fil, True
'            End If
'
'            oINI3.Key = "FSO"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Scr.AddObject "FSO", FSO, True
'            End If
'
'            oINI3.Key = "Groups"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set GRP = New Groups
'                Scr.AddObject "Group", GRP, True
'            End If
'
'            oINI3.Key = "Help"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set HLP = New Help
'                Scr.AddObject "Help", HLP, True
'            End If
'
'            oINI3.Key = "INI1"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set INI = New INI1
'                Scr.AddObject "INI", INI, True
'            End If
'
'            oINI3.Key = "INI2"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set IN2 = New INI2
'                Scr.AddObject "INI2", IN2, True
'            End If
'
'            oINI3.Key = "INI3"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set IN3 = New INI3
'                Scr.AddObject "INI3", IN3, True
'            End If
'
'            oINI3.Key = "Logic"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set LGC = New Logic
'                Scr.AddObject "Logic", LGC, True
'            End If
'
'            oINI3.Key = "MathParser"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set MTH = New MathParser
'                Scr.AddObject "MathParser", MTH, True
'            End If
'
'            oINI3.Key = "Misc"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set MSC = New Misc
'                Scr.AddObject "Misc", MSC, True
'            End If
'
'            oINI3.Key = "Mouse"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set MSE = New Mouse
'                Scr.AddObject "Mouse", MSE, True
'            End If
'
'            oINI3.Key = "Number"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set NUM = New Number
'                Scr.AddObject "Number", NUM, True
'            End If
'
'            oINI3.Key = "Panel"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Scr.AddObject "Panel", Form1, True
'            End If
'
'            oINI3.Key = "Pict"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set PCT = New Pict
'                Scr.AddObject "Pict", PCT, True
'            End If
'
'            oINI3.Key = "Picture"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set pic = New Picture
'                Scr.AddObject "Picture", pic, True
'            End If
'
'            oINI3.Key = "Printer"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set PRN = Printer
'                Scr.AddObject "Printer", PRN, True
'            End If
'
''            oINI3.Key = "Protium"
''            If LCase$(left$(INI3.Value, 1)) = "y" Then
''                Set PRO = New PROTIUM
''                Scr.AddObject "Protium", PRO, True
''            End If
''
'            oINI3.Key = "RegExp"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set RGX = New RegExp
'                Scr.AddObject "RegExp", RGX, True
'            End If
'
'            oINI3.Key = "Registry"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set reg = New Registry
'                Scr.AddObject "Registry", reg, True
'            End If
'
'            oINI3.Key = "Schedule"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set SCH = New Schedule
'                Scr.AddObject "Schedule", SCH, True
'            End If
'
'            oINI3.Key = "Screen"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Scr.AddObject "Screen", Screen, True
'            End If
'
'            oINI3.Key = "Sort"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set SRT = New Sort
'                Scr.AddObject "Sort", SRT, True
'            End If
'
'            oINI3.Key = "Font"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set FNT = New StdFont
'                Scr.AddObject "Font", FNT, True
'            End If
'
'            oINI3.Key = "StopWatch"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set SWT = New StopWatch
'                Scr.AddObject "StopWatch", SWT, True
'            End If
'
'            oINI3.Key = "Strings"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set str = New Strings
'                Scr.AddObject "Strings", str, True
'            End If
'
'            oINI3.Key = "SymbolTable"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set SYM = New SymbolTable
'                Scr.AddObject "SymbolTable", SYM, True
'            End If
'
'            oINI3.Key = "System"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set SYS = New System
'                Scr.AddObject "System", SYS, True
'            End If
'
'            oINI3.Key = "Temp"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set Tmp = New temp
'                Scr.AddObject "Temp", Tmp, True
'            End If
'
'            oINI3.Key = "Testing"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set TST = New Testing
'                Scr.AddObject "Testing", TST, True
'            End If
'
'            oINI3.Key = "Text"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set TXT = New Text
'                TXT.Init PRN
'                Scr.AddObject "Text", TXT, True
'            End If
'
'            oINI3.Key = "Trace"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set TRC = New traces
'                Scr.AddObject "Trace", TRC, True
'            End If
'
'            oINI3.Key = "Unicode"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set UNI = New Unicode
'                Scr.AddObject "Unicode", UNI, True
'            End If
'
'            oINI3.Key = "Users"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set USR = New Users
'                Scr.AddObject "Users", USR, True
'            End If
'
'            oINI3.Key = "Variable"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set VAR = New Variable
'                Scr.AddObject "Variable", VAR, True
'            End If
'
'            oINI3.Key = "Wallpaper"
'            If LCase$(left$(INI3.Value, 1)) = "y" Then
'                Set WAL = New Wallpaper
'                Scr.AddObject "Wallpaper", WAL, True
'            End If
'
'
'        Else
'everything:
            'instantiate everything
'            Set ARR = New Arrays
'            Set BIG = New BigString
'            Set CLP = New Clipboards
'            Set CMD = New CmdLine
'            Set COL = New Collections
'            Set CPX = New Complexes
'            Set CMP = New Computers
'            Set CFG = New Config
'            Set DBF = New Database
'            Set DTM = New DateTime
'            Set Dec = New Decimals
'            Set DOM = New Domains
'            Set DRV = New Drives
'            Set EML = New Emails
'            Set ENV = New Environments
'            Set XCL = New Excel
'            Set Fil = New Files
'            Set GRP = New Groups
'            Set HLP = New Help
'            Set IED = New IEDisplay
'            Set INI = New INI1
'            Set IN2 = New INI2
'            Set IN3 = New INI3
'            Set LGC = New Logic
'            Set MTH = New MathParser
'            Set MSC = New Misc
'            Set MSE = New Mouse
'            Set NUM = New Number
'            Set PCT = New Pict
'            Set pic = New Picture
'            Set PRO = New PROTIUM
'            Set REG = New Registry
'            Set SCH = New Schedule
'            Set SRT = New Sort
'            Set SWT = New StopWatch
'            Set str = New Strings
'            Set SYM = New SymbolTable
'            Set SYS = New System
'            Set Tmp = New temp
'            Set TST = New Testing
'            Set TXT = New Text
'            Set TRC = New traces
'            Set UNI = New Unicode
'            Set USR = New Users
'            Set VAR = New Variable
'            Set WAL = New Wallpaper
'            Set DLG = New CommonDialog
'            Set PRN = Printer
'            Set FNT = New StdFont
'            Set RGX = New RegExp
'            Scr.AddObject "Array", ARR, True
'            Scr.AddObject "BigString", BIG, True
'            Scr.AddObject "Clipboard", CLP, True
'            Scr.AddObject "Cmdline", CMD, True
'            Scr.AddObject "Collection", COL, True
'            Scr.AddObject "Complex", CPX, True
'            Scr.AddObject "Computer", CMP, True
'            Scr.AddObject "Config", CFG, True
'            Scr.AddObject "Database", DBF, True
'            Scr.AddObject "DateTime", DTM, True
'            Scr.AddObject "Decimal", Dec, True
'            Scr.AddObject "Domain", DOM, True
'            Scr.AddObject "Drive", DRV, True
'            Scr.AddObject "Email", EML, True
'            Scr.AddObject "Environment", ENV, True
'            Scr.AddObject "Excel", XCL, True
'            Scr.AddObject "Files", Fil, True
'            Scr.AddObject "Group", GRP, True
'            Scr.AddObject "Help", HLP, True
'            'Scr.AddObject "IEDisplay", IED, True
'            Scr.AddObject "INI", INI, True
'            Scr.AddObject "INI2", IN2, True
'            Scr.AddObject "INI3", IN3, True
'            Scr.AddObject "Logic", LGC, True
'            Scr.AddObject "MathParser", MTH, True
'            Scr.AddObject "Misc", MSC, True
'            Scr.AddObject "Mouse", MSE, True
'            Scr.AddObject "Number", NUM, True
'            Scr.AddObject "Pict", PCT, True
'            Scr.AddObject "Picture", pic, True
'            Scr.AddObject "Protium", PRO, True
'            Scr.AddObject "Registry", REG, True
'            Scr.AddObject "Schedule", SCH, True
'            Scr.AddObject "Sort", SRT, True
'            Scr.AddObject "StopWatch", SWT, True
'            Scr.AddObject "Strings", str, True
'            Scr.AddObject "SymbolTable", SYM, True
'            Scr.AddObject "System", SYS, True
'            Scr.AddObject "Temp", Tmp, True
'            Scr.AddObject "Testing", TST, True
'            Scr.AddObject "Text", TXT, True
'            Scr.AddObject "Trace", TRC, True
'            Scr.AddObject "Unicode", UNI, True
'            Scr.AddObject "Users", USR, True
'            Scr.AddObject "Variable", VAR, True
'            Scr.AddObject "Wallpaper", WAL, True
'
'            Scr.AddObject "Dialogs", DLG, True
'            Scr.AddObject "FSO", FSO, True
'            Scr.AddObject "Printer", PRN, True
'            Scr.AddObject "Font", FNT, True
'            Scr.AddObject "RegExp", RGX, True
'
'            Scr.AddObject "Application", App, True
'            Scr.AddObject "Screen", Screen, True
'            Scr.AddObject "Panel", Form1, True
'        End If
        Dim sExtensions As String
        'sExtensions = App.Path & "\" & App.ExeName & ".ext"
        sExtensions = "Option Explicit" & vbNewLine
        
'        & vbNewLine & _
'        "Sub Include(scriptName)" & vbNewLine & _
'        "    Dim sScript" & vbNewLine & _
'        "    Dim oStream" & vbNewLine & _
'        "    With CreateObject(""Scripting.FileSystemobject"")" & vbNewLine & _
'        "        Set oStream = .OpenTextFile(scriptName)" & vbNewLine & _
'        "    End With" & vbNewLine & _
'        "    sScript = oStream.ReadAll()" & vbNewLine & _
'        "    oStream.Close" & vbNewLine & _
'        "    ExecuteGlobal sScript" & vbNewLine & _
'        "End Sub" & vbNewLine
'
        Dim lExtensions As Long
        'lExtensions = UBound(Split(sExtensions, vbNewLine))
        lExtensions = 0
        
'        Dim fn As Integer
'        fn = FreeFile(1)
'        Open sFilename For Binary As fn
'        sData = String$(LOF(fn), " ")
'        Get fn, , sData
'        Close fn
        sData = oFiles.ReadFileE(sFilename)
        sCodeBlob = sExtensions & vbNewLine & sData
        Dim bFound As Boolean
        Do
            processIncludes bFound, sCodeBlob
            If bFound = False Then Exit Do
        Loop
        
        processUses sCodeBlob
        
        On Error Resume Next
        processNeeds sCodeBlob
        If Err.Number = vbObjectError + 999 Then
            MsgBox Err.Description
            End
        End If
    
        On Error Resume Next
        Form1.Die Number(sCodeBlob)
        Scr.AddCode sCodeBlob
        If Err.Number <> 0 Then
            MsgBox Scr.Error.source & vbNewLine & Scr.Error.Description & vbNewLine & "Line: " & Scr.Error.Line - lExtensions & vbNewLine & "Column: " & Scr.Error.Column & vbNewLine & Scr.Error.Text
            Form1.Die Number(sCodeBlob) & vbNewLine & Scr.Error.source & vbNewLine & Scr.Error.Description & vbNewLine & "Line: " & Scr.Error.Line - lExtensions & vbNewLine & "Column: " & Scr.Error.Column & vbNewLine & Scr.Error.Text, True
        End If
        
    End If
End Sub

Function Number(s As String) As String
    Dim a As Variant
    a = Split(s, vbNewLine)
    For i = 0 To UBound(a) - 1
        a(i) = CStr(i + 1) & ": " & a(i)
    Next
    Number = Join(a, vbNewLine)
End Function

Sub processIncludes(ByRef bFound As Boolean, ByRef sBlock As String)
    Dim b As Boolean
    Dim a As Variant
    Dim i As Long
    Dim e As String
    Dim F As String
    
    b = False
    a = Split(sBlock, vbNewLine)
    For i = 0 To UBound(a) - 1
        e = LTrim$(a(i))
        If LCase$(oStrings.FirstWord(e)) = "include" Then
            b = True
            F = LTrim$(Mid$(e, 8))
            If oFiles.Exists(F) Then
                a(i) = oFiles.ReadFileE(F)
            Else
                a(i) = "' File not found error: " & a(i)
            End If
        End If
    Next
    bFound = b
    sBlock = Join(a, vbNewLine)
End Sub

Sub processUses(ByRef sBlob As String)
    Dim a As Variant
    Dim i As Long
    Dim e As String
    Dim F As String
    Dim sUsed As String
    
    sUsed = vbNullChar
    a = Split(sBlob, vbNewLine)
    For i = 0 To UBound(a) - 1
        e = LTrim$(a(i))
        If LCase$(oStrings.FirstWord(e)) = "uses" Then
            a(i) = "'" & a(i)
            b = True
            F = LCase$(Trim$(Mid$(e, 5)))
            If Not oStrings.Contains(sUsed, vbNullChar & F & vbNullChar) Then
                Select Case F
                Case "application"
                    Scr.AddObject "Application", App, True
                
                Case "arrays"
                    Set ARR = New Arrays
                    Scr.AddObject "Arrays", ARR, True
                
                Case "bigstring"
                    Set BIG = New BigString
                    Scr.AddObject "BigString", BIG, True
                    
                Case "clipboard"
                    Set CLP = New Clipboards
                    Scr.AddObject "Clipboard", CLP, True
    
                Case "cmdline"
                    Set CMD = New CmdLine
                    Scr.AddObject "CmdLine", CMD, True
    
                Case "complex"
                    Set CPX = New Complexes
                    Scr.AddObject "Complex", CPX, True
                
                Case "computers"
                    Set CMP = New Computers
                    Scr.AddObject "Computer", CMP, True
                
                Case "config"
                    Set CFG = New Config
                    Scr.AddObject "Config", CFG, True
                
                Case "database"
                    Set DBF = New Database
                    Scr.AddObject "Database", DBF, True
                
                Case "datetime"
                    Set DTM = New DateTime
                    Scr.AddObject "DateTime", DTM, True
                
                Case "decimal"
                    Set Dec = New Decimals
                    Scr.AddObject "Decimal", Dec, True
                
                Case "dialogs"
                    Set DLG = New CommonDialog
                    Scr.AddObject "Dialogs", DLG, True
                
                Case "domains"
                    Set DOM = New Domains
                    Scr.AddObject "Domain", DOM, True
                
                Case "drives"
                    Set DRV = New Drives
                    Scr.AddObject "Drive", DRV, True
                
                Case "email"
                    Set EML = New Emails
                    Scr.AddObject "Email", EML, True
                
                Case "environment"
                    Set ENV = New Environments
                    Scr.AddObject "Environment", ENV, True
                
                Case "excel"
                    Set XCL = New Excel
                    Scr.AddObject "Excel", XCL, True
                
                Case "files"
                    Set Fil = New Files
                    Scr.AddObject "Files", Fil, True
                
                Case "fso"
                    Scr.AddObject "FSO", FSO, True
                
                Case "groups"
                    Set GRP = New Groups
                    Scr.AddObject "Group", GRP, True
                
                Case "help"
                    Set HLP = New Help
                    Scr.AddObject "Help", HLP, True
                
                Case "ini1"
                    Set INI = New INI1
                    Scr.AddObject "INI1", INI, True
                
                Case "ini2"
                    Set IN2 = New INI2
                    Scr.AddObject "INI2", IN2, True
                
                Case "ini3"
                    Set IN3 = New INI3
                    Scr.AddObject "INI3", IN3, True
                
                Case "logic"
                    Set LGC = New Logic
                    Scr.AddObject "Logic", LGC, True
                
                Case "mathparser"
                    Set MTH = New MathParser
                    Scr.AddObject "MathParser", MTH, True
                
                Case "misc"
                    Set MSC = New Misc
                    Scr.AddObject "Misc", MSC, True
                
                Case "mouse"
                    Set MSE = New Mouse
                    Scr.AddObject "Mouse", MSE, True
                
                Case "number"
                    Set NUM = New Number
                    Scr.AddObject "Number", NUM, True
                
                Case "panel"
                    Scr.AddObject "Panel", Form1, True
                
                Case "pict"
                    Set PCT = New Pict
                    PCT.Init PRN
                    Scr.AddObject "Pict", PCT, True
                
                Case "picture"
                    Set pic = New Picture
                    Scr.AddObject "Picture", pic, True
                
                Case "printer"
                    Set PRN = Printer
                    Scr.AddObject "Printer", PRN, True
                
    '            Case "protium"
    '                Set PRO = New PROTIUM
    '                Scr.AddObject "Protium", PRO, True
    '
                Case "regexp"
                    Set RGX = New RegExp
                    Scr.AddObject "RegExp", RGX, True
                
                Case "registry"
                    Set reg = New Registry
                    Scr.AddObject "Registry", reg, True
                
                Case "schedule"
                    Set SCH = New Schedule
                    Scr.AddObject "Schedule", SCH, True
                
                Case "screen"
                    Scr.AddObject "Screen", Screen, True
                
                Case "sort"
                    Set SRT = New Sort
                    Scr.AddObject "Sort", SRT, True
                
                Case "font"
                    Set FNT = New StdFont
                    Scr.AddObject "Font", FNT, True
                
                Case "stopwatch"
                    Set SWT = New StopWatch
                    Scr.AddObject "StopWatch", SWT, True
                
                Case "strings"
                    Set str = New Strings
                    Scr.AddObject "Strings", str, True
                
                Case "symboltable"
                    Set SYM = New SymbolTable
                    Scr.AddObject "SymbolTable", SYM, True
                
                Case "system"
                    Set SYS = New System
                    Scr.AddObject "System", SYS, True
                
                Case "temp"
                    Set Tmp = New temp
                    Scr.AddObject "Temp", Tmp, True
                
                Case "testing"
                    Set TST = New Testing
                    Scr.AddObject "Testing", TST, True
                
                Case "text"
                    Set TXT = New Text
                    TXT.Init PRN
                    Scr.AddObject "Text", TXT, True
                
                Case "trace"
                    Set TRC = New traces
                    Scr.AddObject "Trace", TRC, True
                
                Case "unicode"
                    Set UNI = New Unicode
                    Scr.AddObject "Unicode", UNI, True
                
                Case "users"
                    Set USR = New Users
                    Scr.AddObject "Users", USR, True
                
                Case "variable"
                    Set VAR = New Variable
                    Scr.AddObject "Variable", VAR, True
                
                Case "wallpaper"
                    Set WAL = New Wallpaper
                    Scr.AddObject "Wallpaper", WAL, True
                            
                End Select
                sUsed = sUsed & F & vbNullChar
            End If
        End If
    Next
    sBlob = Join(a, vbNewLine)
End Sub

Sub processNeeds(ByRef sBlob As String)
    Dim a As Variant
    Dim i As Long
    Dim e As String
    Dim F As String
    Dim sUsed As String
    Dim reg As New Registry
    
    sUsed = vbNullChar
    a = Split(sBlob, vbNewLine)
    For i = 0 To UBound(a) - 1
        e = LTrim$(a(i))
        If LCase$(oStrings.FirstWord(e)) = "needs" Then
            a(i) = "'" & a(i)
            b = True
            F = LCase$(Trim$(Mid$(e, 6)))
            If Not oStrings.Contains(sUsed, vbNullChar & F & vbNullChar) Then
                reg.Needs F
                sUsed = sUsed & F & vbNullChar
            End If
        End If
    Next
    sBlob = Join(a, vbNewLine)
End Sub


'Printer Object Constants
'
'
'Printer Color Mode
'Constant Value Description
'vbPRCMMonochrome 1 Monochrome output
'vbPRCMColor 2 Color output
'
'
'Duplex Printing
'Constant Value Description
'vbPRDPSimplex 1 Single-sided printing
'vbPRDPHorizontal 2 Double-sided horizontal printing
'vbPRDPVertical 3 Double-sided vertical printing
'
'
'Printer Orientation
'Constant Value Description
'vbPRORPortrait 1 Documents print with the top at the narrow side of the paper
'vbPRORLandscape 2 Documents print with the top at the wide side of the paper
'
'
'Print Quality
'Constant Value Description
'vbPRPQDraft -1 Draft print quality
'vbPRPQLow -2 Low print quality
'vbPRPQMedium -3 Medium print quality
'vbPRPQHigh -4 High print quality
'
'
'PaperBin Property
'Constant Value Description
'vbPRBNUpper 1 Use paper from the upper bin
'vbPRBNLower 2 Use paper from the lower bin
'vbPRBNMiddle 3 Use paper from the middle bin
'vbPRBNManual 4 Wait for manual insertion of each sheet of paper
'vbPRBNEnvelope 5 Use envelopes from the envelope feeder
'vbPRBNEnvManual 6 Use envelopes from the envelope feeder, but wait for manual insertion
'vbPRBNAuto 7 (Default) Use paper from the current default bin
'vbPRBNTractor 8 Use paper fed from the tractor feeder
'vbPRBNSmallFmt 9 Use paper from the small paper feeder
'vbPRBNLargeFmt 10 Use paper from the large paper bin
'vbPRBNLargeCapacity 11 Use paper from the large capacity feeder
'vbPRBNCassette 14 Use paper from the attached cassette cartridge
'
'
'PaperSize Property
'Constant Value Description
'vbPRPSLetter 1 Letter, 8 1/2 x 11 in
'vbPRPSLetterSmall 2 +A611Letter Small, 8 1/2 x 11 in
'vbPRPSTabloid 3 Tabloid, 11 x 17 in
'vbPRPSLedger 4 Ledger, 17 x 11 in
'vbPRPSLegal 5 Legal, 8 1/2 x 14 in
'vbPRPSStatement 6 Statement, 5 1/2 x 8 1/2 in
'vbPRPSExecutive 7 Executive, 7 1/2 x 10 1/2 in
'vbPRPSA3 8 A3, 297 x 420 mm
'vbPRPSA4 9 A4, 210 x 297 mm
'vbPRPSA4Small 10 A4 Small, 210 x 297 mm
'vbPRPSA5 11 A5, 148 x 210 mm
'vbPRPSB4 12 B4, 250 x 354 mm
'vbPRPSB5 13 B5, 182 x 257 mm
'vbPRPSFolio 14 Folio, 8 1/2 x 13 in
'vbPRPSQuarto 15 Quarto, 215 x 275 mm
'vbPRPS1&H14 16 10 x 14 in
'vbPRPS11x17 17 11 x 17 in
'vbPRPSNote 18 Note, 8 1/2 x 11 in
'vbPRPSEnv9 19 Envelope #9, 3 7/8 x 8 7/8 in
'vbPRPSEnv10 20 Envelope #10, 4 1/8 x 9 1/2 in
'vbPRPSEnv11 21 Envelope #11, 4 1/2 x 10 3/8 in
'vbPRPSEnv12 22 Envelope #12, 4 1/2 x 11 in
'vbPRPSEnv14 23 Envelope #14, 5 x 11 1/2 in
'vbPRPSCSheet 24 C size sheet
'vbPRPSDSheet 25 D size sheet
'vbPRPSESheet 26 E size sheet
'vbPRPSEnvDL 27 Envelope DL, 110 x 220 mm
'vbPRPSEnvC3 29 Envelope C3, 324 x 458 mm
'vbPRPSEnvC4 30 Envelope C4, 229 x 324 mm
'vbPRPSEnvC5 28 Envelope C5, 162 x 229 mm
'vbPRPSEnvC6 31 Envelope C6, 114 x 162 mm
'vbPRPSEnvC65 32 Envelope C65, 114 x 229 mm
'vbPRPSEnvB4 33 Envelope B4, 250 x 353 mm
'vbPRPSEnvB5 34 Envelope B5, 176 x 250 mm
'vbPRPSEnvB6 35 Envelope B6, 176 x 125 mm
'vbPRPSEnvItaly 36 Envelope, 110 x 230 mm
'vbPRPSEnvMonarch 37 Envelope Monarch, 3 7/8 x 7 1/2 in
'vbPRPSEnvPersonal 38 Envelope, 3 5/8 x 6 1/2 in
'vbPRPSFanfoldUS 39 U.S. Standard Fanfold, 14 7/8 x 11 in
'vbPRPSFanfoldStdGerman 40 German Standard Fanfold, 8 1/2 x 12 in
'vbPRPSFanfoldLglGerman 41 German Legal Fanfold, 8 1/2 x 13 in
'vbPRPSUser 256 User-defined
'




