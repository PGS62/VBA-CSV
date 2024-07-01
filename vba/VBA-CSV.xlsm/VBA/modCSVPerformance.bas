Attribute VB_Name = "modCSVPerformance"
' VBA-CSV

' Copyright (C) 2021 - Philip Swannell (https://github.com/PGS62/VBA-CSV )
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/VBA-CSV#readme

Option Explicit

'Code of the "Performance" worksheet

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Wrap_ws_garcia
' Purpose    : Wrap to https://github.com/ws-garcia/VBA-CSV-interface
'              Wraps version 3.1.5, (the four class module CSVInterface, ECPArrayList, ECPTextStream, parserConfig)
' -----------------------------------------------------------------------------------------------------------------------
Public Function Wrap_ws_garcia(FileName As String, Delimiter As String, ByVal EOL As String, Optional SkipEmptyLines As Boolean, Optional supportMixedLineEndings As Boolean) As Variant

    Dim CSVint As CSVinterface
    Dim oArray()

    On Error GoTo ErrHandler

    Set CSVint = New CSVinterface
    With CSVint.parseConfig
        .path = FileName            ' Full path to the file, including its extension.
        .fieldsDelimiter = Delimiter         ' Columns delimiter
        .recordsDelimiter = EOL     ' Rows delimiter
        .skipCommentLines = False  'I think code runs faster if not testing for skipping comment lines or empty lines
        .SkipEmptyLines = SkipEmptyLines
        If supportMixedLineEndings Then
        .turnStreamRecDelimiterToLF = True
        End If
    End With
    With CSVint
        .ImportFromCSV .parseConfig    ' Import the CSV to internal object
        .DumpToArray oArray
    End With

    Wrap_ws_garcia = oArray

    Exit Function
ErrHandler:
    Wrap_ws_garcia = ReThrow("Wrap_ws_garcia", Err, True)
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Wrap_sdkn104
' Purpose    : Wrapper to https://github.com/sdkn104/VBA-CSV
'              Wraps version 1.9 - module CSVUtils imported as sdkn104_CSVUtils
' -----------------------------------------------------------------------------------------------------------------------
Public Function Wrap_sdkn104(FileName As String, Unicode As Boolean) As Variant
    Dim Contents As String
    Dim FSO As New FileSystemObject
    Dim t As Scripting.TextStream

    On Error GoTo ErrHandler

    Set t = FSO.GetFile(FileName).OpenAsTextStream(ForReading, IIf(Unicode, TristateTrue, TristateFalse))
    Contents = t.ReadAll
    t.Close
    Wrap_sdkn104 = ParseCSVToArray(Contents)
    
    Exit Function
ErrHandler:
    Wrap_sdkn104 = ReThrow("Wrap_sdkn104", Err, True)
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RunSpeedTests
' Purpose    : Attached to the "Run Speed Tests..." button. Note the significance of the "PasteResultsHere" ranges
' -----------------------------------------------------------------------------------------------------------------------
Private Sub RunSpeedTests()

    Dim NumColsInTFPRet As Long
    Const Timeout As Long = 5
    Const Title As String = "VBA-CSV Speed Tests"
    Dim c As Range
    Dim JuliaResultsFile As String
    Dim n As Name
    Dim Prompt As String
    Dim TestResults As Variant
    Dim ws As Worksheet

    On Error GoTo ErrHandler
    
    Set ws = ActiveSheet
    
    Prompt = "Run Speed tests?" & vbLf & vbLf & _
        "Note this will generate approx 227MB of files in folder" & vbLf & _
        Environ$("Temp") & "\VBA-CSV\Performance"

    If MsgBox(Prompt, vbOKCancel + vbQuestion, Title) <> vbOK Then Exit Sub
    
    ws.Protect , , False
    
    ws.Range("TimeStamp").value = "This data generated " & Format$(Now, "dd-mmmm-yyyy hh:mm:ss")
    
    'Julia results file created by Julia function benchmark. See julia/benchmarkCSV.jl
    
    NumColsInTFPRet = ws.Range("ParserNames").Columns.count * 2 + 2
    JuliaResultsFile = Left$(ThisWorkbook.path, InStrRev(ThisWorkbook.path, "\")) & "julia\juliaparsetimes.csv"
    If Not FileExists(JuliaResultsFile) Then
        Throw "Cannot find file '" & JuliaResultsFile & "'"
    End If
    
    For Each n In ws.Names
        If InStr(n.Name, "PasteResultsHere") > 1 Then
            If NameRefersToRange(n) Then
                Application.GoTo n.RefersToRange
                For Each c In n.RefersToRange.Cells
                    c.Resize(1, NumColsInTFPRet).ClearContents
                    TestResults = ThrowIfError(TimeParsers(ws.Range("ParserNames").value, c.Offset(0, -3).value, c.Offset(0, -2).value, _
                        c.Offset(0, -1).value, Timeout))
                    c.Resize(1, NumColsInTFPRet).value = TestResults
                    ws.Calculate
                    DoEvents
                    Application.ScreenUpdating = True
                Next
                ThisWorkbook.Save
            End If
        End If
    Next n

    AddCharts False

    ws.Protect , , True

    Exit Sub

    Exit Sub
ErrHandler:
    MsgBox ReThrow("RunSpeedTests", Err, True), vbCritical
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : TimeParsers
' Purpose    : Core of the method RunSpeedTests. Note the functions being timed are called many times in a loop that
'              exits after TimeOut seconds have elapsed. Leads to much more reliable timings than timing a single call.
' -----------------------------------------------------------------------------------------------------------------------
Public Function TimeParsers(ByVal ParserNames As Variant, EachFieldContains As Variant, _
          NumRows As Long, NumCols As Long, Timeout As Double) As Variant

    Const Unicode As Boolean = False
    Dim Data As Variant
    Dim DataReread As Variant
    Dim ExtraInfo As String
    Dim FileName As String
    Dim Folder As String
    Dim i As Long
    Dim j As Long
    Dim JuliaResults As Variant
    Dim JuliaResultsFile As String
    Dim k As Double
    Dim NumCalls As Variant
    Dim NumFns As Long
    Dim OS As String
    Dim Ret As Variant
    Dim t1 As Double
    Dim t2 As Double
    Dim timeTaken As Variant
    Dim Tstart As Double
    Static Overhead As Double

    On Error GoTo ErrHandler
    
    'The timing loop has an inside-the-loop overhead of calling ElapsedTime (approx 6 microseconds),
    'about 5% of the execution time for reading a one-character file from local disc.
    If Overhead = 0 Then
        t1 = ElapsedTime()
        For i = 1 To 100000
            t2 = ElapsedTime()
        Next i
        Overhead = (ElapsedTime() - t1) / 100000
    End If
    
    JuliaResultsFile = Left$(ThisWorkbook.path, InStrRev(ThisWorkbook.path, "\")) & "julia\juliaparsetimes.csv"
    
    JuliaResults = ThrowIfError(CSVRead(JuliaResultsFile, True))
    
    OS = vbNullString
    
    If VarType(EachFieldContains) = vbDouble Then
        ExtraInfo = "Doubles"
    ElseIf VarType(EachFieldContains) = vbString Then
        If Left$(EachFieldContains, 1) = """" And Right$(EachFieldContains, 1) = """" Then
            If InStr(EachFieldContains, vbLf) > 0 Then
                ExtraInfo = "Quoted_Strings_with LF_length_" & Len(EachFieldContains)
            Else
                ExtraInfo = "Quoted_Strings_length_" & Len(EachFieldContains)
            End If
        Else
            If InStr(EachFieldContains, vbLf) > 0 Then
                ExtraInfo = "Strings_with_LF_length_" & Len(EachFieldContains)
            Else
                ExtraInfo = "Strings_length_" & Len(EachFieldContains)
            End If
        End If
    Else
        ExtraInfo = "Unknown"
    End If

    Folder = Environ$("Temp") & "\VBA-CSV\Performance"

    ThrowIfError CreatePath(Folder)

    FileName = NameThatFile(Folder, OS, NumRows, NumCols, Replace(ExtraInfo, " ", "-"), Unicode, False)
    If Not FileExists(FileName) Then 'Assumes filename acts like a hash of the file!
        Data = Fill(EachFieldContains, NumRows, NumCols)
        ThrowIfError CSVWrite(Data, FileName, False)
    End If

    Force2DArrayR ParserNames
    NumFns = NCols(ParserNames)
    
    Ret = Fill("", 1, NumFns * 2 + 2)

    For j = 1 To NumFns
        k = 0
        
        If InStr(ParserNames(1, j), "CSVRead") > 0 Then
            Tstart = ElapsedTime()
            Do
                DataReread = ThrowIfError(CSVRead(FileName, False, ",", , , , False, , , , , , , , , , "ANSI"))
                k = k + 1
                If ElapsedTime() - Tstart > Timeout Then Exit Do
            Loop
            timeTaken = ElapsedTime - Tstart
            NumCalls = k
        ElseIf InStr(ParserNames(1, j), "sdkn104") > 0 Then
            Tstart = ElapsedTime()
            Do
                DataReread = ThrowIfError(Wrap_sdkn104(FileName, Unicode))
                k = k + 1
                If ElapsedTime() - Tstart > Timeout Then Exit Do
            Loop
            timeTaken = ElapsedTime - Tstart
            NumCalls = k
        ElseIf InStr(ParserNames(1, j), "garcia") > 0 Then
            Tstart = ElapsedTime()
            Do
                DataReread = ThrowIfError(Wrap_ws_garcia(FileName, ",", vbCrLf))
                k = k + 1
                If ElapsedTime() - Tstart > Timeout Then Exit Do
            Loop
            timeTaken = ElapsedTime - Tstart
            NumCalls = k
        ElseIf InStr(ParserNames(1, j), "CSV.jl") > 0 Then
            DataReread = Empty
            NumCalls = "Not found"
            timeTaken = "Not found"
            On Error Resume Next
            NumCalls = Application.WorksheetFunction.VLookup(FileName, JuliaResults, 4, False)
            timeTaken = Application.WorksheetFunction.VLookup(FileName, JuliaResults, 2, False)
            timeTaken = timeTaken * NumCalls
            On Error GoTo ErrHandler
        Else
            Throw "Unrecognised element of ParserNames: " & CStr(ParserNames(1, j))
        End If
        On Error Resume Next
        Ret(1, j) = (timeTaken / NumCalls) - Overhead
        On Error GoTo ErrHandler
        Ret(1, NumFns + j) = NumCalls
    Next j
    
    Ret(1, 2 * NumFns + 1) = FileName
    Ret(1, 2 * NumFns + 2) = FileSize(FileName)

    TimeParsers = Ret

    Exit Function
ErrHandler:
    TimeParsers = ReThrow("TimeParsers", Err, True)
End Function

Sub AddChartsNoExport()
    On Error GoTo ErrHandler
    AddCharts False

    Exit Sub
ErrHandler:
    MsgBox ReThrow("AddChartsNoExport", Err, True), vbCritical
End Sub

Sub AddCharts(Optional Export As Boolean = True)
    
    Dim c As ChartObject
    Dim n As Name
    Dim NumSeries As Long
    Dim prot As Boolean
    Dim ws As Worksheet
    Dim xData As Range
    Dim yData As Range

    On Error GoTo ErrHandler
    
    Set ws = ActiveSheet
    
    prot = ws.ProtectContents

    NumSeries = ws.Range("ParserNames").Columns.count

    ws.Unprotect
    
    For Each c In ws.ChartObjects
        c.Delete
    Next

    For Each n In ws.Names
        If InStr(n.Name, "PasteResultsHere") > 1 Then
            If NameRefersToRange(n) Then
                Set yData = n.RefersToRange
                With yData
                    Set yData = .Offset(-1).Resize(.Rows.count + 1, NumSeries)
                End With
                Set xData = yData.Offset(, -1).Resize(, 1)
                With xData
                    If .Cells(2, 1).value = .Cells(.Rows.count, 1).value Then
                        Set xData = xData.Offset(, -1)
                    End If
                End With
                With xData
                    If .Cells(2, 1).value = .Cells(.Rows.count, 1).value Then
                        Set xData = xData.Offset(, -2)
                    End If
                End With
                AddChart xData, yData, Export
            End If
        End If
    Next n

    Application.GoTo ws.Cells(1, 1)
    ws.Protect , , prot

    Exit Sub
ErrHandler:
    MsgBox ReThrow("AddCharts", Err, True), vbCritical
End Sub

Private Function NameRefersToRange(n As Name) As Boolean
    Dim R As Range

    On Error GoTo ErrHandler
    Set R = n.RefersToRange
    NameRefersToRange = True
    
    Exit Function
ErrHandler:
    NameRefersToRange = False
End Function

Sub AddChartAtSelection()
    On Error GoTo ErrHandler
    
    AddChart

    Exit Sub
ErrHandler:
    MsgBox ReThrow("AddChartAtSelection", Err, True), vbCritical
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : AddChart
' Purpose    : Adds a chart to the sheet Timings. First select the data to plot then run this macro by clicking the
'              "Add Chart" button
' -----------------------------------------------------------------------------------------------------------------------
Sub AddChart(Optional xData As Range, Optional yData As Range, Optional Export As Boolean)

    Dim ChartsInCol As String
    
    Const Err_BadSelection As String = "That selection does not look correct." & vbLf & vbLf & _
        "Select two areas to define the data to plot. The first area should contain " & _
        "the independent data and have a single column with top cell giving the x axis " & _
        "label. The second area should contain the dependent data with one column per data " & _
        "series and top row giving the series names. Both areas should have the same number of rows"
        
    Dim ch As Chart
    Dim shp As Shape
    Dim SourceData As Range
    Dim Title As String
    Dim TitleCell As Range
    Dim TitlesInCol As String
    Dim TopLeftCell As Range
    Dim wsh As Worksheet

    On Error GoTo ErrHandler

    If xData Is Nothing Then
        Set SourceData = Selection

        If SourceData.Areas.count <> 2 Then
            Throw Err_BadSelection
        ElseIf SourceData.Areas(1).Rows.count <> SourceData.Areas(2).Rows.count Then
            Throw Err_BadSelection
        End If
        Set xData = SourceData.Areas(1)
        Set yData = SourceData.Areas(2)
        Set wsh = xData.Parent
    Else
        'Actually selecting the ranges seems to be necessary to get the legends to appear in the generated charts...
        Set wsh = xData.Parent
        wsh.Activate
        wsh.Range(xData.Address & "," & yData.Address).Select
    End If
    
    With xData.Parent.Range("ParserNames")

        ChartsInCol = .Offset(0, .Columns.count * 2 + 3).Resize(1, 1).Address
        ChartsInCol = Mid(ChartsInCol, 2, InStr(2, ChartsInCol, "$") - 2)

        TitlesInCol = .Offset(0, .Columns.count * 2).Resize(1, 1).Address
        TitlesInCol = Mid(TitlesInCol, 2, InStr(2, TitlesInCol, "$") - 2)
    End With
    
    Set shp = wsh.Shapes.AddChart2(240, xlXYScatterLines)
    Set ch = shp.Chart
    ch.SetSourceData Source:=Application.Union(xData, yData)
    Set TopLeftCell = Application.Intersect(xData.Cells(1, 1).EntireRow, wsh.Range(ChartsInCol & ":" & ChartsInCol))
    Set TitleCell = Application.Intersect(xData.Cells(0, 1).EntireRow, wsh.Range(TitlesInCol & ":" & TitlesInCol))

    Title = "='" & wsh.Name & "'!R" & TitleCell.Row & "C" & TitleCell.Column

    ch.Axes(xlCategory).ScaleType = xlLogarithmic
    ch.Axes(xlValue).ScaleType = xlLogarithmic
    ch.Axes(xlValue, xlPrimary).HasTitle = True
    ch.Axes(xlValue, xlPrimary).AxisTitle.text = "Seconds to read. Log Scale"
    ch.Axes(xlCategory).HasTitle = True
    ch.Axes(xlCategory).AxisTitle.text = xData.Cells(1, 1).value & ". Log Scale"
    ch.ChartTitle.Caption = Title
    With xData
        ch.Axes(xlCategory).MinimumScale = .Cells(2, 1).value
        ch.Axes(xlCategory).MaximumScale = .Cells(.Rows.count, 1).value
    End With

    shp.Top = TopLeftCell.Top
    shp.Left = TopLeftCell.Left
    shp.Height = 394
    shp.Width = 561
    shp.Placement = xlMove
    
    If Export Then
        Dim FileName As String
        Dim Folder As String
        FileName = Replace(TitleCell.Offset(-1).value, " ", "_")
        Folder = Left$(ThisWorkbook.path, InStrRev(ThisWorkbook.path, "\")) & "images\"
        ch.Export Folder + FileName
    End If

    Exit Sub
ErrHandler:
    ReThrow "AddChart", Err

End Sub

