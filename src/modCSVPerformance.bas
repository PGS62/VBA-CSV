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
    Wrap_ws_garcia = "#Wrap_ws_garcia: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Wrap_sdkn104
' Purpose    : Wrapper to https://github.com/sdkn104/VBA-CSV
'              Wraps version 1.9 - module CSVUtils imported as sdkn104_CSVUtils
' -----------------------------------------------------------------------------------------------------------------------
Public Function Wrap_sdkn104(FileName As String, Unicode As Boolean) As Variant
    Dim Contents As String
    Dim FSO As New FileSystemObject
    Dim T As Scripting.TextStream

    On Error GoTo ErrHandler

    Set T = FSO.GetFile(FileName).OpenAsTextStream(ForReading, IIf(Unicode, TristateTrue, TristateFalse))
    Contents = T.ReadAll
    T.Close
    Wrap_sdkn104 = ParseCSVToArray(Contents)
    
    Exit Function
ErrHandler:
    Wrap_sdkn104 = "#Wrap_sdkn104: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RunSpeedTests
' Purpose    : Attached to the "Run Speed Tests..." button. Note the significance of the "PasteResultsHere" ranges
' -----------------------------------------------------------------------------------------------------------------------
Private Sub RunSpeedTests()

    Const NumColsInTFPRet As Long = 10
    Const Timeout As Long = 5
    Const Title As String = "VBA-CSV Speed Tests"
    Dim c As Range
    Dim JuliaResultsFile As String
    Dim N As Name
    Dim Prompt As String
    Dim TestResults As Variant
    Dim ws As Worksheet

    On Error GoTo ErrHandler
    
    Set ws = ActiveSheet
    
    Prompt = "Run Speed tests?" + vbLf + vbLf + _
        "Note this will generate approx 227MB of files in folder" + vbLf + _
        Environ$("Temp") & "\VBA-CSV\Performance"

    If MsgBox(Prompt, vbOKCancel + vbQuestion, Title) <> vbOK Then Exit Sub
    
    ws.Protect , , False
    
    ws.Range("TimeStamp").value = "This data generated " & Format$(Now, "dd-mmmm-yyyy hh:mm:ss")
    
    'Julia results file created by Julia function benchmark. See julia/benchmarkCSV.jl
    
    JuliaResultsFile = Left$(ThisWorkbook.path, InStrRev(ThisWorkbook.path, "\")) + "julia\juliaparsetimes.csv"
    If Not FileExists(JuliaResultsFile) Then
        Throw "Cannot find file '" + JuliaResultsFile + "'"
    End If
    
    For Each N In ws.Names
        If InStr(N.Name, "PasteResultsHere") > 1 Then
            Application.Goto N.RefersToRange

            For Each c In N.RefersToRange.Cells
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
    Next N

    AddCharts False

    ws.Protect , , True

    Exit Sub

    Exit Sub
ErrHandler:
    MsgBox "#RunSpeedTests (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical
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
    Dim j As Long
    Dim JuliaResults As Variant
    Dim k As Double
    Dim NumCalls As Variant
    Dim OS As String
    Dim Ret As Variant
    Dim timeTaken As Variant
    Dim Tstart As Double
    Dim NumFns As Long
    Dim JuliaResultsFile As String

    On Error GoTo ErrHandler
    
    JuliaResultsFile = Left$(ThisWorkbook.path, InStrRev(ThisWorkbook.path, "\")) + "julia\juliaparsetimes.csv"
    
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

    Data = Fill(EachFieldContains, NumRows, NumCols)
    FileName = NameThatFile(Folder, OS, NumRows, NumCols, Replace(ExtraInfo, " ", "-"), Unicode, False)

    ThrowIfError CSVWrite(Data, FileName, False)

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
        ElseIf InStr(ParserNames(1, j), "ArrayFromCSV") > 0 Then
            Tstart = ElapsedTime()
            Do
                DataReread = ThrowIfError(ArrayFromCSVfile(FileName))
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
            Throw "Unrecognised element of ParserNames: " + CStr(ParserNames(1, j))
        End If
        On Error Resume Next
        Ret(1, j) = timeTaken / NumCalls
        On Error GoTo ErrHandler
        Ret(1, NumFns + j) = NumCalls
    Next j
    
    Ret(1, 2 * NumFns + 1) = FileName
    Ret(1, 2 * NumFns + 2) = FileSize(FileName)

    TimeParsers = Ret

    Exit Function
ErrHandler:
    TimeParsers = "#TimeParsers (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function


Sub AddChartsNoExport()
    On Error GoTo ErrHandler
    AddCharts False

    Exit Sub
ErrHandler:
    MsgBox "#AddChartsNoExport (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Sub AddCharts(Optional Export As Boolean = True)
          
          Dim c As ChartObject
          Dim N As Name
          Dim prot As Boolean
          Dim ws As Worksheet
          Dim xData As Range
          Dim yData As Range
          Dim NumSeries As Long

1         On Error GoTo ErrHandler
          
2         Set ws = ActiveSheet
          
3         prot = ws.ProtectContents

          NumSeries = ws.Range("ParserNames").Columns.count

4         ws.Unprotect
          
5         For Each c In ws.ChartObjects
6             c.Delete
7         Next

8         For Each N In ws.Names
9             If InStr(N.Name, "PasteResultsHere") > 1 Then
10                Set yData = N.RefersToRange
11                With yData
12                    Set yData = .Offset(-1).Resize(.Rows.count + 1, NumSeries)
16                End With
17                Set xData = yData.Offset(, -1).Resize(, 1)
18                With xData
19                    If .Cells(2, 1).value = .Cells(.Rows.count, 1).value Then
20                        Set xData = xData.Offset(, -1)
21                    End If
22                End With
23                With xData
24                    If .Cells(2, 1).value = .Cells(.Rows.count, 1).value Then
25                        Set xData = xData.Offset(, -2)
26                    End If
27                End With
28                AddChart xData, yData, Export
29            End If
30        Next N

31        Application.Goto ws.Cells(1, 1)
32        ws.Protect , , prot

33        Exit Sub
ErrHandler:
34        MsgBox "#AddCharts (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Sub AddChartAtSelection()
    On Error GoTo ErrHandler
    
    AddChart

    Exit Sub
ErrHandler:
    MsgBox "#AddChartAtSelection (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : AddChart
' Purpose    : Adds a chart to the sheet Timings. First select the data to plot then run this macro by clicking the
'              "Add Chart" button
' -----------------------------------------------------------------------------------------------------------------------
Sub AddChart(Optional xData As Range, Optional yData As Range, Optional Export As Boolean)

          Dim ChartsInCol As String
          
          Const Err_BadSelection As String = "That selection does not look correct." + vbLf + vbLf + _
              "Select two areas to define the data to plot. The first area should contain " + _
              "the independent data and have a single column with top cell giving the x axis " + _
              "label. The second area should contain the dependent data with one column per data " + _
              "series and top row giving the series names. Both areas should have the same number of rows"
              
          Dim ch As Chart
          Dim shp As Shape
          Dim SourceData As Range
          Dim Title As String
          Dim TitleCell As Range
          Dim TitlesInCol As String
          Dim TopLeftCell As Range
          Dim wsh As Worksheet

1         On Error GoTo ErrHandler

2         If xData Is Nothing Then
3             Set SourceData = Selection

4             If SourceData.Areas.count <> 2 Then
5                 Throw Err_BadSelection
6             ElseIf SourceData.Areas(1).Rows.count <> SourceData.Areas(2).Rows.count Then
7                 Throw Err_BadSelection
8             End If
9             Set xData = SourceData.Areas(1)
10            Set yData = SourceData.Areas(2)
11            Set wsh = xData.Parent
12        Else
              'Actually selecting the ranges seems to be necessary to get the legends to appear in the generated charts...
13            Set wsh = xData.Parent
14            wsh.Activate
15            wsh.Range(xData.Address & "," & yData.Address).Select
16        End If
          
17        With xData.Parent.Range("ParserNames")

18            ChartsInCol = .Offset(0, .Columns.count * 2 + 3).Resize(1, 1).Address
19            ChartsInCol = Mid(ChartsInCol, 2, InStr(2, ChartsInCol, "$") - 2)

20            TitlesInCol = .Offset(0, .Columns.count * 2).Resize(1, 1).Address
21            TitlesInCol = Mid(TitlesInCol, 2, InStr(2, TitlesInCol, "$") - 2)
22        End With
          
23        Set shp = wsh.Shapes.AddChart2(240, xlXYScatterLines)
24        Set ch = shp.Chart
25        ch.SetSourceData Source:=Application.Union(xData, yData)
26        Set TopLeftCell = Application.Intersect(xData.Cells(1, 1).EntireRow, wsh.Range(ChartsInCol & ":" & ChartsInCol))
27        Set TitleCell = Application.Intersect(xData.Cells(0, 1).EntireRow, wsh.Range(TitlesInCol & ":" & TitlesInCol))

28        Title = "='" & wsh.Name & "'!R" & TitleCell.Row & "C" & TitleCell.Column

29        ch.Axes(xlCategory).ScaleType = xlLogarithmic
30        ch.Axes(xlValue).ScaleType = xlLogarithmic
31        ch.Axes(xlValue, xlPrimary).HasTitle = True
32        ch.Axes(xlValue, xlPrimary).AxisTitle.text = "Seconds to read. Log Scale"
33        ch.Axes(xlCategory).HasTitle = True
34        ch.Axes(xlCategory).AxisTitle.text = xData.Cells(1, 1).value + ". Log Scale"
35        ch.ChartTitle.Caption = Title
36        With xData
37            ch.Axes(xlCategory).MinimumScale = .Cells(2, 1).value
38            ch.Axes(xlCategory).MaximumScale = .Cells(.Rows.count, 1).value
39        End With

40        shp.Top = TopLeftCell.Top
41        shp.Left = TopLeftCell.Left
42        shp.Height = 394
43        shp.Width = 561
44        shp.Placement = xlMove
          
45        If Export Then
              Dim FileName As String
              Dim Folder As String
46            FileName = Replace(TitleCell.Offset(-1).value, " ", "_")
47            Folder = Left$(ThisWorkbook.path, InStrRev(ThisWorkbook.path, "\")) + "images\"
48            ch.Export Folder + FileName
49        End If

50        Exit Sub
ErrHandler:
51        Throw "#AddChart (line " & CStr(Erl) + "): " & Err.Description & "!"

End Sub
