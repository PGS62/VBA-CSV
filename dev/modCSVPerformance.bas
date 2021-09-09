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
Function Wrap_ws_garcia(FileName As String, Delimiter As String, ByVal EOL As String)

    Dim CSVint As CSVinterface
    Dim oArray() As Variant

    On Error GoTo ErrHandler

    Set CSVint = New CSVinterface
    With CSVint.parseConfig
        .path = FileName            ' Full path to the file, including its extension.
        .fieldsDelimiter = Delimiter         ' Columns delimiter
        .recordsDelimiter = EOL     ' Rows delimiter
        .skipCommentLines = False  'I think code runs faster if not testing for skipping comment lines or empty lines
        .skipEmptyLines = False
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
Public Function Wrap_sdkn104(FileName As String, Unicode As Boolean)
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

    Const NumColsInTFPRet = 10
    Const ReadFiles  As Boolean = True
    Const Timeout = 5
    Const Title = "VBA-CSV Speed Tests"
    Const WriteFiles As Boolean = True
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
        Environ("Temp") & "\VBA-CSV\Performance"

    If MsgBox(Prompt, vbOKCancel + vbQuestion, Title) <> vbOK Then Exit Sub
    
    ws.Protect , , False
    
    ws.Range("TimeStamp").value = "This data generated " & Format(Now, "dd-mmmm-yyyy hh:mm:ss")
    
    'Julia results file created by function benchmark. See julia/benchmarkCSV.jl, function benchmark
    
    JuliaResultsFile = Left$(ThisWorkbook.path, InStrRev(ThisWorkbook.path, "\")) + "\julia\juliaparsetimes.csv"
    If Not FileExists(JuliaResultsFile) Then
        Throw "Cannot find file '" + JuliaResultsFile + "'"
    End If
    
    For Each N In ws.Names
        If InStr(N.Name, "PasteResultsHere") > 1 Then
            Application.GoTo N.RefersToRange

            For Each c In N.RefersToRange.Cells
                c.Resize(1, NumColsInTFPRet).ClearContents
                TestResults = TimeFourParsers(WriteFiles, ReadFiles, c.Offset(0, -3).value, c.Offset(0, -2).value, _
                    c.Offset(0, -1).value, Timeout, False, JuliaResultsFile)
                c.Resize(1, NumColsInTFPRet).value = TestResults
                ws.Calculate
                DoEvents
                Application.ScreenUpdating = True
            Next
        End If
    Next N

    AddCharts

    ws.Protect , , True

    Exit Sub

    Exit Sub
ErrHandler:
    MsgBox "#RunSpeedTests (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : TimeFourParsers
' Purpose    : Core of the method RunSpeedTests. Note the functions being timed are called many times in a loop that exits
'              after TimeOut seconds have elapsed. Leads to much more reliable timings than timing a single call.
' -----------------------------------------------------------------------------------------------------------------------
Function TimeFourParsers(WriteFiles As Boolean, ReadFiles As Boolean, EachFieldContains As Variant, NumRows As Long, _
    NumCols As Long, Timeout As Double, WithHeaders As Boolean, JuliaResultsFile As String)

    Const Unicode = False
    Dim Data As Variant
    Dim DataReread1
    Dim DataReread2
    Dim DataReread3
    Dim DataReread4
    Dim DataRow As Long
    Dim ExtraInfo As String
    Dim FileName As String
    Dim FnName1 As String
    Dim FnName2 As String
    Dim FnName3 As String
    Dim FnName4 As String
    Dim Folder As String
    Dim j As Long
    Dim JuliaResults As Variant
    Dim k As Double
    Dim NumCalls1 As Long
    Dim NumCalls2 As Long
    Dim NumCalls3 As Long
    Dim NumCalls4 As Variant
    Dim OS As String
    Dim Ret As Variant
    Dim t1 As Double
    Dim t2 As Double
    Dim t3 As Double
    Dim t4 As Variant
    Dim Tend As Double
    Dim Tstart As Double

    On Error GoTo ErrHandler
    
    JuliaResults = CSVRead(JuliaResultsFile, True)
    
    OS = ""
    
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

    Folder = Environ("Temp") & "\VBA-CSV\Performance"

    ThrowIfError CreatePath(Folder)

    Data = Fill(EachFieldContains, NumRows, NumCols)
    FileName = NameThatFile(Folder, OS, NumRows, NumCols, Replace(ExtraInfo, " ", "-"), Unicode, False)
    If WriteFiles Then
        ThrowIfError CSVWrite(Data, FileName, False)
    End If
        
    If ReadFiles Then
        For j = 1 To 4
            k = 0
            Tstart = ElapsedTime()
            Do
                k = k + 1
                Select Case j
                    Case 1
                        FnName1 = "CSVRead" + vbLf + "v0.1"
                        DataReread1 = ThrowIfError(CSVRead(FileName, False, ",", , , , False, , , , , , , , , , "ANSI"))
                    Case 2
                        FnName2 = "sdkn104" + vbLf + "v1.9"
                        DataReread2 = ThrowIfError(Wrap_sdkn104(FileName, Unicode))
                    Case 3
                        FnName3 = "ws_garcia" + vbLf + "v3.1.5"

                        DataReread3 = ThrowIfError(Wrap_ws_garcia(FileName, ",", vbCrLf))
                    Case 4
                        FnName4 = "CSV.jl" + vbLf + "v0.8.5+"
                        DataReread4 = Empty
                End Select
                If ElapsedTime() - Tstart > Timeout Then Exit Do
            Loop

            Tend = ElapsedTime()
        
            Select Case j
                Case 1
                    NumCalls1 = k
                    t1 = (Tend - Tstart) / k
                Case 2
                    NumCalls2 = k
                    t2 = (Tend - Tstart) / k
                Case 3
                    NumCalls3 = k
                    t3 = (Tend - Tstart) / k
                Case 4
                    NumCalls4 = "Not found"
                    t4 = "Not found"
                    On Error Resume Next
                    NumCalls4 = Application.WorksheetFunction.VLookup(FileName, JuliaResults, 4, False)
                    t4 = Application.WorksheetFunction.VLookup(FileName, JuliaResults, 2, False)
                    On Error GoTo ErrHandler
            End Select
        Next j
    Else
        t1 = Rnd(): t2 = Rnd(): t3 = Rnd(): t4 = Rnd()
        NumCalls1 = 0: NumCalls2 = 0: NumCalls3 = 0: NumCalls4 = 0
    End If

    ReDim Ret(1 To IIf(WithHeaders, 2, 1), 1 To 10) As Variant
    
    DataRow = IIf(WithHeaders, 2, 1)
    
    Ret(DataRow, 1) = t1: If WithHeaders Then Ret(1, 1) = FnName1
    Ret(DataRow, 2) = t2: If WithHeaders Then Ret(1, 2) = FnName2
    Ret(DataRow, 3) = t3: If WithHeaders Then Ret(1, 3) = FnName3
    Ret(DataRow, 4) = t4: If WithHeaders Then Ret(1, 4) = FnName4
    Ret(DataRow, 5) = NumCalls1: If WithHeaders Then Ret(1, 5) = "NCalls" + vbLf + FnName1
    Ret(DataRow, 6) = NumCalls2: If WithHeaders Then Ret(1, 6) = "NCalls" + vbLf + FnName2
    Ret(DataRow, 7) = NumCalls3: If WithHeaders Then Ret(1, 7) = "NCalls" + vbLf + FnName3
    Ret(DataRow, 8) = NumCalls4: If WithHeaders Then Ret(1, 8) = "NCalls" + vbLf + FnName3
    Ret(DataRow, 9) = FileName: If WithHeaders Then Ret(1, 9) = "File"
    Ret(DataRow, 10) = FileSize(FileName): If WithHeaders Then Ret(1, 10) = "Size"

    TimeFourParsers = Ret

    Exit Function
ErrHandler:
    TimeFourParsers = "#TimeFourParsers (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Sub AddCharts(Optional Export As Boolean = True)
    
    Dim c As ChartObject
    Dim N As Name
    Dim prot As Boolean
    Dim ws As Worksheet
    Dim xData As Range
    Dim yData As Range

    On Error GoTo ErrHandler
    
    Set ws = ActiveSheet
    
    prot = ws.ProtectContents
    ws.Unprotect
    
    For Each c In ws.ChartObjects
        c.Delete
    Next

    For Each N In ws.Names
        If InStr(N.Name, "PasteResultsHere") > 1 Then
            Set yData = N.RefersToRange
            With yData
                Set yData = .Offset(-1).Resize(.Rows.Count + 1, 4)
            End With
            Set xData = yData.Offset(, -1).Resize(, 1)
            With xData
                If .Cells(2, 1).value = .Cells(.Rows.Count, 1).value Then
                    Set xData = xData.Offset(, -1)
                End If
            End With
            With xData
                If .Cells(2, 1).value = .Cells(.Rows.Count, 1).value Then
                    Set xData = xData.Offset(, -2)
                End If
            End With
            AddChart xData, yData, Export
        End If
    Next N

    Application.GoTo ws.Cells(1, 1)
    ws.Protect , , prot

    Exit Sub
ErrHandler:
    MsgBox "#AddCharts (line " & CStr(Erl) + "): " & Err.Description & "!"
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

    Const ChartsInCol = "P"
    Const Err_BadSelection = "That selection does not look correct." + vbLf + vbLf + _
        "Select two areas to define the data to plot. The first area should contain " + _
        "the independent data and have a single column with top cell giving the x axis " + _
        "label. The second area should contain the dependent data with one column per data " + _
        "series and top row giving the series names. Both areas should have the same number of rows"
    Const TitlesInCol = "M"
    Dim ch As Chart
    Dim shp As Shape
    Dim SourceData As Range
    Dim Title As String
    Dim TitleCell As Range
    Dim TopLeftCell As Range
    Dim wsh As Worksheet

    On Error GoTo ErrHandler

    If xData Is Nothing Then
        Set SourceData = Selection

        If SourceData.Areas.Count <> 2 Then
            Throw Err_BadSelection
        ElseIf SourceData.Areas(1).Rows.Count <> SourceData.Areas(2).Rows.Count Then
            Throw Err_BadSelection
        End If
        Set xData = SourceData.Areas(1)
        Set yData = SourceData.Areas(2)
        Set wsh = xData.Parent
    Else
        'Actually selecting the ranges seems to be necessary to get the legends to appear in the generated charts...
        Set wsh = xData.Parent
        wsh.Activate
        Range(xData.Address & "," & yData.Address).Select
    End If
    
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
    ch.Axes(xlCategory).AxisTitle.text = xData.Cells(1, 1).value + ". Log Scale"
    ch.ChartTitle.Caption = Title
    With xData
        ch.Axes(xlCategory).MinimumScale = .Cells(2, 1).value
        ch.Axes(xlCategory).MaximumScale = .Cells(.Rows.Count, 1).value
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
        Folder = Left$(ThisWorkbook.path, InStrRev(ThisWorkbook.path, "\")) + "charts\"
        ch.Export Folder + FileName
    End If

    Exit Sub
ErrHandler:
    Throw "#AddChart (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

