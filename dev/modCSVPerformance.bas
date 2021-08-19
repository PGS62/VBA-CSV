Attribute VB_Name = "modCSVPerformance"
Option Explicit

'Code of the "Performance" worksheet

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RunSpeedTests
' Purpose    : Attached to the "Run Speed Tests..." button. Note the significance of the "PutFormulasHere" ranges
' -----------------------------------------------------------------------------------------------------------------------
Sub RunSpeedTests()

    Const Timeout = 5
    Dim C As Range
    Dim N As Name
    Dim TestResults As Variant
    Const NumColsInTTPRet = 10
    Dim Prompt As String

    On Error GoTo ErrHandler
    
    Prompt = "Run Speed tests?" + vbLf + vbLf + _
    "Note this will generate approx 227MB of files in folder" + vbLf + _
    Environ("Temp") & "\VBA-CSV\Performance"

    If MsgBox(Prompt, vbOKCancel + vbQuestion) <> vbOK Then Exit Sub
    
    ActiveSheet.Protect , , False
    
    For Each N In ActiveSheet.Names
        If InStr(N.Name, "PutFormulasHere") > 1 Then
            Application.Goto N.RefersToRange
            For Each C In N.RefersToRange.Cells
                C.Resize(1, NumColsInTTPRet).ClearContents
                TestResults = TimeThreeParsers(C.Offset(0, -3).value, C.Offset(0, -2).value, C.Offset(0, -1).value, Timeout, False)
                C.Resize(1, NumColsInTTPRet).value = TestResults
                ActiveSheet.Calculate
                Application.ScreenUpdating = True
            Next
        End If
    Next N

    ActiveSheet.Protect , , True

    Exit Sub

    Exit Sub
ErrHandler:
    MsgBox "#RunSpeedTests (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : TimeThreeParsers
' Purpose    : Core of the method RunSpeedTests. Note the functions being timed are called many times in a loop that exits
'              after TimeOut seconds have elapsed. Leads to much more reliable timings than timing a single call.
' -----------------------------------------------------------------------------------------------------------------------
Function TimeThreeParsers(EachFieldContains As Variant, NumRows As Long, NumCols As Long, _
    Optional Timeout As Double = 1, Optional WithHeaders As Boolean)

    Const Unicode = False
    Dim data As Variant
    Dim DataReread1
    Dim DataReread2
    Dim DataReread3
    Dim DataRow As Long
    Dim ExtraInfo As String
    Dim FileName As String
    Dim FnName1 As String
    Dim FnName2 As String
    Dim FnName3 As String
    Dim j As Long
    Dim k As Double
    Dim NumCalls1 As Long
    Dim NumCalls2 As Long
    Dim NumCalls3 As Long
    Dim OS As String
    Dim Ret As Variant
    Dim t1 As Double
    Dim t2 As Double
    Dim t3 As Double
    Dim Tend As Double
    Dim Tstart As Double
    Dim Folder As String

    On Error GoTo ErrHandler
    OS = "Windows"
    
    If VarType(EachFieldContains) = vbDouble Then
        ExtraInfo = "Doubles"
    ElseIf VarType(EachFieldContains) = vbString Then
        If Left(EachFieldContains, 1) = """" & Right(EachFieldContains, 1) = """" Then
            ExtraInfo = "Quoted_Strings_length_" & Len(EachFieldContains)
        Else
            ExtraInfo = "Strings_length_" & Len(EachFieldContains)
        End If
    Else
        ExtraInfo = "Unknown"
    End If

    Folder = Environ("Temp") & "\VBA-CSV\Performance"

    ThrowIfError CreatePath(Folder)

    data = sFill(EachFieldContains, NumRows, NumCols)
    FileName = NameThatFile(Folder, OS, NumRows, NumCols, Replace(ExtraInfo, " ", "-"), Unicode, False)
    ThrowIfError Application.Run("sFileSave", FileName, data, ",", , , , True)
        
    For j = 1 To 6
        k = 0
        Tstart = sElapsedTime()
        Do
            k = k + 1
            Select Case j
                Case 1
                    FnName1 = "CSVRead"
                    DataReread1 = ThrowIfError(CSVRead(FileName, False, ",", , , , , , , , , Unicode))
                Case 2
                    FnName2 = "CSVRead_sdkn104"
                    DataReread2 = ThrowIfError(CSVRead_sdkn104(FileName, Unicode))
                Case 3
                    FnName3 = "CSVRead_ws_garcia"
                    DataReread3 = ThrowIfError(CSVRead_ws_garcia(FileName, ",", vbCrLf))
            End Select
            If sElapsedTime() - Tstart > Timeout Then Exit Do
        Loop

        Tend = sElapsedTime()
        
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
        End Select
    Next j

    ReDim Ret(1 To IIf(WithHeaders, 2, 1), 1 To 8) As Variant
    
    DataRow = IIf(WithHeaders, 2, 1)
    
    Ret(DataRow, 1) = t1: If WithHeaders Then Ret(1, 1) = FnName1
    Ret(DataRow, 2) = t2: If WithHeaders Then Ret(1, 2) = FnName2
    Ret(DataRow, 3) = t3: If WithHeaders Then Ret(1, 3) = FnName3
    Ret(DataRow, 4) = NumCalls1: If WithHeaders Then Ret(1, 4) = "NCalls" + vbLf + FnName1
    Ret(DataRow, 5) = NumCalls2: If WithHeaders Then Ret(1, 5) = "NCalls" + vbLf + FnName2
    Ret(DataRow, 6) = NumCalls3: If WithHeaders Then Ret(1, 6) = "NCalls" + vbLf + FnName3
    Ret(DataRow, 7) = FileName: If WithHeaders Then Ret(1, 7) = "File"
    Ret(DataRow, 8) = FileSize(FileName): If WithHeaders Then Ret(1, 8) = "Size"

    TimeThreeParsers = Ret

    Exit Function
ErrHandler:
    TimeThreeParsers = "#TimeThreeParsers: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : AddChart
' Purpose    : Adds a chart to the sheet Timings. First select the data to plot then run this macro by clicking the
'              "Add Chart" button
' -----------------------------------------------------------------------------------------------------------------------
Sub AddChart()

    Const ChartsInCol = "N"
    Const Err_BadSelection = "That selection does not look correct." + vbLf + vbLf + _
        "Select two areas to define the data to plot. The first area should contain the independent data and have a single column with top cell giving the x axis label. The second area should contain the dependent data with one column per data series and top row giving the series names. Both areas should have the same number of rows"
    Const TitlesInCol = "K"
    Dim ch As Chart
    Dim shp As Shape
    Dim SourceData As Range
    Dim Title As String
    Dim TitleCell As Range
    Dim TopLeftCell As Range
    Dim wsh As Worksheet

    On Error GoTo ErrHandler

    Set SourceData = Selection

    If SourceData.Areas.Count <> 2 Then
        Throw Err_BadSelection
    ElseIf SourceData.Areas(1).Rows.Count <> SourceData.Areas(2).Rows.Count Then
        Throw Err_BadSelection
    End If

    Set wsh = SourceData.Parent
    Set shp = wsh.Shapes.AddChart2(240, xlXYScatterLines)
    Set ch = shp.Chart
    ch.SetSourceData Source:=SourceData
    Set TopLeftCell = Application.Intersect(Selection.Areas(1).Cells(1, 1).EntireRow, wsh.Range(ChartsInCol & ":" & ChartsInCol))
    Set TitleCell = Application.Intersect(Selection.Areas(1).Cells(0, 1).EntireRow, wsh.Range(TitlesInCol & ":" & TitlesInCol))

    Title = "='" & wsh.Name & "'!R" & TitleCell.Row & "C" & TitleCell.Column

    ch.Axes(xlCategory).ScaleType = xlLogarithmic
    ch.Axes(xlValue).ScaleType = xlLogarithmic
    ch.Axes(xlValue, xlPrimary).HasTitle = True
    ch.Axes(xlValue, xlPrimary).AxisTitle.text = "Seconds to read. Log Scale"
    ch.Axes(xlCategory).HasTitle = True
    ch.Axes(xlCategory).AxisTitle.text = Selection.Areas(1).Cells(1, 1).value + ". Log Scale"
    ch.ChartTitle.Caption = Title

    shp.Top = TopLeftCell.Top
    shp.Left = TopLeftCell.Left
    shp.Height = 337
    shp.Width = 561

    Exit Sub
ErrHandler:
    MsgBox "#AddChart: " & Err.Description & "!", vbCritical
End Sub
