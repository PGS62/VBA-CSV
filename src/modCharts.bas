Attribute VB_Name = "modCharts"
Option Explicit

Sub AddChart2()

    Dim ch As Chart
    Dim sh As Shape
    Dim Title As String
    Dim TitleCell As Range
    Dim TopLeftCell As Range

    On Error GoTo ErrHandler

    Set sh = ActiveSheet.Shapes.AddChart2(240, xlXYScatterLines)
    Set ch = sh.Chart
    Set TopLeftCell = Application.Intersect(Selection.Areas(1).Cells(1, 1).EntireRow, ActiveSheet.Range("P:P"))
    Set TitleCell = Application.Intersect(Selection.Areas(1).Cells(0, 1).EntireRow, ActiveSheet.Range("M:M"))

    Title = "='" & ActiveSheet.Name & "'!R" & TitleCell.Row & "C" & TitleCell.Column

    ch.SetSourceData Selection
    ch.Axes(xlCategory).ScaleType = xlLogarithmic
    ch.Axes(xlValue).ScaleType = xlLogarithmic
    ch.Axes(xlValue, xlPrimary).HasTitle = True
    ch.Axes(xlValue, xlPrimary).AxisTitle.text = "Seconds to read. Log Scale"
    ch.Axes(xlCategory).HasTitle = True
    ch.Axes(xlCategory).AxisTitle.text = Selection.Areas(1).Cells(1, 1).value + ". Log Scale"
    ch.ChartTitle.Caption = Title

    sh.Top = TopLeftCell.Top
    sh.Left = TopLeftCell.Left
    sh.Height = 337
    sh.Width = 561

    Exit Sub
ErrHandler:
    MsgBox "#AddChart2: " & Err.Description & "!"
End Sub

Sub RunSpeedTests()

    Const Timeout = 5
    Dim c As Range
    Dim n As Name
    Dim TestResults As Variant

    On Error GoTo ErrHandler

    If MsgBox("Run Speed Tests?", vbOKCancel + vbQuestion) <> vbOK Then Exit Sub
    For Each n In ActiveSheet.Names
        If InStr(n.Name, "PutFormulasHere") > 1 Then
            Application.GoTo n.RefersToRange
            For Each c In n.RefersToRange.Cells
                c.Resize(1, 13).ClearContents
                TestResults = TimeThreeParsers(c.Offset(0, -3).value, c.Offset(0, -2).value, c.Offset(0, -1).value, Timeout, False)
                c.Resize(1, 10).value = TestResults
                ActiveSheet.Calculate
                Application.ScreenUpdating = True
            Next
        End If
    Next n

    Exit Sub

    Exit Sub
ErrHandler:
    MsgBox "#RunSpeedTests (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical
End Sub
