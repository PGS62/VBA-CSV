Attribute VB_Name = "modCharts"
Option Explicit

Sub AddChart2()
    Dim ch As Chart
    Dim sh As Shape
    Dim topLeftCell As Range

    On Error GoTo ErrHandler
    Set sh = ActiveSheet.Shapes.AddChart2(240, xlXYScatterLines)
    Set ch = sh.Chart
    Set topLeftCell = Application.Intersect(Selection.Areas(1).Cells(1, 1).EntireRow, ActiveSheet.Range("R:R"))

    ch.SetSourceData Selection
    ch.Axes(xlCategory).ScaleType = xlLogarithmic
    ch.Axes(xlValue).ScaleType = xlLogarithmic

    sh.Top = topLeftCell.Top
    sh.Left = topLeftCell.Left
    sh.Height = 337
    sh.Width = 561

    Exit Sub
ErrHandler:
    MsgBox "#AddChart2 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

