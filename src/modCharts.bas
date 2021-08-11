Attribute VB_Name = "modCharts"
Option Explicit

Sub AddChart2()
          Dim ch As Chart
          Dim sh As Shape
          Dim topLeftCell As Range
          Dim TitleCell As Range
          Dim Title As String

1         On Error GoTo ErrHandler
2         Set sh = ActiveSheet.Shapes.AddChart2(240, xlXYScatterLines)
3         Set ch = sh.Chart
4         Set topLeftCell = Application.Intersect(Selection.Areas(1).Cells(1, 1).EntireRow, ActiveSheet.Range("T:T"))
5         Set TitleCell = Application.Intersect(Selection.Areas(1).Cells(0, 1).EntireRow, ActiveSheet.Range("O:O"))
6         Title = "='" & ActiveSheet.Name & "'!R" & TitleCell.Row & "C" & TitleCell.Column
          


7         ch.SetSourceData Selection
8         ch.Axes(xlCategory).ScaleType = xlLogarithmic
9         ch.Axes(xlValue).ScaleType = xlLogarithmic

10        ch.Axes(xlValue, xlPrimary).HasTitle = True
11        ch.Axes(xlValue, xlPrimary).AxisTitle.text = "Seconds to read. Log Scale"
          
12        ch.Axes(xlCategory).HasTitle = True
13        ch.Axes(xlCategory).AxisTitle.text = Selection.Areas(1).Cells(1, 1).value + ". Log Scale"
          
14        ch.ChartTitle.Caption = Title


15        sh.Top = topLeftCell.Top
16        sh.Left = topLeftCell.Left
17        sh.Height = 337
18        sh.Width = 561

19        Exit Sub
ErrHandler:
20        MsgBox "#AddChart2 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub


Sub RunSpeedTests()
          Dim c As Range
          Dim n As Name

1         On Error GoTo ErrHandler

2         If MsgBox("Run Speed Tests?", vbOKCancel) <> vbOK Then Exit Sub
3         For Each n In ActiveSheet.Names
4             If InStr(n.Name, "PutFormulasHere") > 1 Then
5                 Application.GoTo n.RefersToRange
6                 For Each c In n.RefersToRange.Cells
7                     c.Resize(1, 13).ClearContents
8                     c.Formula2R1C1 = "=TimeSixParsers(RC[-4],RC[-3],RC[-2],RC[-1])"
9                     With c.SpillingToRange
10                        .value = .value
11                    End With
12                    ActiveSheet.Calculate
13                    Application.ScreenUpdating = True
14                Next
15            End If
16        Next n


17        Exit Sub
ErrHandler:
18        SomethingWentWrong "#RunSpeedTests (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub


Sub dfbsdfbs()

Dim n As Name

For Each n In ActiveSheet.Names
Debug.Print n.Name
Next



End Sub
