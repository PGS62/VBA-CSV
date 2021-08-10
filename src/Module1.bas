Attribute VB_Name = "Module1"
Option Explicit

Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    ActiveChart.SetElement (msoElementPrimaryValueAxisTitleAdjacentToAxis)
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.text = "Axis TitleNum Rows"
End Sub
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    ActiveChart.ChartTitle.Select
    Application.CutCopyMode = False
    Selection.Caption = "='TimingResults (3)'!R4C15"
    ActiveChart.Axes(xlValue).AxisTitle.Select
    Application.CutCopyMode = False
    Selection.Caption = "='TimingResults (3)'!R5C3"
    Range("U33").Select
End Sub
