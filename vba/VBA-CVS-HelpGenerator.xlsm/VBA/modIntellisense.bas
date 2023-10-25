Attribute VB_Name = "modIntellisense"
Option Explicit

Sub CreateIntellisenseWorkbook()

          Const TargetBookName = "VBA-CSV-Intellisense.xlsx"
          Dim FnName As String
          Dim i As Long
          Dim j As Long
          Dim Prompt As String
          Dim SourceRange As Range
          Dim targetsheet As Worksheet
          Dim wb As Workbook

1         On Error GoTo ErrHandler
2         Prompt = "Create intellisense data and save to " & ThisWorkbook.Path & "\" & TargetBookName

3         If MsgBox(Prompt, vbOKCancel + vbQuestion) <> vbOK Then Exit Sub

4         On Error Resume Next
5         Set wb = Application.Workbooks(TargetBookName)
6         On Error GoTo ErrHandler
7         If Not wb Is Nothing Then
8             Err.Raise vbObjectError + 1, , "Please close workbook " & TargetBookName
9         End If

10        Set wb = Application.Workbooks.Add()
11        Set targetsheet = wb.Worksheets(1)

12        targetsheet.Name = "_Intellisense_"

13        targetsheet.Cells(1, 1).value = "FunctionInfo"
14        targetsheet.Cells(1, 2).value = "'1.0"
15        shHelp.Calculate

16        For i = 1 To 2
17            FnName = Choose(i, "CSVRead", "CSVWrite")
18            Set SourceRange = shHelp.Range(FnName & "Args")
19            targetsheet.Cells(1 + i, 1) = FnName
20            targetsheet.Cells(1 + i, 2) = vbCrLf & Replace(SourceRange.Cells(1, 1).Offset(-2, 1).value, vbLf, vbCrLf) & vbCrLf
21            For j = 1 To SourceRange.Rows.Count
22                targetsheet.Cells(1 + i, 2 * (1 + j)).value = SourceRange.Cells(j, 1).value
23                targetsheet.Cells(1 + i, 1 + 2 * (1 + j)).value = Replace(SourceRange.Cells(j, 4).value, vbLf, vbCrLf) 'Use the "long form" of the help...
24            Next j
25        Next i

26        With targetsheet.UsedRange
27            .Columns.ColumnWidth = 40
28            .WrapText = True
29            .VerticalAlignment = xlVAlignCenter
30            .Columns.AutoFit
31        End With
32        Application.DisplayAlerts = False
33        wb.SaveAs ThisWorkbook.Path & "\" & TargetBookName, xlOpenXMLWorkbook
34        wb.Close False

35        Exit Sub
ErrHandler:
36        MsgBox "#CreateIntellisenseWorkbook (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
