Attribute VB_Name = "modDevUtils"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SaveWorkbookAndExportModules
' Author     : Philip Swannell
' Date       : 29-Jul-2021
' Purpose    : Export the modules of this workbook to the src folder, also save the workbook in its current location
' -----------------------------------------------------------------------------------------------------------------------
Sub SaveWorkbookAndExportModules()

          Dim wb As Workbook
          Dim bExport As Boolean
          Dim C As VBIDE.VBComponent
          Dim FileName As String
          Dim Folder As String
          Dim Prompt As String
          Const Title = "VBA-CSV"

1         On Error GoTo ErrHandler

2         Set wb = ThisWorkbook
3         Folder = ThisWorkbook.Path
4         Folder = Left(Folder, InStrRev(Folder, "\")) + "src"

5         Prompt = "Save the workbook and export modules to '" + Folder + "'?"
6         If MsgBox(Prompt, vbOKCancel + vbQuestion, Title) <> vbOK Then Exit Sub

7         If wb.VBProject.Protection = 1 Then
8             Err.Raise vbObjectError + 1, , "VBProject is protected"
9             Exit Sub
10        End If

11        If Right$(Folder, 1) <> "\" Then Folder = Folder + "\"
12        On Error Resume Next
13        Kill Folder & "*.bas*"
14        Kill Folder & "*.cls*"
15        Kill Folder & "*.frm*"
16        Kill Folder & "*.frx*"
17        On Error GoTo ErrHandler

18        For Each C In wb.VBProject.VBComponents
19            bExport = True
20            FileName = C.Name

21            Select Case C.Type
                  Case vbext_ct_ClassModule
22                    FileName = FileName & ".cls"
23                Case vbext_ct_MSForm
24                    FileName = FileName & ".frm"
25                Case vbext_ct_StdModule
26                    FileName = FileName & ".bas"
27                Case vbext_ct_Document
28                    If C.CodeModule.CountOfLines <= 2 Then        'Only export sheet module if it contains code. Test CountOfLines <= 2 likely to be good enough in practice -
29                        bExport = False
30                    Else
31                        bExport = True
32                        FileName = FileName & ".cls"
33                    End If
34                Case Else
35                    bExport = False
36            End Select

37            If bExport Then
38                C.Export Folder & FileName
39            End If

40        Next C

41        On Error Resume Next
42        Kill Folder & "*.frx"        'These are binary files that we don't want to check in to Git
43        On Error GoTo ErrHandler
44        PrepareForRelease
45        ThisWorkbook.Save


46        Exit Sub
ErrHandler:
47        MsgBox "#SaveWorkbookAndExportModules (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical
End Sub



Sub PrepareForRelease()
          Dim i As Long
          Dim ws As Worksheet
          
1         On Error GoTo ErrHandler

2         If Application.DisplayFormulaBar Then Application.FormulaBarHeight = 1

3         For Each ws In ThisWorkbook.Worksheets
4             If ws.Visible = xlSheetVisible Then
5                 Application.Goto ws.Cells(1, 1)
6                 ActiveWindow.DisplayGridlines = False
7                 ActiveWindow.DisplayHeadings = False
8             End If
9             ws.Protect , True, True
10        Next
11        For i = 1 To ThisWorkbook.Worksheets.Count
12            If ThisWorkbook.Worksheets(i).Visible Then
13                Application.Goto ThisWorkbook.Worksheets(i).Cells(1, 1)
14                Exit For
15            End If
16        Next i
17        Exit Sub
ErrHandler:
18        Throw "#PrepareForRelease (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

