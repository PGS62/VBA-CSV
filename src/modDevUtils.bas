Attribute VB_Name = "modDevUtils"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ExportModules
' Author     : Philip Swannell
' Date       : 29-Jul-2021
' Purpose    : Export the modules of this workbook to the src folder.
' -----------------------------------------------------------------------------------------------------------------------
Sub ExportModules()

          Dim wb As Workbook
          Dim bExport As Boolean
          Dim C As VBIDE.VBComponent
          Dim FileName As String
          Dim Folder As String

1         On Error GoTo ErrHandler

2         Set wb = ThisWorkbook
3         Folder = ThisWorkbook.Path
4         Folder = Left(Folder, InStrRev(Folder, "\")) + "src"

5         If wb.VBProject.Protection = 1 Then
6             Err.Raise vbObjectError + 1, , "VBProject is protected"
7             Exit Sub
8         End If

9         If Right$(Folder, 1) <> "\" Then Folder = Folder + "\"
10        On Error Resume Next
11        Kill Folder & "*.bas*"
12        Kill Folder & "*.cls*"
13        Kill Folder & "*.frm*"
14        Kill Folder & "*.frx*"
15        On Error GoTo ErrHandler

16        For Each C In wb.VBProject.VBComponents
17            bExport = True
18            FileName = C.Name

19            Select Case C.Type
                  Case vbext_ct_ClassModule
20                    FileName = FileName & ".cls"
21                Case vbext_ct_MSForm
22                    FileName = FileName & ".frm"
23                Case vbext_ct_StdModule
24                    FileName = FileName & ".bas"
25                Case vbext_ct_Document
26                    If C.CodeModule.CountOfLines <= 2 Then        'Only export sheet module if it contains code. Test CountOfLines <= 2 likely to be good enough in practice -
27                        bExport = False
28                    Else
29                        bExport = True
30                        FileName = FileName & ".cls"
31                    End If
32                Case Else
33                    bExport = False
34            End Select

35            If bExport Then
36                C.Export Folder & FileName
37            End If

38        Next C

39        On Error Resume Next
40        Kill Folder & "*.frx"        'These are binary files that we don't want to check in to Git
41        On Error GoTo ErrHandler

42        Exit Sub
ErrHandler:
43        MsgBox "#ExportModules (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical
End Sub

