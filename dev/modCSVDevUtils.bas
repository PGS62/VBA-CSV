Attribute VB_Name = "modCSVDevUtils"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SaveWorkbookAndExportModules
' Purpose    : Export the modules of this workbook to the src folder, also save the workbook in its current location, and save a backup of the workbook to my OneDrive folder.
' -----------------------------------------------------------------------------------------------------------------------
Sub SaveWorkbookAndExportModules()

          Const Title = "VBA-CSV"
          Dim AuditData
          Dim BackUpBookName
          Dim bExport As Boolean
          Dim C As VBIDE.VBComponent
          Dim FileName As String
          Dim Folder As String
          Dim Folder2 As String
          Dim i As Long
          Dim Prompt As String
          Dim wb As Workbook

1         On Error GoTo ErrHandler

2         Set wb = ThisWorkbook
3         Folder = Left(ThisWorkbook.path, InStrRev(ThisWorkbook.path, "\")) + "src"
4         Folder2 = Left(ThisWorkbook.path, InStrRev(ThisWorkbook.path, "\")) + "dev"

5         Prompt = "Save the workbook and export modules to '" + Folder + "'?"
6         If MsgBox(Prompt, vbOKCancel + vbQuestion, Title) <> vbOK Then Exit Sub

7         If wb.VBProject.Protection = 1 Then
8             Throw "VBProject is protected"
9             Exit Sub
10        End If

11        If Right$(Folder, 1) <> "\" Then Folder = Folder + "\"
12        If Right$(Folder2, 1) <> "\" Then Folder2 = Folder2 + "\"
13        On Error Resume Next
14        Kill Folder & "*.bas*"
15        Kill Folder & "*.cls*"
16        Kill Folder2 & "*.bas*"
17        Kill Folder2 & "*.cls*"
18        On Error GoTo ErrHandler
          
          'No longer export all modules

19        For Each C In wb.VBProject.VBComponents
20            bExport = True
21            FileName = C.Name

22            Select Case C.Type
                  Case vbext_ct_ClassModule
23                    FileName = FileName & ".cls"
24                Case vbext_ct_MSForm
25                    FileName = FileName & ".frm"
26                Case vbext_ct_StdModule
27                    FileName = FileName & ".bas"
28                Case vbext_ct_Document
29                    If C.CodeModule.CountOfLines <= 2 Then        'Only export sheet module if it contains code. Test CountOfLines <= 2 likely to be good enough in practice -
30                        bExport = False
31                    Else
32                        bExport = True
33                        FileName = FileName & ".cls"
34                    End If
35                Case Else
36                    bExport = False
37            End Select

              'only export files of the PGS62 project, not those from other CSV parsers that I have imported for perfromance comparison.
38            If Left(FileName, 6) <> "modCSV" Then
39                bExport = False
40            End If

41            If bExport Then
42                If FileName = "modCSVReadWrite.bas" Then
43                    C.Export Folder & FileName
44                Else
45                    C.Export Folder2 & FileName
46                End If
47            End If
48        Next C
          
49        On Error Resume Next
50        Kill Folder & "*.frx"        'These are binary files that we don't want to check in to Git
51        On Error GoTo ErrHandler
          
52        AuditData = Range(shAudit.Range("Headers").Cells(1, 1), shAudit.Range("Headers").Cells(1, 1).End(xlToRight).End(xlDown))
53        For i = LBound(AuditData, 1) + 1 To UBound(AuditData, 1)
54            AuditData(i, 3) = CDate(AuditData(i, 3))
55        Next
          
56        ThrowIfError CSVWrite(ThisWorkbook.path & "\AuditSheetComments.csv", AuditData, True, "dd-mmm-yyyy", "hh:mm:ss")
          
57        PrepareForRelease
58        ThisWorkbook.Save
          
59        BackUpBookName = Environ("OneDriveConsumer") + "\Excel Sheets\VBA-CSV_Backups\" + Replace(ThisWorkbook.Name, ".", "_v" & shAudit.Range("B6") & ".")
          
60        ThrowIfError Application.Run("sfilecopy", ThisWorkbook.FullName, BackUpBookName)

61        Exit Sub
ErrHandler:
62        MsgBox "#SaveWorkbookAndExportModules (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Sub PrepareForRelease()

    Dim i As Long
    Dim ws As Worksheet
    
    On Error GoTo ErrHandler

    If Application.DisplayFormulaBar Then Application.FormulaBarHeight = 1

    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            Application.Goto ws.Cells(1, 1)
            ActiveWindow.DisplayGridlines = False
            ActiveWindow.DisplayHeadings = False
        End If
        ws.Protect , True, True
    Next
    For i = 1 To ThisWorkbook.Worksheets.Count
        If ThisWorkbook.Worksheets(i).Visible Then
            Application.Goto ThisWorkbook.Worksheets(i).Cells(1, 1)
            Exit For
        End If
    Next i
    Exit Sub
ErrHandler:
    Throw "#PrepareForRelease: " & Err.Description & "!"
End Sub
