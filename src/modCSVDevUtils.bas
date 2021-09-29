Attribute VB_Name = "modCSVDevUtils"

' VBA-CSV

' Copyright (C) 2021 - Philip Swannell (https://github.com/PGS62/VBA-CSV )
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/VBA-CSV#readme

Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SaveWorkbookAndExportModules
' Purpose    : Export the modules of this workbook to the src folder, also save the workbook in its current location,
'              and save a backup of the workbook to my OneDrive folder.
' -----------------------------------------------------------------------------------------------------------------------
Sub SaveWorkbookAndExportModules()

    Const Title = "VBA-CSV"
    Dim AuditData
    Dim BackUpBookName As String
    Dim bExport As Boolean
    Dim c As VBIDE.VBComponent
    Dim FileName As String
    Dim Folder As String
    Dim i As Long
    Dim Prompt As String
    Dim wb As Workbook

    On Error GoTo ErrHandler

    Set wb = ThisWorkbook
    Folder = Left$(ThisWorkbook.path, InStrRev(ThisWorkbook.path, "\")) + "src"

    Prompt = "Save the workbook and export modules to '" + Folder + "'?"
    If MsgBox(Prompt, vbOKCancel + vbQuestion, Title) <> vbOK Then Exit Sub

    If wb.VBProject.Protection = 1 Then
        Throw "VBProject is protected"
        Exit Sub
    End If

    If Right$(Folder, 1) <> "\" Then Folder = Folder + "\"
    On Error Resume Next
    Kill Folder & "*.bas*"
    Kill Folder & "*.cls*"
    On Error GoTo ErrHandler
    
    For Each c In wb.VBProject.VBComponents
        bExport = True
        FileName = c.Name

        Select Case c.Type
            Case vbext_ct_ClassModule
                FileName = FileName & ".cls"
            Case vbext_ct_MSForm
                FileName = FileName & ".frm"
            Case vbext_ct_StdModule
                FileName = FileName & ".bas"
            Case vbext_ct_Document
                If c.CodeModule.CountOfLines <= 2 Then        'Only export sheet module if it contains code. Test CountOfLines <= 2 likely to be good enough in practice -
                    bExport = False
                Else
                    bExport = True
                    FileName = FileName & ".cls"
                End If
            Case Else
                bExport = False
        End Select

        'only export files of the PGS62 project, not those from other _
         CSV parsers that I have imported to compare performance.
        If Left$(FileName, 6) <> "modCSV" Then
            bExport = False
        End If

        If bExport Then
            c.Export Folder & FileName
        End If
    Next c
    
    On Error Resume Next
    Kill Folder & "*.frx"        'These are binary files that we don't want to check in to Git
    On Error GoTo ErrHandler
    
    AuditData = Range(shAudit.Range("Headers").Cells(1, 1), shAudit.Range("Headers").Cells(1, 1).End(xlToRight).End(xlDown))
    For i = LBound(AuditData, 1) + 1 To UBound(AuditData, 1)
        AuditData(i, 3) = CDate(AuditData(i, 3))
    Next
    
    ThrowIfError CSVWrite(AuditData, ThisWorkbook.path & "\AuditSheetComments.csv", True, "dd-mmm-yyyy", "hh:mm:ss")
    
    PrepareForRelease
    ThisWorkbook.Save
    
    BackUpBookName = Environ("OneDriveConsumer") + "\Excel Sheets\VBA-CSV_Backups\" + Replace(ThisWorkbook.Name, ".", "_v" & shAudit.Range("B6") & ".")
    
    FileCopy ThisWorkbook.FullName, BackUpBookName

    Exit Sub
ErrHandler:
    MsgBox "#SaveWorkbookAndExportModules (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : PrepareForRelease
' Purpose    : Tidy up the worksheets of this workbook.
' -----------------------------------------------------------------------------------------------------------------------
Sub PrepareForRelease()

    Dim i As Long
    Dim ws As Worksheet
    
    On Error GoTo ErrHandler

    If Application.DisplayFormulaBar Then Application.FormulaBarHeight = 1

    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            Application.Goto ws.Cells(1, 1)
            If InStr(ws.Name, "GIF") = 0 Then
                ActiveWindow.DisplayGridlines = False
                ActiveWindow.DisplayHeadings = False
            End If
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
