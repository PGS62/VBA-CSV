Attribute VB_Name = "modDevUtils"
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
    Dim c As VBIDE.VBComponent
    Dim FileName As String
    Dim Folder As String
    Dim i As Long
    Dim Prompt As String
    Dim wb As Workbook

    On Error GoTo ErrHandler

    Set wb = ThisWorkbook
    Folder = ThisWorkbook.path
    Folder = Left(Folder, InStrRev(Folder, "\")) + "src"

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
    Kill Folder & "*.frm*"
    Kill Folder & "*.frx*"
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
    
    ThrowIfError CSVWrite(ThisWorkbook.path & "\AuditSheetComments.csv", AuditData, True, "dd-mmm-yyyy", "hh:mm:ss")
    
    PrepareForRelease
    ThisWorkbook.Save
    
    
    
    BackUpBookName = Environ("OneDriveConsumer") + "\Excel Sheets\VBA-CSV_Backups\" + Replace(ThisWorkbook.Name, ".", "_v" & shAudit.Range("B6") & ".")
    
    ThrowIfError Application.Run("sfilecopy", ThisWorkbook.FullName, BackUpBookName)

    Exit Sub
ErrHandler:
    MsgBox "#SaveWorkbookAndExportModules: " & Err.Description & "!", vbCritical
End Sub

Sub PrepareForRelease()

    Dim i As Long
    Dim ws As Worksheet
    
    On Error GoTo ErrHandler

    If Application.DisplayFormulaBar Then Application.FormulaBarHeight = 1

    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            Application.GoTo ws.Cells(1, 1)
            ActiveWindow.DisplayGridlines = False
            ActiveWindow.DisplayHeadings = False
        End If
        ws.Protect , True, True
    Next
    For i = 1 To ThisWorkbook.Worksheets.count
        If ThisWorkbook.Worksheets(i).Visible Then
            Application.GoTo ThisWorkbook.Worksheets(i).Cells(1, 1)
            Exit For
        End If
    Next i
    Exit Sub
ErrHandler:
    Throw "#PrepareForRelease: " & Err.Description & "!"
End Sub


