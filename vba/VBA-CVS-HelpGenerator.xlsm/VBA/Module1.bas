Attribute VB_Name = "Module1"
'Option Explicit
'
'' -----------------------------------------------------------------------------------------------------------------------
'' Procedure  : SaveAndBackupWorkbook
'' Purpose    : Export the modules of this workbook to the src folder, also save the workbook in its current location,
''              and save a backup of the workbook to my OneDrive folder.
'' -----------------------------------------------------------------------------------------------------------------------
'Sub SaveAndBackupWorkbook()
'
'          Const Title = "VBA-CSV"
'          Dim BackUpBookName As String
'
'1         On Error GoTo ErrHandler
'
'2         PrepareForRelease
'3         CreateIntellisenseWorkbook
'4         ThisWorkbook.Save
'
'5         BackUpBookName = Environ("OneDriveConsumer") + "\Excel Sheets\VBA-CSV_Backups\" + Replace(ThisWorkbook.Name, ".", "_v" & shAudit.Range("B6") & ".")
'
'6         FileCopy ThisWorkbook.FullName, BackUpBookName
'
'7         Exit Sub
'ErrHandler:
'8         MsgBox "#SaveAndBackupWorkbook (line " & CStr(Erl) + "): " & Err.Description & "!", , Title
'End Sub
'
'' -----------------------------------------------------------------------------------------------------------------------
'' Procedure  : PrepareForRelease
'' Purpose    : Tidy up the worksheets of this workbook.
'' -----------------------------------------------------------------------------------------------------------------------
'Sub PrepareForRelease()
'
'    Dim i As Long
'    Dim ws As Worksheet
'
'    On Error GoTo ErrHandler
'
'    If Application.DisplayFormulaBar Then Application.FormulaBarHeight = 1
'
'    For Each ws In ThisWorkbook.Worksheets
'        If ws.Visible = xlSheetVisible Then
'            Application.GoTo ws.Cells(1, 1)
'            ActiveWindow.DisplayGridlines = False
'            ActiveWindow.DisplayHeadings = False
'        End If
'        ws.Protect , True, True
'    Next
'    For i = 1 To ThisWorkbook.Worksheets.Count
'        If ThisWorkbook.Worksheets(i).Visible Then
'            Application.GoTo ThisWorkbook.Worksheets(i).Cells(1, 1)
'            Exit For
'        End If
'    Next i
'    Exit Sub
'ErrHandler:
'    Throw "#PrepareForRelease: " & Err.Description & "!"
'End Sub
'
'Function FileCopy(SourceFile As String, TargetFile As String)
'    Dim F As Scripting.File
'    Dim FSO As Scripting.FileSystemObject
'    Dim CopyOfErr As String
'    On Error GoTo ErrHandler
'    Set FSO = New FileSystemObject
'    Set F = FSO.GetFile(SourceFile)
'    F.Copy TargetFile, True
'    FileCopy = TargetFile
'    Set FSO = Nothing: Set F = Nothing
'    Exit Function
'ErrHandler:
'    CopyOfErr = Err.Description
'    Set FSO = Nothing: Set F = Nothing
'    Throw "#" + CopyOfErr + "!"
'End Function
'
