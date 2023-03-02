VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    On Error GoTo ErrHandler
    Dim CSVReadRes As Range
    Dim i As Long
    Dim NewTestNo As Variant
    Dim protContents As Boolean
    Dim R As Range
    Const MAXWIDTH As Long = 50
    Const MINWIDTH As Long = 4

    Set R = ActiveSheet.ListObjects("Tests").DataBodyRange

    If Not Application.Intersect(Target, R) Is Nothing Then
        NewTestNo = Application.Intersect(Target.EntireRow, R.Columns(2)).value
        If Not IsNumeric(NewTestNo) Then Throw "TestNo must be a number"
    End If

    If NewTestNo > 0 Then
        Me.Range("TestNo").value = NewTestNo
        Me.Calculate
        Cancel = True
        On Error Resume Next
        Set CSVReadRes = Me.Range("CallToCSVRead").SpillingToRange
        On Error GoTo ErrHandler
        If CSVReadRes Is Nothing Then
            Set CSVReadRes = Me.Range("CallToCSVRead")
        End If
        
        Application.Goto Me.Cells(1, 1)
        Me.Range("TheVBACode").Select
        
        protContents = Me.ProtectContents

        If protContents Then Me.Unprotect
        With CSVReadRes
            .NumberFormat = "General"
            .Columns.AutoFit
        End With
        
        With Me.Range("FormattedRange1")
            .Columns.AutoFit
            If .Columns(1).ColumnWidth < MINWIDTH Then
                .Columns(1).ColumnWidth = MINWIDTH
            ElseIf .Columns(1).ColumnWidth > MAXWIDTH Then
                .Columns(1).ColumnWidth = MAXWIDTH
            End If
        End With
        
        With Me.Range("FormattedRange2")
            For i = 1 To .Columns.count
                If .Cells(1, i).ColumnWidth < MINWIDTH Then
                    .Columns(i).ColumnWidth = MINWIDTH
                ElseIf .Cells(1, i).ColumnWidth > MAXWIDTH Then
                    .Columns(i).ColumnWidth = MAXWIDTH
                End If
            Next i
        End With
        
        ClearUsedRange
        
        If protContents Then Me.Protect , , True

    End If
    
    Exit Sub
ErrHandler:
    MsgBox ReThrow("Worksheet_BeforeDoubleClick", Err, True), vbCritical
End Sub

Sub ClearUsedRange()
    Dim i As Long
    Dim LastCol As Long
    Dim LastRow As Long
    Dim RealLastCol As Long
    Dim RealLastRow As Long
    Dim UsedRange As Range

    Dim NeedThis As Range

    On Error GoTo ErrHandler
    For i = 1 To 6

        Select Case i
            Case 1, 2, 3
                Set NeedThis = Me.Range("FormattedRange" & i)

            Case 4
                Set NeedThis = Me.Cells.Find(What:="*", After:=Me.Cells(1, 1), LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
            Case 5
                Set NeedThis = Me.Cells.Find(What:="*", After:=Me.Cells(1, 1), LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
            Case 6
                Set NeedThis = Me.ListObjects("Tests").DataBodyRange
        End Select
        
        If Not NeedThis Is Nothing Then
            With NeedThis
                If .Row + .Rows.count - 1 > RealLastRow Then RealLastRow = .Row + .Rows.count - 1
                If .Column + .Columns.count - 1 > RealLastCol Then RealLastCol = .Column + .Columns.count - 1
            End With
        End If
    Next i

    Set UsedRange = Me.UsedRange
    With UsedRange
        LastRow = .Row + .Rows.count - 1
        LastCol = .Column + .Columns.count - 1
    End With

    If LastRow > RealLastRow Then
        Debug.Print Me.Cells(RealLastRow + 1, 1).Resize(LastRow - RealLastRow + 1).EntireRow.Address
        Me.Cells(RealLastRow + 1, 1).Resize(LastRow - RealLastRow + 1).EntireRow.Delete
    End If
    If LastCol > RealLastCol Then
        Debug.Print Me.Cells(1, RealLastCol + 1).Resize(, LastCol - RealLastCol + 1).EntireColumn.Address
        Me.Cells(1, RealLastCol + 1).Resize(, LastCol - RealLastCol + 1).EntireColumn.Delete
    End If

    Set UsedRange = Me.UsedRange

    Exit Sub
ErrHandler:
    ReThrow "ClearUsedRange", Err
End Sub
