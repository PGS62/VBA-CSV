VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    On Error GoTo ErrHandler
    RegisterCSVRead
    RegisterCSVWrite

    Exit Sub
ErrHandler:
    ReThrow "Workbook_Open", Err
End Sub
