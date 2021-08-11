Attribute VB_Name = "ws_garcia"
Option Explicit

'Wrap to https://github.com/ws-garcia/VBA-CSV-interface 3.1.5
Function CSVRead_ws_garcia(FileName As String, Delimiter As String, ByVal EOL As String)

    Dim oArray() As Variant
    Dim CSVint As CSVinterface

    On Error GoTo ErrHandler

    EOL = OStoEOL(EOL, "EOL")

    Set CSVint = New CSVinterface
    With CSVint.parseConfig
        .path = FileName        ' Full path to the file, including its extension.
        .fieldsDelimiter = Delimiter         ' Columns delimiter
        .recordsDelimiter = EOL     ' Rows delimiter
    End With
    With CSVint
        .ImportFromCSV .parseConfig    ' Import the CSV to internal object
        .DumpToArray oArray
    End With

    CSVRead_ws_garcia = oArray

    Exit Function
ErrHandler:
    CSVRead_ws_garcia = "#CSVRead_ws_garcia: " & Err.Description & "!"
End Function


Private Function OStoEOL(OS As String, ArgName As String) As String

    Const Err_Invalid = " must be one of ""Windows"", ""Unix"" or ""Mac"", or the associented end of line characters."

    Select Case LCase(OS)
        Case "windows", vbCrLf
            OStoEOL = vbCrLf
        Case "unix", vbLf
            OStoEOL = vbLf
        Case "mac", vbCr
            OStoEOL = vbCr
        Case Else
            Throw ArgName + Err_Invalid
    End Select
End Function

