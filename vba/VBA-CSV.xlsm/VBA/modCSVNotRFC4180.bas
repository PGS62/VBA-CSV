Attribute VB_Name = "modCSVNotRFC4180"
Option Explicit

'Convenience function to compare parser's handling of non-RFC4180 compliant input, with those of _
the Java library FastCSV (and other Java libraries). See notation used at https://github.com/osiegmar/JavaCsvComparison

Function TestNonStandardInput(ByVal InputString As String, IgnoreEmptyLines As Boolean, ParserName As String, ExpectedResult As String)

    Dim i As Long
    Dim j As Long
    Dim NC As Long
    Dim NR As Long
    Dim OneDArray() As String
    Dim out As String
    Dim ParserRes As Variant

    On Error GoTo ErrHandler

    Dim LFSymbol As String: LFSymbol = ChrW(9226)
    Dim CRSymbol As String: CRSymbol = ChrW(9229)
    Dim NewRowSymbol As String: NewRowSymbol = ChrW(9166)
    Dim NewColumnSymbol As String: NewColumnSymbol = ChrW(8631)
    Dim EmptyFieldSymbol As String: EmptyFieldSymbol = ChrW(9711)
    Dim SpaceSymbol As String: SpaceSymbol = ChrW(9251)
    Dim NotSupportedSymbol As String: NotSupportedSymbol = ChrW(10134)
    Dim CorrectResultSymbol As String: CorrectResultSymbol = ChrW(9989)
    Dim CrashesSymbol As String: CrashesSymbol = WorksheetFunction.Unichar(128165)

    InputString = Replace(InputString, LFSymbol, vbLf)
    InputString = Replace(InputString, CRSymbol, vbCr)
    InputString = Replace(InputString, SpaceSymbol, " ")

    If ParserName = "CSVRead" Then
        ParserRes = CSVRead(InputString, IgnoreEmptyLines:=IgnoreEmptyLines, ShowMissingsAs:=EmptyFieldSymbol, Delimiter:=",")
    ElseIf ParserName = "sdkn104" Then
        If IgnoreEmptyLines Then
            TestNonStandardInput = NotSupportedSymbol
            Exit Function
        End If

        ParserRes = ParseCSVToArray(InputString, True)
        If IsNull(ParserRes) Then
            ' Debug.Print Err.Description
            TestNonStandardInput = CrashesSymbol
            Exit Function
        End If

        For i = LBound(ParserRes, 1) To UBound(ParserRes, 1)
            For j = LBound(ParserRes, 2) To UBound(ParserRes, 2)
                If ParserRes(i, j) = "" Then
                    ParserRes(i, j) = EmptyFieldSymbol
                End If
            Next
        Next
    ElseIf ParserName = "ws_garcia" Then
        Dim FSO As New Scripting.FileSystemObject
        Dim t As Scripting.TextStream
        Set t = FSO.OpenTextFile("c:\Temp\temp.txt", ForWriting, True)
        t.Write InputString
        t.Close
    
        ParserRes = Wrap_ws_garcia("c:\temp\temp.txt", ",", vbLf, IgnoreEmptyLines, True)
        If NumDimensions(ParserRes) = 0 Then
            TestNonStandardInput = CrashesSymbol
            Exit Function
        End If
    
        For i = LBound(ParserRes, 1) To UBound(ParserRes, 1)
            For j = LBound(ParserRes, 2) To UBound(ParserRes, 2)
                If ParserRes(i, j) = "" Then
                    ParserRes(i, j) = EmptyFieldSymbol
                End If
            Next
        Next

    Else
        Throw "ParserName not recognised"
    End If

    NR = NRows(ParserRes)
    NC = NCols(ParserRes)
    ReDim OneDArray(LBound(ParserRes, 1) To UBound(ParserRes, 1))

    For i = LBound(ParserRes, 1) To UBound(ParserRes, 1)
        OneDArray(i) = ParserRes(i, LBound(ParserRes, 1))
        For j = LBound(ParserRes, 2) + 1 To UBound(ParserRes, 2)
            OneDArray(i) = OneDArray(i) & NewColumnSymbol & ParserRes(i, j)
        Next j
    Next i

    out = VBA.Join(OneDArray, NewRowSymbol)
    out = Replace(out, vbLf, LFSymbol)
    out = Replace(out, vbCr, CRSymbol)
    out = Replace(out, " ", SpaceSymbol)
    If out = ExpectedResult Then out = CorrectResultSymbol
    TestNonStandardInput = out

    Exit Function
ErrHandler:
    TestNonStandardInput = ReThrow("TestNonStandardInput", Err, True)
End Function

Function SaveNotCompliantFile(ByVal InputString As String, FileName As String)

    On Error GoTo ErrHandler
    On Error GoTo ErrHandler

    Dim LFSymbol As String: LFSymbol = ChrW(9226)
    Dim CRSymbol As String: CRSymbol = ChrW(9229)
    Dim SpaceSymbol As String: SpaceSymbol = ChrW(9251)

    InputString = Replace(InputString, LFSymbol, vbLf)
    InputString = Replace(InputString, CRSymbol, vbCr)
    InputString = Replace(InputString, SpaceSymbol, " ")

    Dim FSO As New Scripting.FileSystemObject
    Dim t As Scripting.TextStream
    Set t = FSO.OpenTextFile(FileName, ForWriting, True)
    t.Write InputString
    t.Close

    SaveNotCompliantFile = FileName

    Exit Function
ErrHandler:
    SaveNotCompliantFile = ReThrow("SaveNotCompliantFile", Err, True)
End Function

