Attribute VB_Name = "modCSVNotRFC4180"
Option Explicit

'Convenience function to compare parser's handling of non-RFC4180 compliant input, with those of _
the Java library FastCSV (and other Java libraries). See notation used at https://github.com/osiegmar/JavaCsvComparison

Function TestNonStandardInput(ByVal InputString As String, IgnoreEmptyLines As Boolean, Optional ParserName As String = "CSVRead")

          Dim i As Long
          Dim j As Long
          Dim ParserRes As Variant
          Dim OneDArray() As String
          Dim NR As Long, NC As Long
          Dim Out As String

1         On Error GoTo ErrHandler

2         Dim LFSymbol As String: LFSymbol = ChrW(9226)
3         Dim CRSymbol As String: CRSymbol = ChrW(9229)
4         Dim NewRowSymbol As String: NewRowSymbol = ChrW(9166)
5         Dim NewColumnSymbol As String: NewColumnSymbol = ChrW(8631)
6         Dim EmptyFieldSymbol As String: EmptyFieldSymbol = ChrW(9711)
7         Dim SpaceSymbol As String: SpaceSymbol = ChrW(9251)

8         InputString = Replace(InputString, LFSymbol, vbLf)
9         InputString = Replace(InputString, CRSymbol, vbCr)
10        InputString = Replace(InputString, SpaceSymbol, " ")

11        If ParserName = "CSVRead" Then
12            ParserRes = CSVRead(InputString, IgnoreEmptyLines:=IgnoreEmptyLines, ShowMissingsAs:=EmptyFieldSymbol, Delimiter:=",")
13        ElseIf ParserName = "sdkn104" Then
14            If IgnoreEmptyLines Then Throw "sdkn104 does not support IgnoreEmptyLines"
15            ParserRes = ParseCSVToArray(InputString, True)
16            For i = LBound(ParserRes, 1) To UBound(ParserRes, 1)
17                For j = LBound(ParserRes, 2) To UBound(ParserRes, 2)
18                    If ParserRes(i, j) = "" Then
19                        ParserRes(i, j) = EmptyFieldSymbol
20                    End If
21                Next
22            Next
23        ElseIf ParserName = "ws_garcia" Then
          Dim FSO As New Scripting.FileSystemObject
          Dim T As Scripting.TextStream
24        Set T = FSO.OpenTextFile("c:\Temp\temp.txt", ForWriting, True)
25        T.Write InputString
26        T.Close
          
27        ParserRes = Wrap_ws_garcia("c:\temp\temp.txt", ",", vbLf, IgnoreEmptyLines)
          
28            For i = LBound(ParserRes, 1) To UBound(ParserRes, 1)
29                For j = LBound(ParserRes, 2) To UBound(ParserRes, 2)
30                    If ParserRes(i, j) = "" Then
31                        ParserRes(i, j) = EmptyFieldSymbol
32                    End If
33                Next
34            Next


35        Else
36            Throw "ParserName not recognised"
37        End If

38        NR = NRows(ParserRes)
39        NC = NCols(ParserRes)
40        ReDim OneDArray(LBound(ParserRes, 1) To UBound(ParserRes, 1))

41        For i = LBound(ParserRes, 1) To UBound(ParserRes, 1)
42            OneDArray(i) = ParserRes(i, LBound(ParserRes, 1))
43            For j = LBound(ParserRes, 2) + 1 To UBound(ParserRes, 2)
44                OneDArray(i) = OneDArray(i) & NewColumnSymbol & ParserRes(i, j)
45            Next j
46        Next i

47        Out = VBA.Join(OneDArray, NewRowSymbol)
48        Out = Replace(Out, vbLf, LFSymbol)
49        Out = Replace(Out, vbCr, CRSymbol)
50        Out = Replace(Out, " ", SpaceSymbol)
51        TestNonStandardInput = Out

52        Exit Function
ErrHandler:
53        TestNonStandardInput = "#TestNonStandardInput (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function



