Attribute VB_Name = "modCSVNotRFC4180"
Option Explicit

'Convenience function to compare parser's handling of non-RFC4180 compliant input, with those of _
the Java library FastCSV (and other Java libraries). See notation used at https://github.com/osiegmar/JavaCsvComparison

Function TestNonStandardInput(ByVal InputString As String, IgnoreEmptyLines As Boolean, ParserName As String, ExpectedResult As String)

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
8         Dim NotSupportedSymbol As String: NotSupportedSymbol = ChrW(10134)
9         Dim CorrectResultSymbol As String: CorrectResultSymbol = ChrW(9989)
10        Dim CrashesSymbol As String: CrashesSymbol = ThisWorkbook.Worksheets("notrfc4180").Range("CrashSymbol").value 'Hacky but cant find another way...

11        InputString = Replace(InputString, LFSymbol, vbLf)
12        InputString = Replace(InputString, CRSymbol, vbCr)
13        InputString = Replace(InputString, SpaceSymbol, " ")

14        If ParserName = "CSVRead" Then
15            ParserRes = CSVRead(InputString, IgnoreEmptyLines:=IgnoreEmptyLines, ShowMissingsAs:=EmptyFieldSymbol, Delimiter:=",")
16        ElseIf ParserName = "sdkn104" Then
17            If IgnoreEmptyLines Then
18                TestNonStandardInput = NotSupportedSymbol
19                Exit Function
20            End If

21            ParserRes = ParseCSVToArray(InputString, True)
22            If IsNull(ParserRes) Then
                 ' Debug.Print Err.Description
23                TestNonStandardInput = CrashesSymbol
24                Exit Function
25            End If

26            For i = LBound(ParserRes, 1) To UBound(ParserRes, 1)
27                For j = LBound(ParserRes, 2) To UBound(ParserRes, 2)
28                    If ParserRes(i, j) = "" Then
29                        ParserRes(i, j) = EmptyFieldSymbol
30                    End If
31                Next
32            Next
33        ElseIf ParserName = "ws_garcia" Then
              Dim FSO As New Scripting.FileSystemObject
              Dim T As Scripting.TextStream
34            Set T = FSO.OpenTextFile("c:\Temp\temp.txt", ForWriting, True)
35            T.Write InputString
36            T.Close
          
37            ParserRes = Wrap_ws_garcia("c:\temp\temp.txt", ",", vbLf, IgnoreEmptyLines, True)
38            If NumDimensions(ParserRes) = 0 Then
39                TestNonStandardInput = CrashesSymbol
40                Exit Function
41            End If
          
42            For i = LBound(ParserRes, 1) To UBound(ParserRes, 1)
43                For j = LBound(ParserRes, 2) To UBound(ParserRes, 2)
44                    If ParserRes(i, j) = "" Then
45                        ParserRes(i, j) = EmptyFieldSymbol
46                    End If
47                Next
48            Next


49        Else
50            Throw "ParserName not recognised"
51        End If

52        NR = NRows(ParserRes)
53        NC = NCols(ParserRes)
54        ReDim OneDArray(LBound(ParserRes, 1) To UBound(ParserRes, 1))

55        For i = LBound(ParserRes, 1) To UBound(ParserRes, 1)
56            OneDArray(i) = ParserRes(i, LBound(ParserRes, 1))
57            For j = LBound(ParserRes, 2) + 1 To UBound(ParserRes, 2)
58                OneDArray(i) = OneDArray(i) & NewColumnSymbol & ParserRes(i, j)
59            Next j
60        Next i

61        Out = VBA.Join(OneDArray, NewRowSymbol)
62        Out = Replace(Out, vbLf, LFSymbol)
63        Out = Replace(Out, vbCr, CRSymbol)
64        Out = Replace(Out, " ", SpaceSymbol)
65        If Out = ExpectedResult Then Out = CorrectResultSymbol
66        TestNonStandardInput = Out

67        Exit Function
ErrHandler:
68        TestNonStandardInput = "#TestNonStandardInput (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function



