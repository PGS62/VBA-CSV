Attribute VB_Name = "modCSVTestUtils"

' VBA-CSV

' Copyright (C) 2021 - Philip Swannell (https://github.com/PGS62/VBA-CSV )
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/VBA-CSV#readme

Option Explicit

'Module contains functions called from the worksheet "Test", mostly meta-programming to construct the code of method RunTests

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ArrayToVBALitteral
' Purpose    : Metaprogramming. Given an array of arbitrary data (strings, doubles, booleans, empties, errors) returns a
'              snippet of VBA code that would generate that data and assign it to a variable AssignTo. The generated code
'              assumes functions HStack and VStack are available.
' -----------------------------------------------------------------------------------------------------------------------
Function ArrayToVBALitteral(TheData As Variant, Optional AssignTo As String, Optional LengthLimit As Long = 5000)
          Dim NR As Long, NC As Long, i As Long, j As Long
          Dim res As String

1         On Error GoTo ErrHandler
2         If TypeName(TheData) = "Range" Then
3             TheData = TheData.value
4         End If

5         Force2DArray TheData, NR, NC

6         If AssignTo <> "" Then
7             res = AssignTo & " = "
8         End If

9         res = res + "HStack( _" + vbLf

10        For j = 1 To NC
11            If NR > 1 Then
12                res = res + "Array("
13            End If
14            For i = 1 To NR
15                res = res + ElementToVBALitteral(TheData(i, j))
                  'Avoid attempting to build massive string in a manner which will be slow
16                If Len(res) > LengthLimit Then Throw "Length limit (" + CStr(LengthLimit) + ") reached"
17                If i < NR Then
18                    res = res + ","
19                End If
20            Next i
21            If NR > 1 Then
22                res = res + ")"
23            End If
24            If j < NC Then
25                res = res + ", _" + vbLf
26            End If
27        Next j
28        res = res + ")"

29        If Len(res) < 100 Then
30            ArrayToVBALitteral = Replace(res, " _" & vbLf, "")
31        Else
32            ArrayToVBALitteral = Transpose(VBA.Split(res, vbLf))
33        End If

34        Exit Function
ErrHandler:
35        ArrayToVBALitteral = "#ArrayToVBALitteral (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function IsWideString(TheStr As String) As Boolean
          Dim i As Long

1         On Error GoTo ErrHandler
2         For i = 1 To Len(TheStr)
3             If AscW(Mid(TheStr, i, 1)) > 255 Then
4                 IsWideString = True
5             End If
6             Exit For
7         Next i

8         Exit Function
ErrHandler:
9         Throw "#IsWideString (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function HandleWideString(TheStr As String)

          Dim i As Long
          Dim res As String

1         res = "ChrW(" + CStr(AscW(Left(TheStr, 1))) + ")"
2         For i = 2 To Len(TheStr)
3             res = res + " + ChrW(" + CStr(AscW(Mid(TheStr, i, 1))) + ")"
4             If i Mod 10 = 1 Then
5                 res = res + " _" & vbLf
6             End If
7         Next i
8         HandleWideString = res

9         Exit Function
ErrHandler:
10        Throw "#HandleWideString (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function ElementToVBALitteral(x)

1         On Error GoTo ErrHandler
2         If VarType(x) = vbDate Then
3             If x <= 1 Then
4                 ElementToVBALitteral = "CDate(""" + Format(x, "hh:mm:ss") + """)"
5             ElseIf x = CLng(x) Then
6                 ElementToVBALitteral = "CDate(""" + Format(x, "yyyy-mmm-dd") + """)"
7             Else
8                 ElementToVBALitteral = "CDate(""" + Format(x, "yyyy-mmm-dd hh:mm:ss") + """)"
9             End If

10        ElseIf IsNumberOrDate(x) Then
11            ElementToVBALitteral = CStr(x) + "#"
12        ElseIf VarType(x) = vbString Then
13            If x = vbTab Then
14                ElementToVBALitteral = "vbTab"

15            ElseIf x = "I'm missing!" Then 'Hack
16                ElementToVBALitteral = "Empty"
17            Else
18                If IsWideString(CStr(x)) Then
19                    ElementToVBALitteral = HandleWideString(CStr(x))
20                Else
21                    x = Replace(x, """", """""")
22                    x = Replace(x, vbCrLf, """ + vbCrLf + """)
23                    x = Replace(x, vbLf, """ + vbLf + """)
24                    x = Replace(x, vbCr, """ + vbCr + """)
25                    x = Replace(x, vbTab, """ + vbTab + """)
26                    ElementToVBALitteral = """" + x + """"
27                End If
28            End If
29        ElseIf VarType(x) = vbBoolean Then
30            ElementToVBALitteral = CStr(x)
31        ElseIf IsEmpty(x) Then
32            ElementToVBALitteral = "Empty"
33        ElseIf IsError(x) Then
34            ElementToVBALitteral = "CVErr(" & Mid(CStr(x), 7) & ")"
35        End If

36        Exit Function
ErrHandler:
37        Throw "#ElementToVBALitteral (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function GenerateTestCode(ConvertTypes As Variant, Delimiter As String, IgnoreRepeated As Boolean, DateFormat As String, _
    Comment As String, SkipToRow As Long, SkipToCol As Long, NumRows As Long, NumCols As Long, TrueStrings As String, FalseStrings As String, MissingStrings As String, Encoding As Variant, DecimalSeparator As String)

    Dim res As String
    Const IndentBy = 4

    On Error GoTo ErrHandler
    res = "TestRes = TestCSVRead(i, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers"

    If ConvertTypes <> False Then
        res = res + ", _" + vbLf + String(IndentBy, " ") + "ConvertTypes := " & ElementToVBALitteral(ConvertTypes)
    End If

    If Delimiter <> "" Then
        res = res + ", _" + vbLf + String(IndentBy, " ") + "Delimiter := " & ElementToVBALitteral(Delimiter)
    End If
    If IgnoreRepeated = True Then
        res = res + ", _" + vbLf + String(IndentBy, " ") + "IgnoreRepeated := True"
    End If
    If DateFormat <> "" Then
        res = res + ", _" + vbLf + String(IndentBy, " ") + "DateFormat := " & ElementToVBALitteral(DateFormat)
    End If
    If Comment <> "" Then
        res = res + ", _" + vbLf + String(IndentBy, " ") + "Comment := " & ElementToVBALitteral(Comment)
    End If
    If SkipToRow <> 1 And SkipToRow <> 0 Then
        res = res + ", _" + vbLf + String(IndentBy, " ") + "SkipToRow := " & CStr(SkipToRow)
    End If
    If SkipToCol <> 1 And SkipToCol <> 0 Then
        res = res + ", _" + vbLf + String(IndentBy, " ") + "SkipToCol := " & CStr(SkipToCol)
    End If
    If NumRows <> 0 Then
        res = res + ", _" + vbLf + String(IndentBy, " ") + "NumRows := " & CStr(NumRows)
    End If
    If NumCols <> 0 Then
        res = res + ", _" + vbLf + String(IndentBy, " ") + "NumCols := " & CStr(NumCols)
    End If
    If TrueStrings <> "" Then
        If InStr(TrueStrings, ",") = 0 Then
            res = res + ", _" + vbLf + String(IndentBy, " ") + "TrueStrings := " & ElementToVBALitteral(TrueStrings)
        Else
            res = res + ", _" + vbLf + String(IndentBy, " ") + "TrueStrings := " & ArrayToVBALitteral(VBA.Split(TrueStrings, ","))
        End If
    End If
    If FalseStrings <> "" Then
        If InStr(FalseStrings, ",") = 0 Then
            res = res + ", _" + vbLf + String(IndentBy, " ") + "FalseStrings := " & ElementToVBALitteral(FalseStrings)
        Else
            res = res + ", _" + vbLf + String(IndentBy, " ") + "FalseStrings := " & ArrayToVBALitteral(VBA.Split(FalseStrings, ","))
        End If
    End If
    If MissingStrings <> "" Then
        If InStr(MissingStrings, ",") = 0 Then
            res = res + ", _" + vbLf + String(IndentBy, " ") + "MissingStrings := " & ElementToVBALitteral(MissingStrings)
        Else
            res = res + ", _" + vbLf + String(IndentBy, " ") + "MissingStrings := " & ArrayToVBALitteral(VBA.Split(MissingStrings, ","))
        End If
    End If
    
    res = res + ", _" + vbLf + String(IndentBy, " ") + "ShowMissingsAs := Empty"
    If Encoding <> "" And Not IsEmpty(Encoding) Then
        res = res + ", _" + vbLf + String(IndentBy, " ") + "Encoding := " & ElementToVBALitteral(Encoding)
    End If
    If DecimalSeparator <> "." And DecimalSeparator <> "" Then
        res = res + ", _" + vbLf + String(IndentBy, " ") + "DecimalSeparator := " & ElementToVBALitteral(DecimalSeparator)
    End If

    res = res + ")"

    GenerateTestCode = Transpose(Split(res, vbLf))

    Exit Function
ErrHandler:
    GenerateTestCode = "#GenerateTestCode (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
