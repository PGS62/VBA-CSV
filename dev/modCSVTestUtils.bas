Attribute VB_Name = "modCSVTestUtils"
' VBA-CSV

' Copyright (C) 2021 - Philip Swannell (https://github.com/PGS62/VBA-CSV )
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/VBA-CSV#readme

'Module contains functions called from the worksheet "Test", mostly meta-programming to construct the code of method RunTests

Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : GenerateTestCode
' Purpose    : Metaprogramming - generate the VBA code for a given test, used on worksheet Test to genenerate a single
'              method to be called from RunTests.
' -----------------------------------------------------------------------------------------------------------------------
Function GenerateTestCode(TestNo As Long, FileName, ExpectedReturn As Variant, ConvertTypes As Variant, Delimiter As Variant, IgnoreRepeated As Boolean, DateFormat As String, _
    Comment As String, IgnoreEmptyLines As Boolean, HeaderRowNum As Long, SkipToRow As Long, SkipToCol As Long, NumRows As Long, NumCols As Long, TrueStrings As String, _
    FalseStrings As String, MissingStrings As String, Encoding As Variant, DecimalSeparator As String, ExpectedHeaderRow As Variant)

    Dim LitteralExpected
    Dim Res As String
    
    Const Indent = "    "
    Const Indent2 = "        "
    
    On Error GoTo ErrHandler
    
    Res = "Sub Test" & TestNo & "(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)"
    Res = Res & vbLf & Indent & "Dim Expected"
    Res = Res & vbLf & Indent & "Dim FileName As String"
    Res = Res & vbLf & Indent & "Dim Observed"
    Res = Res & vbLf & Indent & "Dim TestDescription As String"
    Res = Res & vbLf & Indent & "Dim TestRes As Boolean"
    Res = Res & vbLf & Indent & "Dim WhatDiffers As String"
    Res = Res & vbLf
    Res = Res & vbLf & Indent & "On Error GoTo ErrHandler"
    
    Res = Res & vbLf & _
        Indent + "TestDescription = """ & Replace(Replace(FileName, "_", " "), ".csv", "") & """"
    
    If Not IsArray(ExpectedReturn) Then
        LitteralExpected = ElementToVBALitteral(ExpectedReturn)
    Else
        LitteralExpected = ArrayToVBALitteral(ExpectedReturn, , 10000)
        If Left$(LitteralExpected, 1) = "#" Then
        End If
    End If
    
    Res = Res + vbLf + Indent + "Expected = " & LitteralExpected

    Res = Res + vbLf + Indent + "FileName = """ & FileName & """"

    If Left$(FileName, 4) = "http" Then
        Res = Res + vbLf + Indent + "TestRes = TestCSVRead(" & TestNo & ", TestDescription, Expected, FileName, Observed, WhatDiffers"
    Else
        Res = Res + vbLf + Indent + "TestRes = TestCSVRead(" & TestNo & ", TestDescription, Expected, Folder + FileName, Observed, WhatDiffers"
    End If

    If IsArray(ConvertTypes) Then
        Res = Res + ", _" + vbLf + Indent2 + "ConvertTypes := " & ArrayToVBALitteral(ConvertTypes)
    ElseIf ConvertTypes <> False And ConvertTypes <> "" Then
        Res = Res + ", _" + vbLf + Indent2 + "ConvertTypes := " & ElementToVBALitteral(ConvertTypes)
    End If

    If Delimiter <> "" Then
        Res = Res + ", _" + vbLf + Indent2 + "Delimiter := " & ElementToVBALitteral(Delimiter)
    End If
    If IgnoreRepeated = True Then
        Res = Res + ", _" + vbLf + Indent2 + "IgnoreRepeated := True"
    End If
    If DateFormat <> "" Then
        Res = Res + ", _" + vbLf + Indent2 + "DateFormat := " & ElementToVBALitteral(DateFormat)
    End If
    If Comment <> "" Then
        Res = Res + ", _" + vbLf + Indent2 + "Comment := " & ElementToVBALitteral(Comment)
    End If
    If IgnoreEmptyLines <> True Then
        Res = Res + ", _" + vbLf + Indent2 + "IgnoreEmptyLines := " & ElementToVBALitteral(IgnoreEmptyLines)
    End If
    
    If SkipToRow <> 1 And SkipToRow <> 0 Then
        Res = Res + ", _" + vbLf + Indent2 + "SkipToRow := " & CStr(SkipToRow)
    End If
    If SkipToCol <> 1 And SkipToCol <> 0 Then
        Res = Res + ", _" + vbLf + Indent2 + "SkipToCol := " & CStr(SkipToCol)
    End If
    If NumRows <> 0 Then
        Res = Res + ", _" + vbLf + Indent2 + "NumRows := " & CStr(NumRows)
    End If
    If NumCols <> 0 Then
        Res = Res + ", _" + vbLf + Indent2 + "NumCols := " & CStr(NumCols)
    End If
    If TrueStrings <> "" Then
        If InStr(TrueStrings, ",") = 0 Then
            Res = Res + ", _" + vbLf + Indent2 + "TrueStrings := " & ElementToVBALitteral(TrueStrings)
        Else
            Res = Res + ", _" + vbLf + Indent2 + "TrueStrings := " & ArrayToVBALitteral(VBA.Split(TrueStrings, ","))
        End If
    End If
    If FalseStrings <> "" Then
        If InStr(FalseStrings, ",") = 0 Then
            Res = Res + ", _" + vbLf + Indent2 + "FalseStrings := " & ElementToVBALitteral(FalseStrings)
        Else
            Res = Res + ", _" + vbLf + Indent2 + "FalseStrings := " & ArrayToVBALitteral(VBA.Split(FalseStrings, ","))
        End If
    End If
    If MissingStrings <> "" Then
        If InStr(MissingStrings, ",") = 0 Then
            Res = Res + ", _" + vbLf + Indent2 + "MissingStrings := " & ElementToVBALitteral(MissingStrings)
        Else
            Res = Res + ", _" + vbLf + Indent2 + "MissingStrings := " & ArrayToVBALitteral(VBA.Split(MissingStrings, ","))
        End If
    End If
    
    Res = Res + ", _" + vbLf + Indent2 + "ShowMissingsAs := Empty"
    If Encoding <> "" And Not IsEmpty(Encoding) Then
        Res = Res + ", _" + vbLf + Indent2 + "Encoding := " & ElementToVBALitteral(Encoding)
    End If
    If DecimalSeparator <> "." And DecimalSeparator <> "" Then
        Res = Res + ", _" + vbLf + Indent2 + "DecimalSeparator := " & ElementToVBALitteral(DecimalSeparator)
    End If
    If HeaderRowNum <> 0 Then
        Res = Res + ", _" + vbLf + Indent2 + "HeaderRowNum := " & ElementToVBALitteral(HeaderRowNum)
    End If
    
    If Not ArraysIdentical(ExpectedHeaderRow, "#Not requested!") Then
        Res = Res + ", _" + vbLf + Indent2 + "ExpectedHeaderRow := " & ArrayToVBALitteral(ExpectedHeaderRow)
    End If

    Res = Res + ")"
    
    Res = Res & vbLf & Indent + "AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures"
    
    Res = Res & vbLf & ""
    Res = Res & vbLf & "    Exit Sub"
    Res = Res & vbLf & "ErrHandler:"
    Res = Res & vbLf & "    Throw ""#Test" & TestNo & " (line "" & CStr(Erl) + ""): "" & Err.Description & ""!"""
    Res = Res & vbLf & "End Sub"

    GenerateTestCode = Transpose(Split(Res, vbLf))

    Exit Function
ErrHandler:
    GenerateTestCode = "#GenerateTestCode (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function ElementToVBALitteral(x)

    On Error GoTo ErrHandler
    If VarType(x) = vbDate Then
        If x <= 1 Then
            ElementToVBALitteral = "CDate(""" + Format(x, "hh:mm:ss") + """)"
        ElseIf x = CLng(x) Then
            ElementToVBALitteral = "CDate(""" + Format(x, "yyyy-mmm-dd") + """)"
        Else
            ElementToVBALitteral = "CDate(""" + Format(x, "yyyy-mmm-dd hh:mm:ss") + """)"
        End If

    ElseIf IsNumberOrDate(x) Then
        ElementToVBALitteral = CStr(x) + "#"
    ElseIf VarType(x) = vbString Then
        If x = vbTab Then
            ElementToVBALitteral = "vbTab"

        ElseIf x = "I'm missing!" Then 'Hack
            ElementToVBALitteral = "Empty"
        Else
            If IsWideString(CStr(x)) Then
                ElementToVBALitteral = HandleWideString(CStr(x))
            Else
                x = Replace(x, """", """""")
                x = Replace(x, vbCrLf, """ + vbCrLf + """)
                x = Replace(x, vbLf, """ + vbLf + """)
                x = Replace(x, vbCr, """ + vbCr + """)
                x = Replace(x, vbTab, """ + vbTab + """)
                ElementToVBALitteral = """" + x + """"
            End If
        End If
    ElseIf VarType(x) = vbBoolean Then
        ElementToVBALitteral = CStr(x)
    ElseIf IsEmpty(x) Then
        ElementToVBALitteral = "Empty"
    ElseIf IsError(x) Then
        ElementToVBALitteral = "CVErr(" & Mid$(CStr(x), 7) & ")"
    End If

    Exit Function
ErrHandler:
    Throw "#ElementToVBALitteral (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ArrayToVBALitteral
' Purpose    : Metaprogramming. Given an array of arbitrary data (strings, doubles, booleans, empties, errors) returns a
'              snippet of VBA code that would generate that data and assign it to a variable AssignTo. The generated code
'              assumes functions HStack and VStack are available.
' -----------------------------------------------------------------------------------------------------------------------
Function ArrayToVBALitteral(TheData As Variant, Optional AssignTo As String, Optional LengthLimit As Long = 5000)
    Dim i As Long
    Dim j As Long
    Dim NC As Long
    Dim NR As Long
    Dim Res As String

    On Error GoTo ErrHandler
    If TypeName(TheData) = "Range" Then
        TheData = TheData.value
    End If

    Force2DArray TheData, NR, NC

    If AssignTo <> "" Then
        Res = AssignTo & " = "
    End If

    Res = Res + "HStack( _" + vbLf

    For j = 1 To NC
        If NR > 1 Then
            Res = Res + "Array("
        End If
        For i = 1 To NR
            Res = Res + ElementToVBALitteral(TheData(i, j))
            'Avoid attempting to build massive string in a manner which will be slow
            If Len(Res) > LengthLimit Then Throw "Length limit (" + CStr(LengthLimit) + ") reached"
            If i < NR Then
                Res = Res + ","
            End If
        Next i
        If NR > 1 Then
            Res = Res + ")"
        End If
        If j < NC Then
            Res = Res + ", _" + vbLf
        End If
    Next j
    Res = Res + ")"

    If Len(Res) < 130 Then
        ArrayToVBALitteral = Replace(Res, " _" & vbLf, "")
    Else
        ArrayToVBALitteral = Res
    End If

    Exit Function
ErrHandler:
    ArrayToVBALitteral = "#ArrayToVBALitteral (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function HandleWideString(TheStr As String)

    Dim i As Long
    Dim Res As String

    Res = "ChrW(" + CStr(AscW(Left$(TheStr, 1))) + ")"
    For i = 2 To Len(TheStr)
        Res = Res + " + ChrW(" + CStr(AscW(Mid$(TheStr, i, 1))) + ")"
        If i Mod 10 = 1 Then
            Res = Res + " _" & vbLf
        End If
    Next i
    HandleWideString = Res

    Exit Function
ErrHandler:
    Throw "#HandleWideString (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function IsWideString(TheStr As String) As Boolean
    Dim i As Long

    On Error GoTo ErrHandler
    For i = 1 To Len(TheStr)
        If AscW(Mid$(TheStr, i, 1)) > 255 Then
            IsWideString = True
        End If
        Exit For
    Next i

    Exit Function
ErrHandler:
    Throw "#IsWideString (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : UnPack
' Purpose    : Allow saving of arrays into cells of the Test sheet
' -----------------------------------------------------------------------------------------------------------------------
Function UnPack(Str As Variant)
    UnPack = Str
    If VarType(Str) = vbString Then
        If InStr(Str, vbLf) > 0 Then
            UnPack = CSVRead(CStr(Str), "NDBE")
        End If
    End If
End Function
