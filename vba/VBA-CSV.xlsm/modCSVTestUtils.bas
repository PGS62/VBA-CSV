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
Function GenerateTestCode(TestNo As Long, FileName As String, ExpectedReturn As Variant, ConvertTypes As Variant, _
    Delimiter As Variant, IgnoreRepeated As Boolean, DateFormat As String, Comment As String, _
    IgnoreEmptyLines As Boolean, HeaderRowNum As Long, SkipToRow As Long, SkipToCol As Variant, NumRows As Long, _
    NumCols As Variant, TrueStrings As String, FalseStrings As String, MissingStrings As String, Encoding As Variant, _
    DecimalSeparator As String, ExpectedHeaderRow As Variant) As Variant

    Dim LitteralExpected As String
    Dim Res As String
    
    Const Indent As String = "    "
    Const Indent2 As String = "        "
    
    On Error GoTo ErrHandler
    
    Res = "Private Sub Test" & TestNo & "(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)"
    Res = Res & vbLf & Indent & "Dim Expected as Variant"
    Res = Res & vbLf & Indent & "Dim FileName As String"
    Res = Res & vbLf & Indent & "Dim Observed As Variant"
    Res = Res & vbLf & Indent & "Dim TestDescription As String"
    Res = Res & vbLf & Indent & "Dim TestRes As Boolean"
    Res = Res & vbLf & Indent & "Dim WhatDiffers As String"
    Res = Res & vbLf
    Res = Res & vbLf & Indent & "On Error GoTo ErrHandler"
    
    Res = Res & vbLf & _
        Indent & "TestDescription = """ & Replace(Replace(Replace(Replace(Replace(Replace(FileName, vbCrLf, "<CRLF>"), """", "DQ"), vbLf, "<LF>"), vbCr, "<CR>"), "_", " "), ".csv", vbNullString) & """"
    
    If Not IsArray(ExpectedReturn) Then
        LitteralExpected = ElementToVBALiteral(ExpectedReturn)
    Else
        LitteralExpected = ArrayToVBALiteral(ExpectedReturn, , 10000)
        If Left$(LitteralExpected, 1) = "#" Then
            Dim TestSizeOnly As Boolean
            TestSizeOnly = True
        End If
    End If
    
    If TestSizeOnly Then
        Res = Res & vbLf & Indent & "Expected = Empty"
    Else
        Res = Res & vbLf & Indent & "Expected = " & LitteralExpected
    End If

    Res = Res & vbLf & Indent & "FileName =  " & ElementToVBALiteral(FileName)

    If Left$(FileName, 4) = "http" Or InStr(FileName, ",") > 0 Then
        Res = Res & vbLf & Indent & "TestRes = TestCSVRead(" & TestNo & ", TestDescription, Expected, FileName, Observed, WhatDiffers"
    Else
        Res = Res & vbLf & Indent & "TestRes = TestCSVRead(" & TestNo & ", TestDescription, Expected, Folder & FileName, Observed, WhatDiffers"
    End If
    If TestSizeOnly Then
        Res = Res & ", _" & vbLf & Indent2 & "NumRowsExpected := " & CStr(NRows(ExpectedReturn))
        Res = Res & ", _" & vbLf & Indent2 & "NumColsExpected := " & CStr(NCols(ExpectedReturn))
    End If

    If IsArray(ConvertTypes) Then
        Res = Res & ", _" & vbLf & Indent2 & "ConvertTypes := " & ArrayToVBALiteral(ConvertTypes)
    ElseIf ConvertTypes <> False And ConvertTypes <> vbNullString Then
        Res = Res & ", _" & vbLf & Indent2 & "ConvertTypes := " & ElementToVBALiteral(ConvertTypes)
    End If

    If Delimiter <> vbNullString Then
        Res = Res & ", _" & vbLf & Indent2 & "Delimiter := " & ElementToVBALiteral(Delimiter)
    End If
    If IgnoreRepeated = True Then
        Res = Res & ", _" & vbLf & Indent2 & "IgnoreRepeated := True"
    End If
    If DateFormat <> vbNullString Then
        Res = Res & ", _" & vbLf & Indent2 & "DateFormat := " & ElementToVBALiteral(DateFormat)
    End If
    If Comment <> vbNullString Then
        Res = Res & ", _" & vbLf & Indent2 & "Comment := " & ElementToVBALiteral(Comment)
    End If
    '    If IgnoreEmptyLines <> False Then
    Res = Res & ", _" & vbLf & Indent2 & "IgnoreEmptyLines := " & ElementToVBALiteral(IgnoreEmptyLines)
    '    End If
    If SkipToRow <> (HeaderRowNum & 1) And (SkipToRow <> 0) Then
        Res = Res & ", _" & vbLf & Indent2 & "SkipToRow := " & CStr(SkipToRow)
    End If
    
    If IsNumber(SkipToCol) Then
        If SkipToCol <> 1 And SkipToCol <> 0 Then
            Res = Res & ", _" & vbLf & Indent2 & "SkipToCol := " & CStr(SkipToCol)
        End If
    Else
        Res = Res & ", _" & vbLf & Indent2 & "SkipToCol := " & ElementToVBALiteral(SkipToCol)
    End If
    
    If NumRows <> 0 Then
        Res = Res & ", _" & vbLf & Indent2 & "NumRows := " & CStr(NumRows)
    End If
    
    If IsNumber(NumCols) Then
        If NumCols <> 0 Then
            Res = Res & ", _" & vbLf & Indent2 & "NumCols := " & CStr(NumCols)
        End If
    Else
        Res = Res & ", _" & vbLf & Indent2 & "NumCols := " & ElementToVBALiteral(NumCols)
    End If
    
    
    If TrueStrings <> vbNullString Then
        If InStr(TrueStrings, ",") = 0 Then
            Res = Res & ", _" & vbLf & Indent2 & "TrueStrings := " & ElementToVBALiteral(TrueStrings)
        Else
            Res = Res & ", _" & vbLf & Indent2 & "TrueStrings := " & ArrayToVBALiteral(VBA.Split(TrueStrings, ","))
        End If
    End If
    If FalseStrings <> vbNullString Then
        If InStr(FalseStrings, ",") = 0 Then
            Res = Res & ", _" & vbLf & Indent2 & "FalseStrings := " & ElementToVBALiteral(FalseStrings)
        Else
            Res = Res & ", _" & vbLf & Indent2 & "FalseStrings := " & ArrayToVBALiteral(VBA.Split(FalseStrings, ","))
        End If
    End If
    If MissingStrings <> vbNullString Then
        If InStr(MissingStrings, ",") = 0 Then
            Res = Res & ", _" & vbLf & Indent2 & "MissingStrings := " & ElementToVBALiteral(MissingStrings)
        Else
            Res = Res & ", _" & vbLf & Indent2 & "MissingStrings := " & ArrayToVBALiteral(VBA.Split(MissingStrings, ","))
        End If
    End If
    
    Res = Res & ", _" & vbLf & Indent2 & "ShowMissingsAs := Empty"
    If Encoding <> vbNullString And Not IsEmpty(Encoding) Then
        Res = Res & ", _" & vbLf & Indent2 & "Encoding := " & ElementToVBALiteral(Encoding)
    End If
    If DecimalSeparator <> "." And DecimalSeparator <> vbNullString Then
        Res = Res & ", _" & vbLf & Indent2 & "DecimalSeparator := " & ElementToVBALiteral(DecimalSeparator)
    End If
    If HeaderRowNum <> 0 Then
        Res = Res & ", _" & vbLf & Indent2 & "HeaderRowNum := " & ElementToVBALiteral(HeaderRowNum)
    End If
    
    If Not ArraysIdentical(ExpectedHeaderRow, "#Not requested!") Then
        Res = Res & ", _" & vbLf & Indent2 & "ExpectedHeaderRow := " & ArrayToVBALiteral(ExpectedHeaderRow)
    End If

    Res = Res & ")"
    
    Res = Res & vbLf & Indent & "AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures"
    
    Res = Res & vbLf
    Res = Res & vbLf & "    Exit Sub"
    Res = Res & vbLf & "ErrHandler:"
    Res = Res & vbLf & "    ReThrow ""Test" & TestNo & """, Err"
    Res = Res & vbLf & "End Sub"

    GenerateTestCode = Transpose(Split(Res, vbLf))

    Exit Function
ErrHandler:
    GenerateTestCode = ReThrow("GenerateTestCode", Err, True)
End Function

Function ElementToVBALiteral(x As Variant) As String

    On Error GoTo ErrHandler
    If VarType(x) = vbDate Then
        If x <= 1 Then
            ElementToVBALiteral = "CDate(""" & Format$(x, "hh:mm:ss") & """)"
        ElseIf x = CLng(x) Then
            ElementToVBALiteral = "CDate(""" & Format$(x, "yyyy-mmm-dd") & """)"
        Else
            ElementToVBALiteral = "CDate(""" & Format$(x, "yyyy-mmm-dd hh:mm:ss") & """)"
        End If

    ElseIf IsNumberOrDate(x) Then
        ElementToVBALiteral = CStr(x) & "#"
    ElseIf VarType(x) = vbString Then
        If x = vbTab Then
            ElementToVBALiteral = "vbTab"
        ElseIf x = vbNullString Then
            ElementToVBALiteral = "vbNullString"
        ElseIf x = "I'm missing!" Then 'Hack!!!!!!!!!!
            ElementToVBALiteral = "Empty"
        Else
            If IsWideString(CStr(x)) Then
                ElementToVBALiteral = HandleWideString(CStr(x))
            Else
                x = Replace(x, """", """""")
                x = Replace(x, vbCrLf, """ & vbCrLf & """)
                x = Replace(x, vbLf, """ & vbLf & """)
                x = Replace(x, vbCr, """ & vbCr & """)
                x = Replace(x, vbTab, """ & vbTab & """)
                ElementToVBALiteral = """" & x & """"
            End If
        End If
    ElseIf VarType(x) = vbBoolean Then
        ElementToVBALiteral = CStr(x)
    ElseIf IsEmpty(x) Then
        ElementToVBALiteral = "Empty"
    ElseIf IsError(x) Then
        ElementToVBALiteral = "CVErr(" & Mid$(CStr(x), 7) & ")"
    End If

    Exit Function
ErrHandler:
    ReThrow "ElementToVBALiteral", Err
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ArrayToVBALiteral
' Purpose    : Metaprogramming. Given an array of arbitrary data (strings, doubles, booleans, empties, errors) returns a
'              snippet of VBA code that would generate that data and assign it to a variable AssignTo. The generated code
'              assumes functions HStack and VStack are available.
' -----------------------------------------------------------------------------------------------------------------------
Function ArrayToVBALiteral(TheData As Variant, Optional AssignTo As String, Optional LengthLimit As Long = 5000)
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

    If AssignTo <> vbNullString Then
        Res = AssignTo & " = "
    End If

    Res = Res & "HStack( _" & vbLf

    For j = 1 To NC
        If NR > 1 Then
            Res = Res & "Array("
        End If
        For i = 1 To NR
            Res = Res + ElementToVBALiteral(TheData(i, j))
            'Avoid attempting to build massive string in a manner which will be slow
            If Len(Res) > LengthLimit Then Throw "Length limit (" & CStr(LengthLimit) & ") reached"
            If i < NR Then
                Res = Res & ","
            End If
        Next i
        If NR > 1 Then
            Res = Res & ")"
        End If
        If j < NC Then
            Res = Res & ", _" & vbLf
        End If
    Next j
    Res = Res & ")"

    If Len(Res) < 130 Then
        ArrayToVBALiteral = Replace(Res, " _" & vbLf, vbNullString)
    Else
        ArrayToVBALiteral = Res
    End If

    Exit Function
ErrHandler:
    ArrayToVBALiteral = ReThrow("ArrayToVBALiteral", Err, True)
End Function

Private Function HandleWideString(TheStr As String) As String

    Dim i As Long
    Dim Res As String

    Res = "ChrW(" & CStr(AscW(Left$(TheStr, 1))) & ")"
    For i = 2 To Len(TheStr)
        Res = Res & " & ChrW(" & CStr(AscW(Mid$(TheStr, i, 1))) & ")"
        If i Mod 10 = 1 Then
            Res = Res & " _" & vbLf
        End If
    Next i
    HandleWideString = Res

    Exit Function
ErrHandler:
    ReThrow "HandleWideString", Err
End Function

Private Function IsWideString(TheStr As String) As Boolean
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
    ReThrow "IsWideString", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : UnPack
' Purpose    : Allow saving of arrays into cells of the Test sheet
' -----------------------------------------------------------------------------------------------------------------------
Public Function UnPack(Str As Variant)
    UnPack = Str
    If VarType(Str) = vbString Then
        If InStr(Str, vbLf) > 0 Then
            UnPack = CSVRead(CStr(Str), "NDBE")
        End If
    End If
End Function
