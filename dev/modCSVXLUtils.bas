Attribute VB_Name = "modCSVXLUtils"

' VBA-CSV

' Copyright (C) 2021 - Philip Swannell (https://github.com/PGS62/VBA-CSV )
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/VBA-CSV#readme

Option Explicit
'Functions that, in addition to CSVRead and CSVWrite, are called from the worksheets of this workbook _
 see also mod

Function TempFolder()
    TempFolder = Environ("Temp")
End Function

Function TestFolder()
    TestFolder = Left(ThisWorkbook.path, InStrRev(ThisWorkbook.path, "\")) + "testfiles\"
End Function

'---------------------------------------------------------------------------------------------------------
' Procedure : ArrayEquals
' Purpose   : Element-wise testing for equality of two arrays - the array version of Equals. Like the =
'             operator in Excel array formulas, but capable of comparing error values, so
'             always returns an array of logicals. See also ArraysIdentical.
' Arguments
' Array1    : The first array to compare, with arbitrary values - numbers, text, errors, logicals etc.
' Array2    : The second array to compare, with arbitrary values - numbers, text, errors, logicals etc.
' CaseSensitive: Determines if comparison of strings is done in a case sensitive manner. If omitted
'             defaults to FALSE (case insensitive matching).
'---------------------------------------------------------------------------------------------------------
Function ArrayEquals(Array1 As Variant, Array2 As Variant, Optional CaseSensitive As Variant = False)
    On Error GoTo ErrHandler
    Dim i As Long
    Dim j As Long
    Dim NC1 As Long
    Dim NC2 As Long
    Dim NCMax As Long
    Dim NCMin As Long
    Dim NR1 As Long
    Dim NR2 As Long
    Dim NRMax As Long
    Dim NRMin As Long
    Dim Ret() As Variant

    If VarType(Array1) < vbArray And VarType(Array2) < vbArray And VarType(CaseSensitive) = vbBoolean Then
        ArrayEquals = Equals(Array1, Array2, CBool(CaseSensitive))
    Else

        Force2DArrayR Array1, NR1, NC1
        Force2DArrayR Array2, NR2, NC2

        If NR1 < NR2 Then
            NRMax = NR2
            NRMin = NR1
        Else
            NRMax = NR1
            NRMin = NR2
        End If
        If NC1 < NC2 Then
            NCMax = NC2
            NCMin = NC1
        Else
            NCMax = NC1
            NCMin = NC2
        End If
        Ret = Fill(CVErr(xlErrNA), NRMax, NCMax)
        For i = 1 To NRMin
            For j = 1 To NCMin
                Ret(i, j) = Equals(Array1(i, j), Array2(i, j))
            Next j
        Next i
  
        ArrayEquals = Ret

    End If
    Exit Function
ErrHandler:
    ArrayEquals = "#ArrayEquals (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : FileFromPath
' Purpose    : Split file-with-path to file name (if ReturnFileName is True) or path otherwise.
' -----------------------------------------------------------------------------------------------------------------------
Function FileFromPath(FullFileName As Variant, Optional ReturnFileName As Boolean = True) As Variant
    On Error GoTo ErrHandler
    Dim SlashPos As Long
    Dim SlashPos2 As Long
    If VarType(FullFileName) = vbString Then
        SlashPos = InStrRev(FullFileName, "\")
        SlashPos2 = InStrRev(FullFileName, "/")
        If SlashPos2 > SlashPos Then SlashPos = SlashPos2
        If SlashPos = 0 Then Throw "Neither '\' nor '/' found"

        If ReturnFileName Then
            FileFromPath = Mid$(FullFileName, SlashPos + 1)
        Else
            FileFromPath = Left$(FullFileName, SlashPos - 1)
        End If
    Else
        Throw "FullFileName must be a string"
    End If

    Exit Function
ErrHandler:
    FileFromPath = "#" & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : FileReadAll
' Purpose    : Returns the contents of a file as a string
' -----------------------------------------------------------------------------------------------------------------------
Function FileReadAll(FileName As String)
    Dim F As Scripting.File
    Dim FSO As New FileSystemObject
    Dim T As Scripting.TextStream
    On Error GoTo ErrHandler
    Set F = FSO.GetFile(FileName)
    Set T = F.OpenAsTextStream()
    FileReadAll = T.ReadAll
    T.Close

    Exit Function
ErrHandler:
    Throw "#FileReadAll (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------------------------
' Procedure : VStack
' Purpose   : Places arrays vertically on top of one another. If the arrays are of unequal width then they will be
'             padded to the right with #NA! values.
'  Notes   1) Input arrays to range can have 0, 1, or 2 dimensions.
'          2) output array has lower bound 1, whatever the lower bounds of the inputs.
'          3) input arrays of 1 dimension are treated as if they were rows, same as SAI equivalent fn.
'---------------------------------------------------------------------------------------------------------
Function VStack(ParamArray Arrays())
    Dim AllC As Long
    Dim AllR As Long
    Dim c As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim R As Long
    Dim R0 As Long
    Dim ReturnArray()
    On Error GoTo ErrHandler

    Static NA As Variant
    If IsMissing(Arrays) Then
        VStack = CreateMissing()
    Else
        If IsEmpty(NA) Then NA = CVErr(xlErrNA)

        For i = LBound(Arrays) To UBound(Arrays)
            If TypeName(Arrays(i)) = "Range" Then Arrays(i) = Arrays(i).value
            If IsMissing(Arrays(i)) Then
                R = 0: c = 0
            Else
                Select Case NumDimensions(Arrays(i))
                    Case 0
                        R = 1: c = 1
                    Case 1
                        R = 1
                        c = UBound(Arrays(i)) - LBound(Arrays(i)) + 1
                    Case 2
                        R = UBound(Arrays(i), 1) - LBound(Arrays(i), 1) + 1
                        c = UBound(Arrays(i), 2) - LBound(Arrays(i), 2) + 1
                End Select
            End If
            If c > AllC Then AllC = c
            AllR = AllR + R
        Next i

        If AllR = 0 Then
            VStack = CreateMissing
            Exit Function
        End If

        ReDim ReturnArray(1 To AllR, 1 To AllC)

        R0 = 1
        For i = LBound(Arrays) To UBound(Arrays)
            If Not IsMissing(Arrays(i)) Then
                Select Case NumDimensions(Arrays(i))
                    Case 0
                        R = 1: c = 1
                        ReturnArray(R0, 1) = Arrays(i)
                    Case 1
                        R = 1
                        c = UBound(Arrays(i)) - LBound(Arrays(i)) + 1
                        For j = 1 To c
                            ReturnArray(R0, j) = Arrays(i)(j + LBound(Arrays(i)) - 1)
                        Next j
                    Case 2
                        R = UBound(Arrays(i), 1) - LBound(Arrays(i), 1) + 1
                        c = UBound(Arrays(i), 2) - LBound(Arrays(i), 2) + 1

                        For j = 1 To R
                            For k = 1 To c
                                ReturnArray(R0 + j - 1, k) = Arrays(i)(j + LBound(Arrays(i), 1) - 1, k + LBound(Arrays(i), 2) - 1)
                            Next k
                        Next j

                End Select
                If c < AllC Then
                    For j = 1 To R
                        For k = c + 1 To AllC
                            ReturnArray(R0 + j - 1, k) = NA
                        Next k
                    Next j
                End If
                R0 = R0 + R
            End If
        Next i

        VStack = ReturnArray
    End If
    Exit Function
ErrHandler:
    VStack = "#VStack (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function FileSize(FileName As String)
    Static FSO As Scripting.FileSystemObject

    On Error GoTo ErrHandler
    If FSO Is Nothing Then Set FSO = New FileSystemObject

    FileSize = FSO.GetFile(FileName).Size

    Exit Function
ErrHandler:
    Throw "#FileSize (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Fill
' Purpose    : Creates an array filled with the value x
' -----------------------------------------------------------------------------------------------------------------------
Function Fill(ByVal x As Variant, ByVal NumRows As Long, ByVal NumCols As Long)

    On Error GoTo ErrHandler

    Dim i As Long
    Dim j As Long
    Dim Result() As Variant

    ReDim Result(1 To NumRows, 1 To NumCols)

    For i = 1 To NumRows
        For j = 1 To NumCols
            Result(i, j) = x
        Next j
    Next i

    Fill = Result

    Exit Function
ErrHandler:
    Fill = "#Fill: " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------------------------
' Procedure : SplitString
' Purpose   : Breaks up TheString into sub-strings with breaks at the positions at which the Delimiter
'             character appears, and returns the sub-strings as a two-dimensional, 1-based, 1 column array.
' Arguments
' TheString : The string to be split.
' Delimiter : The delimiter string, can be multiple characters. The search for the delimiter
'             is case insensitive.
'---------------------------------------------------------------------------------------------------------
Function SplitString(TheString As String, Optional Delimiter As String = ",")

    Dim i As Long
    Dim LB As Long
    Dim N As Long
    Dim OneDArray
    Dim Res()
    Dim UB As Long
    
    On Error GoTo ErrHandler
    If Len(TheString) = 0 Then
        ReDim Res(1 To 1, 1 To 1)
        Res(1, 1) = ""
        SplitString = Res
        Exit Function
    End If
    
    OneDArray = VBA.Split(TheString, Delimiter, -1, vbTextCompare)
    LB = LBound(OneDArray): UB = UBound(OneDArray)
    N = UB - LB + 1
    ReDim Res(1 To N, 1 To 1)
    For i = 1 To N
        Res(i, 1) = OneDArray(i - 1)
    Next
    SplitString = Res
    Exit Function
ErrHandler:
    SplitString = "#SplitString (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : AllCombinations
' Purpose    : Iterate over all combinations of the elements of the (up to) 4 input arrays, producing an output vector
'              of those elements concatenated with given delimiter.
' -----------------------------------------------------------------------------------------------------------------------
Function AllCombinations(Arg1, Optional Arg2, Optional Arg3, _
    Optional Arg4, Optional Delimiter As String)
    Dim k As Long
    Dim Part1 As Variant
    Dim Part2 As Variant
    Dim Part3 As Variant
    Dim Part4 As Variant
    Dim Res() As String

    On Error GoTo ErrHandler
    If IsMissing(Arg2) Then Arg2 = ""
    If IsMissing(Arg3) Then Arg3 = ""
    If IsMissing(Arg4) Then Arg4 = ""

    Force2DArrayR Arg1
    Force2DArrayR Arg2
    Force2DArrayR Arg3
    Force2DArrayR Arg4

    ReDim Res(1 To NRows(Arg1) * NCols(Arg1) * NRows(Arg2) * NCols(Arg2) * NRows(Arg3) * NCols(Arg4) * NRows(Arg4) * NCols(Arg4), 1 To 1)
    For Each Part1 In Arg1
        Part1 = CStr(Part1)
        For Each Part2 In Arg2
            Part2 = CStr(Part2)
            For Each Part3 In Arg3
                Part3 = CStr(Part3)
                For Each Part4 In Arg4
                    Part4 = CStr(Part4)
                    k = k + 1
                    Res(k, 1) = Part1 & IIf(Len(Part2) > 0, Delimiter, "") & Part2 & _
                        IIf(Len(Part3) > 0, Delimiter, "") & Part3 & _
                        IIf(Len(Part4) > 0, Delimiter, "") & Part4
                Next
            Next
        Next
    Next
    AllCombinations = Res

    Exit Function
ErrHandler:
    AllCombinations = "#AllCombinations (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : MakeGoodStringsBad
' Purpose    : For a vector of input strings, returns a longer vector of those strings "made bad" by injecting the
'              character "x" at all possible positions.
'              Helpful for testing, say, parsing of nearly-but-not-quite-valid ISO8601 strings.
' Parameters :
'  GoodStrings: Array (1 col) of strings.
' -----------------------------------------------------------------------------------------------------------------------
Function MakeGoodStringsBad(GoodStrings, Optional InsertThis As String = "x")

    Dim Res1D() As String

    On Error GoTo ErrHandler
    Force2DArrayR GoodStrings

    Dim i As Long
    Dim j As Long
    Dim k As Long

    ReDim Res1D(1 To 1)
    For i = 1 To NRows(GoodStrings)
        For j = 0 To Len(GoodStrings(i, 1)) + 1
            k = k + 1
            If k > UBound(Res1D) Then
                ReDim Preserve Res1D(1 To k)
            End If
            Res1D(k) = InsertInString(InsertThis, GoodStrings(i, 1), j)
        Next j
    Next i

    MakeGoodStringsBad = Transpose(Res1D)

    Exit Function
ErrHandler:
    MakeGoodStringsBad = "#MakeGoodStringsBad (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : InsertInString
' Purpose    : Sub of MakeGoodStringsBad
' -----------------------------------------------------------------------------------------------------------------------
Private Function InsertInString(InsertThis As String, ByVal InToThis As String, ByVal AtPoint As Long)

    On Error GoTo ErrHandler

    If AtPoint + Len(InsertThis) > Len(InToThis) Then
        InToThis = InToThis + String(AtPoint + Len(InsertThis) - Len(InToThis), " ")
    ElseIf AtPoint <= 0 Then
        InToThis = String(1 - AtPoint, " ") + InToThis
        AtPoint = 1
    End If

    Mid(InToThis, AtPoint, Len(InsertThis)) = InsertThis
    InsertInString = InToThis
    Exit Function
ErrHandler:
    Throw "#InsertInString (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------------------------
' Procedure : IsRegMatch
' Purpose   : Implements Regular Expressions exposed by "Microsoft VBScript Regular Expressions 5.5".
'             The function returns TRUE if StringToSearch matches RegularExpression, FALSE
'             if it does not match, or an error string if RegularExpression contains a
'             syntax error.
' Arguments
' RegularExpression: The regular expression. Must be a string. Example cat|dog to match on either the string
'             cat or the string dog.
' StringToSearch: The string to match. May be an array in which case the return from the function is an
'             array of the same dimensions.
' CaseSensitive: TRUE for case-sensitive matching, FALSE for case-insensitive matching. This argument is
'             optional, defaulting to FALSE for case-insensitive matching.
'
' Notes     : Syntax cheat sheet:
'             Character classes
'             .                 any character except newline
'             \w \d \s          word, digit, whitespace
'             \W \D \S          not word, not digit, not whitespace
'             [abc]             any of a, b, or c
'             [^abc]            not a, b, or c
'             [a-g]             character between a & g
'
'             Anchors
'             ^abc$              start / end of the string
'             \b                 word boundary
'
'             Escaped characters
'             \. \* \\          escaped special characters
'             \t \n \r          tab, linefeed, carriage return
'
'             Groups and Look-arounds
'             (abc)             capture group
'             \1                backreference to group #1
'             (?:abc)           non-capturing group
'             (?=abc)           positive lookahead
'             (?!abc)           negative lookahead
'
'             Quantifiers and Alternation
'             a* a+ a?          0 or more, 1 or more, 0 or 1
'             a{5} a{2,}        exactly five, two or more
'             a{1,3}            between one & three
'             a+? a{2,}?        match as few as possible
'             ab|cd             match ab or cd
'
'             Further reading:
'             http://www.regular-expressions.info/
'             https://en.wikipedia.org/wiki/Regular_expression
'---------------------------------------------------------------------------------------------------------
Function IsRegMatch(RegularExpression As String, ByVal StringToSearch As Variant, Optional CaseSensitive As Boolean = False)
    Dim i As Long
    Dim j As Long
    Dim Result() As Variant
    Dim rx As VBScript_RegExp_55.RegExp

    On Error GoTo ErrHandler

    If Not RegExSyntaxValid(RegularExpression) Then
        IsRegMatch = "#Invalid syntax for RegularExpression!"
        Exit Function
    End If
    Set rx = New RegExp
    With rx
        .IgnoreCase = Not CaseSensitive
        .Pattern = RegularExpression
        .Global = False        'Find first match only
    End With

    If VarType(StringToSearch) = vbString Then
        IsRegMatch = rx.Test(StringToSearch)

        GoTo EarlyExit
    ElseIf VarType(StringToSearch) < vbArray Then
        IsRegMatch = "#StringToSearch must be a string!"
        GoTo EarlyExit
    End If
    If TypeName(StringToSearch) = "Range" Then StringToSearch = StringToSearch.Value2

    Select Case NumDimensions(StringToSearch)
        Case 2
            ReDim Result(LBound(StringToSearch, 1) To UBound(StringToSearch, 1), LBound(StringToSearch, 2) To UBound(StringToSearch, 2))
            For i = LBound(StringToSearch, 1) To UBound(StringToSearch, 1)
                For j = LBound(StringToSearch, 2) To UBound(StringToSearch, 2)
                    If VarType(StringToSearch(i, j)) = vbString Then
                        Result(i, j) = rx.Test(StringToSearch(i, j))
                    Else
                        Result(i, j) = "#StringToSearch must be a string!"
                    End If
                Next j
            Next i
        Case 1
            ReDim Result(LBound(StringToSearch, 1) To UBound(StringToSearch, 1))
            For i = LBound(StringToSearch, 1) To UBound(StringToSearch, 1)
                If VarType(StringToSearch(i)) = vbString Then
                    Result(i) = rx.Test(StringToSearch(i))
                Else
                    Result(i) = "#StringToSearch must be a string!"
                End If
            Next i
        Case Else
            Throw "StringToSearch must be String or array with 1 or 2 dimensions"
    End Select

    IsRegMatch = Result
EarlyExit:
    Set rx = Nothing

    Exit Function
ErrHandler:
    IsRegMatch = "#IsRegMatch (line " & CStr(Erl) + "): " & Err.Description & "!"
    Set rx = Nothing
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RegExSyntaxValid
' Purpose    : Tests syntax of a regular expression.
' -----------------------------------------------------------------------------------------------------------------------
Function RegExSyntaxValid(RegularExpression As String) As Boolean
    Dim Res As Boolean
    Dim rx As VBScript_RegExp_55.RegExp
    On Error GoTo ErrHandler
    Set rx = New RegExp
    With rx
        .IgnoreCase = False
        .Pattern = RegularExpression
        .Global = False        'Find first match only
    End With
    Res = rx.Test("Foo")
    RegExSyntaxValid = True
    Exit Function
ErrHandler:
    RegExSyntaxValid = False
End Function

'---------------------------------------------------------------------------------------------------------
' Procedure : RegExReplace
' Purpose   : Uses regular expressions to make replacement in a set of input strings.
'
'             The function replaces every instance of the regular expression match with the
'             replacement.
' Arguments
' InputString: Input string to be transformed. Can be an array. Non-string elements will be left
'             unchanged.
' RegularExpression: A standard regular expression string.
' Replacement: A replacement template for each match of the regular expression in the input string.
' CaseSensitive: Whether matching should be case-sensitive (TRUE) or not (FALSE).
'
' Notes     : Details of regular expressions are given under sIsRegMatch. The replacement string can be
'             an explicit string, and it can also contain special escape sequences that are
'             replaced by the characters they represent. The options available are:
'
'             Characters Replacement
'             $n        n-th backreference. That is, a copy of the n-th matched group
'             specified with parentheses in the regular expression. n must be an integer
'             value designating a valid backreference, greater than zero, and of two digits
'             at most.
'             $&       A copy of the entire match
'             $`        The prefix, that is, the part of the target sequence that precedes
'             the match.
'             $´        The suffix, that is, the part of the target sequence that follows
'             the match.
'             $$        A single $ character.
'---------------------------------------------------------------------------------------------------------
Function RegExReplace(InputString As Variant, RegularExpression As String, Replacement As String, Optional CaseSensitive As Boolean)
    Dim i As Long
    Dim j As Long
    Dim Result() As String
    Dim rx As VBScript_RegExp_55.RegExp
    On Error GoTo ErrHandler

    If Not RegExSyntaxValid(RegularExpression) Then
        RegExReplace = "#Invalid syntax for RegularExpression!"
        Exit Function
    End If

    Set rx = New RegExp

    With rx
        .IgnoreCase = Not (CaseSensitive)
        .Pattern = RegularExpression
        .Global = True
    End With

    If VarType(InputString) = vbString Then
        RegExReplace = rx.Replace(InputString, Replacement)
        GoTo Cleanup
    ElseIf VarType(InputString) < vbArray Then
        RegExReplace = InputString
        GoTo Cleanup
    End If
    If TypeName(InputString) = "Range" Then InputString = InputString.Value2

    Select Case NumDimensions(InputString)
        Case 2
            ReDim Result(LBound(InputString, 1) To UBound(InputString, 1), LBound(InputString, 2) To UBound(InputString, 2))
            For i = LBound(InputString, 1) To UBound(InputString, 1)
                For j = LBound(InputString, 2) To UBound(InputString, 2)
                    If VarType(InputString(i, j)) = vbString Then
                        Result(i, j) = rx.Replace(InputString(i, j), Replacement)
                    Else
                        Result(i, j) = InputString(i, j)
                    End If
                Next j
            Next i
        Case 1
            ReDim Result(LBound(InputString, 1) To UBound(InputString, 1))
            For i = LBound(InputString, 1) To UBound(InputString, 1)
                If VarType(InputString(i)) = vbString Then
                    Result(i) = rx.Replace(InputString(i), Replacement)
                Else
                    Result(i) = InputString(i)
                End If
            Next i
        Case Else
            Throw "InputString must be a String or an array with 1 or 2 dimensions"
    End Select
    RegExReplace = Result

Cleanup:
    Set rx = Nothing
    Exit Function
ErrHandler:
    RegExReplace = "#RegExReplace (line " & CStr(Erl) + "): " & Err.Description & "!"
    Set rx = Nothing
End Function

