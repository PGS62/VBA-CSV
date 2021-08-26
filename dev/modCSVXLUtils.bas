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

'---------------------------------------------------------------------------------------------------------
' Procedure : ArrayEquals
' Purpose   : Element-wise testing for equality of two arrays - the array version of sEquals. Like the =
'             operator in Excel array formulas, but capable of comparing error values, so
'             always returns an array of logicals. See also sArraysIdentical.
' Arguments
' Array1    : The first array to compare, with arbitrary values - numbers, text, errors, logicals etc.
' Array2    : The second array to compare, with arbitrary values - numbers, text, errors, logicals etc.
' CaseSensitive: Determines if comparison of strings is done in a case sensitive manner. If omitted
'             defaults to FALSE (case insensitive matching).
'---------------------------------------------------------------------------------------------------------
Function ArrayEquals(Array1 As Variant, Array2 As Variant, Optional CaseSensitive As Variant = False)
    On Error GoTo ErrHandler
    Dim NR1 As Long
    Dim NC1 As Long
    Dim NR2 As Long
    Dim NC2 As Long
    Dim Ret() As Variant
    Dim NRMax As Long
    Dim NRMin As Long
    Dim NCMax As Long
    Dim NCMin As Long
    Dim i As Long
    Dim j As Long

    If VarType(Array1) < vbArray And VarType(Array2) < vbArray And VarType(CaseSensitive) = vbBoolean Then
        ArrayEquals = sEquals(Array1, Array2, CBool(CaseSensitive))
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
        Ret = sFill(CVErr(xlErrNA), NRMax, NCMax)
        For i = 1 To NRMin
            For j = 1 To NCMin
                Ret(i, j) = sEquals(Array1(i, j), Array2(i, j))
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
1         On Error GoTo ErrHandler
          Dim SlashPos As Long
          Dim SlashPos2 As Long
2         If VarType(FullFileName) = vbString Then
3             SlashPos = InStrRev(FullFileName, "\")
4             SlashPos2 = InStrRev(FullFileName, "/")
5             If SlashPos2 > SlashPos Then SlashPos = SlashPos2
6             If SlashPos = 0 Then Throw "Neither '\' nor '/' found"

7             If ReturnFileName Then
8                 FileFromPath = Mid$(FullFileName, SlashPos + 1)
9             Else
10                FileFromPath = Left$(FullFileName, SlashPos - 1)
11            End If
12        Else
13            Throw "FullFileName must be a string"
14        End If

15        Exit Function
ErrHandler:
16        FileFromPath = "#" & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : FileReadAll
' Author     : Philip Swannell
' Date       : 26-Aug-2021
' Purpose    : Returns the contents of a file as a string
' -----------------------------------------------------------------------------------------------------------------------
Function FileReadAll(FileName As String)
    Dim FSO As New FileSystemObject, F As Scripting.File, T As Scripting.TextStream
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
    Dim C As Long
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
                R = 0: C = 0
            Else
                Select Case NumDimensions(Arrays(i))
                    Case 0
                        R = 1: C = 1
                    Case 1
                        R = 1
                        C = UBound(Arrays(i)) - LBound(Arrays(i)) + 1
                    Case 2
                        R = UBound(Arrays(i), 1) - LBound(Arrays(i), 1) + 1
                        C = UBound(Arrays(i), 2) - LBound(Arrays(i), 2) + 1
                End Select
            End If
            If C > AllC Then AllC = C
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
                        R = 1: C = 1
                        ReturnArray(R0, 1) = Arrays(i)
                    Case 1
                        R = 1
                        C = UBound(Arrays(i)) - LBound(Arrays(i)) + 1
                        For j = 1 To C
                            ReturnArray(R0, j) = Arrays(i)(j + LBound(Arrays(i)) - 1)
                        Next j
                    Case 2
                        R = UBound(Arrays(i), 1) - LBound(Arrays(i), 1) + 1
                        C = UBound(Arrays(i), 2) - LBound(Arrays(i), 2) + 1

                        For j = 1 To R
                            For k = 1 To C
                                ReturnArray(R0 + j - 1, k) = Arrays(i)(j + LBound(Arrays(i), 1) - 1, k + LBound(Arrays(i), 2) - 1)
                            Next k
                        Next j

                End Select
                If C < AllC Then
                    For j = 1 To R
                        For k = C + 1 To AllC
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
' Procedure  : sFill
' Purpose    : Creates an array filled with the value x
' -----------------------------------------------------------------------------------------------------------------------
Function sFill(ByVal x As Variant, ByVal NumRows As Long, ByVal NumCols As Long)

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

    sFill = Result

    Exit Function
ErrHandler:
    sFill = "#sFill: " & Err.Description & "!"
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
          Dim res()
          Dim UB As Long
          
1         On Error GoTo ErrHandler
2         If Len(TheString) = 0 Then
3             ReDim res(1 To 1, 1 To 1)
4             res(1, 1) = ""
5             SplitString = res
6             Exit Function
7         End If
          
8         OneDArray = VBA.Split(TheString, Delimiter, -1, vbTextCompare)
9         LB = LBound(OneDArray): UB = UBound(OneDArray)
10        N = UB - LB + 1
11        ReDim res(1 To N, 1 To 1)
12        For i = 1 To N
13            res(i, 1) = OneDArray(i - 1)
14        Next
15        SplitString = res
16        Exit Function
ErrHandler:
17        SplitString = "#SplitString (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : AllCombinations
' Author     : Philip Swannell
' Date       : 26-Aug-2021
' Purpose    : Iterate over all combinations of the elements of the (up to) 4 input arrays, producing an output vector
'              of those elements concatenated with given delimiter.
' -----------------------------------------------------------------------------------------------------------------------
Function AllCombinations(Arg1, Optional Arg2, Optional Arg3, _
          Optional Arg4, Optional Delimiter As String)
          Dim res() As String
          Dim Part1 As Variant
          Dim Part2 As Variant
          Dim Part3 As Variant
          Dim Part4 As Variant
          Dim k As Long

1         On Error GoTo ErrHandler
2         If IsMissing(Arg2) Then Arg2 = ""
3         If IsMissing(Arg3) Then Arg3 = ""
4         If IsMissing(Arg4) Then Arg4 = ""

5         Force2DArrayR Arg1
6         Force2DArrayR Arg2
7         Force2DArrayR Arg3
8         Force2DArrayR Arg4

9         ReDim res(1 To sNRows(Arg1) * sNCols(Arg1) * sNRows(Arg2) * sNCols(Arg2) * sNRows(Arg3) * sNCols(Arg4) * sNRows(Arg4) * sNCols(Arg4), 1 To 1)
10        For Each Part1 In Arg1
11            Part1 = CStr(Part1)
12            For Each Part2 In Arg2
13                Part2 = CStr(Part2)
14                For Each Part3 In Arg3
15                    Part3 = CStr(Part3)
16                    For Each Part4 In Arg4
17                        Part4 = CStr(Part4)
18                        k = k + 1
19                        res(k, 1) = Part1 & IIf(Len(Part2) > 0, Delimiter, "") & Part2 & _
                              IIf(Len(Part3) > 0, Delimiter, "") & Part3 & _
                              IIf(Len(Part4) > 0, Delimiter, "") & Part4
20                    Next
21                Next
22            Next
23        Next
24        AllCombinations = res

25        Exit Function
ErrHandler:
26        AllCombinations = "#AllCombinations (line " & CStr(Erl) + "): " & Err.Description & "!"
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

1         On Error GoTo ErrHandler
2         Force2DArrayR GoodStrings

          Dim i As Long
          Dim j As Long
          Dim k As Long

3         ReDim Res1D(1 To 1)
4         For i = 1 To sNRows(GoodStrings)
5             For j = 0 To Len(GoodStrings(i, 1)) + 1
6                 k = k + 1
7                 If k > UBound(Res1D) Then
8                     ReDim Preserve Res1D(1 To k)
9                 End If
10                Res1D(k) = InsertInString(InsertThis, GoodStrings(i, 1), j)
11            Next j
12        Next i

13        MakeGoodStringsBad = Transpose(Res1D)

14        Exit Function
ErrHandler:
15        MakeGoodStringsBad = "#MakeGoodStringsBad (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : InsertInString
' Purpose    : Sub of MakeGoodStringsBad
' -----------------------------------------------------------------------------------------------------------------------
Private Function InsertInString(InsertThis As String, ByVal InToThis As String, ByVal AtPoint As Long)

1         On Error GoTo ErrHandler

2         If AtPoint + Len(InsertThis) > Len(InToThis) Then
3             InToThis = InToThis + String(AtPoint + Len(InsertThis) - Len(InToThis), " ")
4         ElseIf AtPoint <= 0 Then
5             InToThis = String(1 - AtPoint, " ") + InToThis
              AtPoint = 1
6         End If

7         Mid(InToThis, AtPoint, Len(InsertThis)) = InsertThis
8         InsertInString = InToThis
9         Exit Function
ErrHandler:
10        Throw "#InsertInString (line " & CStr(Erl) + "): " & Err.Description & "!"
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

1         On Error GoTo ErrHandler

2         If Not RegExSyntaxValid(RegularExpression) Then
3             IsRegMatch = "#Invalid syntax for RegularExpression!"
4             Exit Function
5         End If
6         Set rx = New RegExp
7         With rx
8             .IgnoreCase = Not CaseSensitive
9             .Pattern = RegularExpression
10            .Global = False        'Find first match only
11        End With

12        If VarType(StringToSearch) = vbString Then
13            IsRegMatch = rx.Test(StringToSearch)

14            GoTo EarlyExit
15        ElseIf VarType(StringToSearch) < vbArray Then
16            IsRegMatch = "#StringToSearch must be a string!"
17            GoTo EarlyExit
18        End If
19        If TypeName(StringToSearch) = "Range" Then StringToSearch = StringToSearch.Value2

20        Select Case NumDimensions(StringToSearch)
              Case 2
21                ReDim Result(LBound(StringToSearch, 1) To UBound(StringToSearch, 1), LBound(StringToSearch, 2) To UBound(StringToSearch, 2))
22                For i = LBound(StringToSearch, 1) To UBound(StringToSearch, 1)
23                    For j = LBound(StringToSearch, 2) To UBound(StringToSearch, 2)
24                        If VarType(StringToSearch(i, j)) = vbString Then
25                            Result(i, j) = rx.Test(StringToSearch(i, j))
26                        Else
27                            Result(i, j) = "#StringToSearch must be a string!"
28                        End If
29                    Next j
30                Next i
31            Case 1
32                ReDim Result(LBound(StringToSearch, 1) To UBound(StringToSearch, 1))
33                For i = LBound(StringToSearch, 1) To UBound(StringToSearch, 1)
34                    If VarType(StringToSearch(i)) = vbString Then
35                        Result(i) = rx.Test(StringToSearch(i))
36                    Else
37                        Result(i) = "#StringToSearch must be a string!"
38                    End If
39                Next i
40            Case Else
41                Throw "StringToSearch must be String or array with 1 or 2 dimensions"
42        End Select

43        IsRegMatch = Result
EarlyExit:
44        Set rx = Nothing

45        Exit Function
ErrHandler:
46        IsRegMatch = "#IsRegMatch (line " & CStr(Erl) + "): " & Err.Description & "!"
47        Set rx = Nothing
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RegExSyntaxValid
' Purpose    : Tests syntax of a regular expression.
' -----------------------------------------------------------------------------------------------------------------------
Private Function RegExSyntaxValid(RegularExpression As String) As Boolean
          Dim res As Boolean
          Dim rx As VBScript_RegExp_55.RegExp
1         On Error GoTo ErrHandler
2         Set rx = New RegExp
3         With rx
4             .IgnoreCase = False
5             .Pattern = RegularExpression
6             .Global = False        'Find first match only
7         End With
8         res = rx.Test("Foo")
9         RegExSyntaxValid = True
10        Exit Function
ErrHandler:
11        RegExSyntaxValid = False
End Function

