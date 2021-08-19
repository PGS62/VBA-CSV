Attribute VB_Name = "modCSVTestDeps"
'This module contains "test dependencies" of CSVReadWrite, i.e. dependencies of the code used to test ModCSVReadWrite, _
but not dependencies of ModCSVReadWrite itself which is (should be) self-contained
Option Explicit
Option Private Module

#If VBA7 Then
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
#Else
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
#End If

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : TestCSVRead
' Purpose    : Kernal of the method RunTests, uses sArryasIdentical to check that data read by function CSVRead is
'              identical to Expected. If not sets WhatDiffers to a description of what went wrong.
' -----------------------------------------------------------------------------------------------------------------------
Function TestCSVRead(CaseNo As Long, ByVal TestDescription As String, Expected As Variant, FileName As String, _
          ByRef WhatDiffers As String, Optional ConvertTypes As Variant = False, _
          Optional ByVal Delimiter As Variant, Optional IgnoreRepeated As Boolean, _
          Optional DateFormat As String, Optional Comment As String, Optional ByVal SkipToRow As Long = 1, _
          Optional ByVal SkipToCol As Long = 1, Optional ByVal NumRows As Long = 0, _
          Optional ByVal NumCols As Long = 0, Optional ByVal ShowMissingsAs As Variant = "", _
          Optional ByVal Encoding As Variant, Optional DecimalSeparator As String = vbNullString) As Boolean

          Dim Observed As Variant

1         On Error GoTo ErrHandler

2         WhatDiffers = ""
3         TestDescription = "Case " + CStr(CaseNo) + " " + TestDescription

4         Observed = CSVRead(FileName, ConvertTypes, Delimiter, IgnoreRepeated, DateFormat, Comment, SkipToRow, _
              SkipToCol, NumRows, NumCols, ShowMissingsAs, Encoding, DecimalSeparator)

5         If VarType(Observed) = vbString Then
6             If VarType(Expected) = vbString Then
7                 If Observed = Expected Then
8                     TestCSVRead = True
9                     Exit Function
10                Else
11                    WhatDiffers = TestDescription + " FAILED, CSVRead returned error: '" + Observed + _
                          "' but expected a different error: '" + Expected + "'"
12                    GoTo Failed
13                End If
14            Else
15                WhatDiffers = TestDescription + " FAILED, CSVRead returned error: '" + Observed + "'"
16                GoTo Failed
17            End If
18        End If

19        If NumDimensions(Observed) = 2 And NumDimensions(Expected) = 2 Then
20            If sArraysIdentical(Observed, Expected, True, False, WhatDiffers) Then
21                TestCSVRead = True
22                Exit Function
23            Else
24                WhatDiffers = TestDescription + " FAILED, observed and expected differed: " + WhatDiffers
25                GoTo Failed
26            End If
27        Else
28            TestCSVRead = False
29            WhatDiffers = TestDescription + " FAILED, observed has " + CStr(NumDimensions(Observed)) + _
                  " dimensions, expected has " + CStr(NumDimensions(Expected)) + " dimensions"
30        End If

Failed:
31        Debug.Print WhatDiffers
32        TestCSVRead = False

33        Exit Function
ErrHandler:
34        Throw "#TestCSVRead (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function


Function NameThatFile(Folder As String, ByVal OS As String, NumRows As Long, NumCols As Long, ExtraInfo As String, Unicode As Boolean, Ragged As Boolean)
    NameThatFile = (Folder & "\" & IIf(ExtraInfo = "", "", ExtraInfo & "_") & OS & "_" & Format(NumRows, "0000") & "_x_" & Format(NumCols, "000") & IIf(Unicode, "_Unicode", "_Ascii") & IIf(Ragged, "_Ragged", "") & ".csv")
End Function

'---------------------------------------------------------------------------------------------------------
' Procedure : CreatePath
' Purpose   : Creates a folder on disk. FolderPath can be passed in as C:\This\That\TheOther even if the
'             folder C:\This does not yet exist. If successful returns the name of the
'             folder. If not successful returns an error string.
' Arguments
' FolderPath: Path of the folder to be created. For example C:\temp\My_New_Folder. It does not matter if
'             this path has a terminating backslash or not.
'---------------------------------------------------------------------------------------------------------
Function CreatePath(ByVal FolderPath As String)

    Dim F As Scripting.Folder
    Dim FSO As Scripting.FileSystemObject
    Dim i As Long
    Dim isUNC As Boolean
    Dim ParentFolderName

    On Error GoTo ErrHandler

    If Left$(FolderPath, 2) = "\\" Then
        isUNC = True
    ElseIf Mid$(FolderPath, 2, 2) <> ":\" Or Asc(UCase$(Left$(FolderPath, 1))) < 65 Or Asc(UCase$(Left$(FolderPath, 1))) > 90 Then
        Throw "First three characters of FolderPath must give drive letter followed by "":\"" or else be""\\"" for " & _
            "UNC folder name"
    End If

    FolderPath = Replace(FolderPath, "/", "\")

    If Right$(FolderPath, 1) <> "\" Then
        FolderPath = FolderPath + "\"
    End If

    Set FSO = New FileSystemObject
    If FolderExists(FolderPath) Then
        GoTo EarlyExit
    End If

    'Go back until we find parent folder that does exist
    For i = Len(FolderPath) - 1 To 3 Step -1
        If Mid$(FolderPath, i, 1) = "\" Then
            If FolderExists(Left$(FolderPath, i)) Then
                Set F = FSO.GetFolder(Left$(FolderPath, i))
                ParentFolderName = Left$(FolderPath, i)
                Exit For
            End If
        End If
    Next i

    If F Is Nothing Then Throw "Cannot create folder " + Left$(FolderPath, 3)

    'now add folders one level at a time
    For i = Len(ParentFolderName) + 1 To Len(FolderPath)
        If Mid$(FolderPath, i, 1) = "\" Then
            Dim ThisFolderName As String
            ThisFolderName = Mid$(FolderPath, InStrRev(FolderPath, "\", i - 1) + 1, i - 1 - InStrRev(FolderPath, "\", i - 1))
            F.SubFolders.Add ThisFolderName
            Set F = FSO.GetFolder(Left$(FolderPath, i))
        End If
    Next i

EarlyExit:
    Set F = FSO.GetFolder(FolderPath)
    CreatePath = F.path
    Set F = Nothing: Set FSO = Nothing

    Exit Function
ErrHandler:
    CreatePath = "#CreatePath: " & Err.Description & "!"
End Function
'---------------------------------------------------------------------------------------
' Procedure : FolderExists
' Purpose   : Returns True or False. Does not matter if FolderPath has a terminating backslash or not.
'---------------------------------------------------------------------------------------
Function FolderExists(ByVal FolderPath As String)
    Dim F As Scripting.Folder
    Dim FSO As Scripting.FileSystemObject
    On Error GoTo ErrHandler
    Set FSO = New FileSystemObject
    Set F = FSO.GetFolder(FolderPath)
    FolderExists = True
    Exit Function
ErrHandler:
    FolderExists = False
End Function
'
'---------------------------------------------------------------------------------------------------------
' Procedure : sElapsedTime
' Purpose   : Retrieves the current value of the performance counter, which is a high resolution (<1us)
'             time stamp that can be used for time-interval measurements.
'
'             See http://msdn.microsoft.com/en-us/library/windows/desktop/ms644904(v=vs.85).aspx
'---------------------------------------------------------------------------------------------------------
Function sElapsedTime() As Double
    Dim a As Currency
    Dim b As Currency
    On Error GoTo ErrHandler

    QueryPerformanceCounter a
    QueryPerformanceFrequency b
    sElapsedTime = a / b

    Exit Function
ErrHandler:
    Throw "#sElapsedTime: " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : sNCols
' Purpose   : Number of columns in an array. Missing has zero rows, 1-dimensional arrays
'             have one row and the number of columns returned by this function.
'---------------------------------------------------------------------------------------
Function sNCols(Optional TheArray) As Long
    If TypeName(TheArray) = "Range" Then
        sNCols = TheArray.Columns.Count
    ElseIf IsMissing(TheArray) Then
        sNCols = 0
    ElseIf VarType(TheArray) < vbArray Then
        sNCols = 1
    Else
        Select Case NumDimensions(TheArray)
            Case 1
                sNCols = UBound(TheArray, 1) - LBound(TheArray, 1) + 1
            Case Else
                sNCols = UBound(TheArray, 2) - LBound(TheArray, 2) + 1
        End Select
    End If
End Function

'Copy of identical function in modCVSReadWrite
Function NumDimensions(x As Variant) As Long
    Dim i As Long
    Dim y As Long
    If Not IsArray(x) Then
        NumDimensions = 0
        Exit Function
    Else
        On Error GoTo ExitPoint
        i = 1
        Do While True
            y = LBound(x, i)
            i = i + 1
        Loop
    End If
ExitPoint:
    NumDimensions = i - 1
End Function

'---------------------------------------------------------------------------------------
' Procedure : sNRows
' Purpose   : Number of rows in an array. Missing has zero rows, 1-dimensional arrays have one row.
'---------------------------------------------------------------------------------------
Function sNRows(Optional TheArray) As Long
    If TypeName(TheArray) = "Range" Then
        sNRows = TheArray.Rows.Count
    ElseIf IsMissing(TheArray) Then
        sNRows = 0
    ElseIf VarType(TheArray) < vbArray Then
        sNRows = 1
    Else
        Select Case NumDimensions(TheArray)
            Case 1
                sNRows = 1
            Case Else
                sNRows = UBound(TheArray, 1) - LBound(TheArray, 1) + 1
        End Select
    End If
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

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Throw
' Purpose    : Simple error handling.
' -----------------------------------------------------------------------------------------------------------------------
Public Sub Throw(ByVal ErrorString As String)
    Err.Raise vbObjectError + 1, , ErrorString
End Sub

'---------------------------------------------------------------------------------------
' Procedure : CreateMissing
' Purpose   : Returns a variant of type Missing
'---------------------------------------------------------------------------------------
Function CreateMissing()
    CreateMissing = CM2()
End Function
Function CM2(Optional OptionalArg As Variant)
    CM2 = OptionalArg
End Function

'---------------------------------------------------------------------------------------
' Procedure : Force2DArray
' Purpose   : In-place amendment of singletons and one-dimensional arrays to two dimensions.
'             singletons and 1-d arrays are returned as 2-d 1-based arrays. Leaves two
'             two dimensional arrays untouched (i.e. a zero-based 2-d array will be left as zero-based).
'             See also Force2DArrayR that also handles Range objects.
'---------------------------------------------------------------------------------------
Sub Force2DArray(ByRef TheArray As Variant, Optional ByRef NR As Long, Optional ByRef NC As Long)
    Dim TwoDArray As Variant

    On Error GoTo ErrHandler

    Select Case NumDimensions(TheArray)
        Case 0
            ReDim TwoDArray(1 To 1, 1 To 1)
            TwoDArray(1, 1) = TheArray
            TheArray = TwoDArray
            NR = 1: NC = 1
        Case 1
            Dim i As Long
            Dim LB As Long
            LB = LBound(TheArray, 1)
            NR = 1: NC = UBound(TheArray, 1) - LB + 1
            ReDim TwoDArray(1 To 1, 1 To NC)
            For i = 1 To UBound(TheArray, 1) - LBound(TheArray) + 1
                TwoDArray(1, i) = TheArray(LB + i - 1)
            Next i
            TheArray = TwoDArray
        Case 2
            NR = UBound(TheArray, 1) - LBound(TheArray, 1) + 1
            NC = UBound(TheArray, 2) - LBound(TheArray, 2) + 1
            'Nothing to do
        Case Else
            Throw "Cannot convert array of dimension greater than two"
    End Select

    Exit Sub
ErrHandler:
    Throw "#Force2DArray (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Force2DArrayR
' Purpose   : When writing functions to be called from sheets, we often don't want to process
'             the inputs as Range objects, but instead as Arrays. This method converts the
'             input into a 2-dimensional 1-based array (even if it's a single cell or single row of cells)
'---------------------------------------------------------------------------------------
Sub Force2DArrayR(ByRef RangeOrArray As Variant, Optional ByRef NR As Long, Optional ByRef NC As Long)
    If TypeName(RangeOrArray) = "Range" Then RangeOrArray = RangeOrArray.Value2
    Force2DArray RangeOrArray, NR, NC
End Sub

Function SafeMin(a, b)
    On Error GoTo ErrHandler
    If Not IsNumberOrDate(a) Then
        SafeMin = "#Non-number found!"
    ElseIf Not IsNumberOrDate(b) Then
        SafeMin = "#Non-number found!"
    ElseIf a > b Then
        SafeMin = b
    Else
        SafeMin = a
    End If
    Exit Function
ErrHandler:
    SafeMin = "#" & Err.Description & "!"
End Function

Function SafeMax(a, b)
    On Error GoTo ErrHandler
    If Not IsNumberOrDate(a) Then
        SafeMax = "#Non-number found!"
    ElseIf Not IsNumberOrDate(b) Then
        SafeMax = "#Non-number found!"
    ElseIf a > b Then
        SafeMax = a
    Else
        SafeMax = b
    End If
    Exit Function
ErrHandler:
    SafeMax = "#" & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : IsNumberOrDate
' Purpose   : Is a singleton a number or date
'---------------------------------------------------------------------------------------
Function IsNumberOrDate(x As Variant) As Boolean
    Select Case VarType(x)
        Case vbDouble, vbInteger, vbSingle, vbLong, vbDate
            IsNumberOrDate = True
    End Select
End Function

'---------------------------------------------------------------------------------------
' Procedure : SafeSubtract
' Purpose   : low-level subtraction with error handling
'---------------------------------------------------------------------------------------
Function SafeSubtract(a, b)
    On Error GoTo ErrHandler
    If Not IsNumberOrDate(a) Then
        SafeSubtract = "#Non-number found!"
    ElseIf Not IsNumberOrDate(b) Then
        SafeSubtract = "#Non-number found!"
    Else
        SafeSubtract = a - b
    End If
    Exit Function
ErrHandler:
    SafeSubtract = "#" & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------------------------
' Procedure : sEquals
' Purpose   : Returns TRUE if a is equal to b, FALSE otherwise. a and b may be numbers, strings,
'             Booleans or Excel error values, but not arrays. For testing equality of
'             arrays see ArrayEquals and sArraysIdentical.
'             Examples
'             sEquals(1,1) = TRUE
'             sEquals(#DIV0!,1) = FALSE
' Arguments
' a         : The first value to compare. Must be a single cell (not an array) but can contain numbers,
'             text, logical value or error values.
' b         : The second value to compare. Must be a single cell (not an array) but can contain numbers,
'             text, logical value or error values.
' CaseSensitive: Determines if comparison of strings is done in a case sensitive manner. If omitted
'             defaults to FALSE (case insensitive matching).
'
'Note:        Avoids VBA booby trap that False = 0 and True = -1
'---------------------------------------------------------------------------------------
Function sEquals(a, b, Optional CaseSensitive As Boolean = False) As Variant
    On Error GoTo ErrHandler
    Dim VTA As Long
    Dim VTB As Long

    VTA = VarType(a)
    VTB = VarType(b)
    If VTA >= vbArray Or VTB >= vbArray Then
        sEquals = "#sEquals: Function does not handle arrays. Use sArrayEquals or sArraysIdentical instead!"
        Exit Function
    End If

    If VTA = VTB Then
        If VTA = vbString And Not CaseSensitive Then
            If Len(a) = Len(b) Then
                sEquals = UCase$(a) = UCase$(b)
            Else
                sEquals = False
            End If
        Else
            sEquals = (a = b)
        End If
    Else
        If VTA = vbBoolean Or VTB = vbBoolean Or VTA = vbString Or VTB = vbString Then
            sEquals = False
        Else
            sEquals = (a = b)
        End If
    End If
    Exit Function
ErrHandler:
    sEquals = False
End Function

'---------------------------------------------------------------------------------------
' Procedure : NonStringToString
' Purpose   : Convert non-string to string in a way that mimics how the non-string would
'             be displayed in an Excel cell. Used by functions such as ConcatenateStrings
'             and Examine (aka g)
'---------------------------------------------------------------------------------------
Function NonStringToString(x As Variant, Optional AddSingleQuotesToStings As Boolean = False)
    Dim res As String
    On Error GoTo ErrHandler
    If IsError(x) Then
        Select Case CStr(x)
            Case "Error 2007"
                res = "#DIV/0!"
            Case "Error 2029"
                res = "#NAME?"
            Case "Error 2023"
                res = "#REF!"
            Case "Error 2036"
                res = "#NUM!"
            Case "Error 2000"
                res = "#NULL!"
            Case "Error 2042"
                res = "#N/A"
            Case "Error 2015"
                res = "#VALUE!"
            Case "Error 2045"
                res = "#SPILL!"
            Case "Error 2047"
                res = "#BLOCKED!"
            Case "Error 2046"
                res = "#CONNECT!"
            Case "Error 2048"
                res = "#UNKNOWN!"
            Case "Error 2043"
                res = "#GETTING_DATA!"
            Case Else
                res = CStr(x)        'should never hit this line...
        End Select
    ElseIf VarType(x) = vbDate Then
        If CDbl(x) = CLng(x) Then
            res = Format$(x, "dd-mmm-yyyy")
        Else
            res = Format$(x, "dd-mmm-yyyy hh:mm:ss")
        End If
    ElseIf IsNull(x) Then
        res = "null" 'Follow how json represents Null as lower-case null
    ElseIf VarType(x) = vbString And AddSingleQuotesToStings Then
        res = "'" + x + "'"
    Else
        res = SafeCStr(x)        'Converts Empty to null string. Prior to 15 Jan 2017 Empty was converted to "Empty"
    End If
    NonStringToString = res
    Exit Function
ErrHandler:
    Throw "#NonStringToString (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function SafeCStr(x As Variant)
    On Error GoTo ErrHandler
    SafeCStr = CStr(x)
    Exit Function
ErrHandler:
    SafeCStr = "#Cannot represent " + TypeName(x) + "!"
End Function

'---------------------------------------------------------------------------------------------------------
' Procedure : Transpose
' Purpose   : Returns the transpose of an array.
' Arguments
' TheArray  : An array of arbitrary values.
'             also converts 0-based to 1-based arrays
'---------------------------------------------------------------------------------------------------------
Function Transpose(ByVal TheArray As Variant)
    Dim Co As Long
    Dim i As Long
    Dim j As Long
    Dim m As Long
    Dim N As Long
    Dim Result As Variant
    Dim Ro As Long
    On Error GoTo ErrHandler
    Force2DArrayR TheArray, N, m
    Ro = LBound(TheArray, 1) - 1
    Co = LBound(TheArray, 2) - 1
    ReDim Result(1 To m, 1 To N)
    For i = 1 To N
        For j = 1 To m
            Result(j, i) = TheArray(i + Ro, j + Co)
        Next j
    Next i
    Transpose = Result
    Exit Function
ErrHandler:
    Transpose = "#Transpose (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------------------------
' Procedure : HStack
' Purpose   : Places arrays horizontally side by side. If the arrays are of unequal height then they will be padded
'             underneath with #NA! values.
'  Notes   1) Input arrays to range can have 0,1, or 2 dimensions
'          2) output array has lower bound 1, whatever the lower bounds of the inputs
'          3) input arrays of 1 dimension are treated as if they were columns, different from SAI equivalent fn.
'---------------------------------------------------------------------------------------------------------
Function HStack(ParamArray Arrays())

          Dim AllC As Long
          Dim AllR As Long
          Dim C As Long
          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim R As Long
          Dim ReturnArray()
          Dim Y0 As Long

1         On Error GoTo ErrHandler

          Static NA As Variant
2         If IsEmpty(NA) Then NA = CVErr(xlErrNA)

3         If IsMissing(Arrays) Then
4             HStack = CreateMissing()
5         Else
6             For i = LBound(Arrays) To UBound(Arrays)
7                 If TypeName(Arrays(i)) = "Range" Then Arrays(i) = Arrays(i).value
8                 If IsMissing(Arrays(i)) Then
9                     R = 0: C = 0
10                Else
11                    Select Case NumDimensions(Arrays(i))
                          Case 0
12                            R = 1: C = 1
13                        Case 1
14                            R = UBound(Arrays(i)) - LBound(Arrays(i)) + 1
15                            C = 1
16                        Case 2
17                            R = UBound(Arrays(i), 1) - LBound(Arrays(i), 1) + 1
18                            C = UBound(Arrays(i), 2) - LBound(Arrays(i), 2) + 1
19                    End Select
20                End If
21                If R > AllR Then AllR = R
22                AllC = AllC + C
23            Next i

24            If AllR = 0 Then
25                HStack = CreateMissing()
26                Exit Function
27            End If

28            ReDim ReturnArray(1 To AllR, 1 To AllC)

29            Y0 = 1
30            For i = LBound(Arrays) To UBound(Arrays)
31                If Not IsMissing(Arrays(i)) Then
32                    Select Case NumDimensions(Arrays(i))
                          Case 0
33                            R = 1: C = 1
34                            ReturnArray(1, Y0) = Arrays(i)
35                        Case 1
36                            R = UBound(Arrays(i)) - LBound(Arrays(i)) + 1
37                            C = 1
38                            For j = 1 To R
39                                ReturnArray(j, Y0) = Arrays(i)(j + LBound(Arrays(i)) - 1)
40                            Next j
41                        Case 2
42                            R = UBound(Arrays(i), 1) - LBound(Arrays(i), 1) + 1
43                            C = UBound(Arrays(i), 2) - LBound(Arrays(i), 2) + 1

44                            For j = 1 To R
45                                For k = 1 To C
46                                    ReturnArray(j, Y0 + k - 1) = Arrays(i)(j + LBound(Arrays(i), 1) - 1, k + LBound(Arrays(i), 2) - 1)
47                                Next k
48                            Next j

49                    End Select
50                    If R < AllR Then        'Pad with #NA! values
51                        For j = R + 1 To AllR
52                            For k = 1 To C
53                                ReturnArray(j, Y0 + k - 1) = NA
54                            Next k
55                        Next j
56                    End If

57                    Y0 = Y0 + C
58                End If
59            Next i
60            HStack = ReturnArray
61        End If

62        Exit Function
ErrHandler:
63        HStack = "#HStack (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------------------------
' Procedure : sArraysIdentical
' Purpose   : Returns TRUE if the two input arrays are identical. That is, they are the same size and
'             shape and every pair of elements are equal.
'
' Arguments
' Array1    : The first array to compare.
' Array2    : The second array to compare.
' CaseSensitive: TRUE for case sensitive comparison of strings. FALSE or omitted for case insensitive
'             comparison.
' PermitBaseDifference: This argument is not relevant when using the function in an Excel formula and should be
'             omitted. If used from VBA code, then setting it to TRUE allows "zero-based"
'             arrays to be compared with "one-based" arrays.
' WhatDiffers: passed by reference and set to a string describing the differences found.
'---------------------------------------------------------------------------------------------------------
Function sArraysIdentical(ByVal Array1, ByVal Array2, Optional CaseSensitive As Boolean, _
          Optional PermitBaseDifference As Boolean = False, Optional ByRef WhatDiffers As String) As Variant
          
          Dim cN As Long
          Dim i As Long
          Dim j As Long
          Dim rN As Long
          Dim NumSame As Long
          Dim NumDiff As Long
          
1         On Error GoTo ErrHandler

          'flips 1-d to 2-d with one row
2         Force2DArrayR Array1: Force2DArrayR Array2

          WhatDiffers = ""
3         If (UBound(Array1, 1) - LBound(Array1, 1)) <> (UBound(Array2, 1) - LBound(Array2, 1)) Then
4             WhatDiffers = "Row count different: " + CStr(1 + (UBound(Array1, 1) - LBound(Array1, 1))) + " vs " _
                  + CStr(1 + (UBound(Array2, 1) - LBound(Array2, 1)))
5             sArraysIdentical = False
6         ElseIf (UBound(Array1, 2) - LBound(Array1, 2)) <> (UBound(Array2, 2) - LBound(Array2, 2)) Then
7             WhatDiffers = "Column count different: " + CStr(1 + (UBound(Array1, 2) - LBound(Array1, 2))) + " vs " _
                  + CStr(1 + (UBound(Array2, 2) - LBound(Array2, 2)))
8             sArraysIdentical = False
9         Else
10            If Not PermitBaseDifference Then
11                If (LBound(Array1, 1) <> LBound(Array2, 1)) Or (LBound(Array1, 2) <> LBound(Array2, 2)) Then
12                    WhatDiffers = "Lower bounds different"
13                    sArraysIdentical = False
14                    Exit Function
15                End If
16            End If
17            rN = LBound(Array2, 1) - LBound(Array1, 1)
18            cN = LBound(Array2, 2) - LBound(Array1, 2)
19            For i = LBound(Array1, 1) To UBound(Array1, 1)
20                For j = LBound(Array1, 2) To UBound(Array1, 2)
21                    If Not sEquals(Array1(i, j), Array2(i + rN, j + cN), CaseSensitive) Then
22                        NumDiff = NumDiff + 1

23                        WhatDiffers = "first difference at " + CStr(i) + "," + CStr(j) + ": " + _
                        TypeName(Array1(i, j)) + " '" + CStr(Array1(i, j)) + "' vs " + _
                              TypeName(Array2(i + rN, j + cN)) + " '" + CStr(Array2(i + rN, j + cN)) + "'"
24                        sArraysIdentical = False
26                    Else
27                        NumSame = NumSame + 1
28                    End If
29                Next j
30            Next i
31            If NumDiff = 0 Then
32                sArraysIdentical = True
33            Else
34                sArraysIdentical = False
35                WhatDiffers = CStr(NumDiff) + " of " + CStr(NumDiff + NumSame) + " elements differ, " + WhatDiffers
36            End If

37        End If

38        Exit Function
ErrHandler:
39        sArraysIdentical = "#sArraysIdentical (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function


Function FileCopy(SourceFile As String, TargetFile As String)
          Dim F As Scripting.File
          Dim FSO As Scripting.FileSystemObject
          Dim CopyOfErr As String
1         On Error GoTo ErrHandler
2         Set FSO = New FileSystemObject
3         Set F = FSO.GetFile(SourceFile)
4         F.Copy TargetFile, True
5         FileCopy = TargetFile
6         Set FSO = Nothing: Set F = Nothing
7         Exit Function
ErrHandler:
8         CopyOfErr = Err.Description
9         Set FSO = Nothing: Set F = Nothing
10        Throw "#" + CopyOfErr + "!"
End Function

