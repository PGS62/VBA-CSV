Attribute VB_Name = "modCSVTestDeps"

' VBA-CSV

' Copyright (C) 2021 - Philip Swannell (https://github.com/PGS62/VBA-CSV )
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/VBA-CSV#readme

'This module contains "test dependencies" of CSVReadWrite, i.e. dependencies of the code used to test ModCSVReadWrite, _
 but not dependencies of ModCSVReadWrite itself which is (should be) self-contained.

Option Explicit
Option Private Module

#If VBA7 And Win64 Then
    Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
    Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
#Else
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
    Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
#End If

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : TestCSVRead
' Purpose    : Kernel of the method RunTests, uses sArryasIdentical to check that data read by function CSVRead is
'              identical to Expected. If not, sets WhatDiffers to a description of what went wrong.
' -----------------------------------------------------------------------------------------------------------------------
Function TestCSVRead(TestNo As Long, ByVal TestDescription As String, Expected As Variant, FileName As String, ByRef Observed, _
    ByRef WhatDiffers As String, Optional AbsTol As Double, Optional RelTol As Double, Optional ConvertTypes As Variant = False, _
    Optional ByVal Delimiter As Variant, Optional IgnoreRepeated As Boolean, _
    Optional DateFormat As String, Optional Comment As String, Optional IgnoreEmptyLines As Boolean = True, Optional ByVal SkipToRow As Long = 0, _
    Optional ByVal SkipToCol As Long = 1, Optional ByVal NumRows As Long = 0, _
    Optional ByVal NumCols As Long = 0, Optional HeaderRowNum As Long, Optional TrueStrings As Variant, Optional FalseStrings As Variant, _
    Optional MissingStrings As Variant, Optional ByVal ShowMissingsAs As Variant = "", _
    Optional ByVal Encoding As Variant, Optional DecimalSeparator As String = vbNullString, _
    Optional NumRowsExpected As Long, Optional NumColsExpected As Long, Optional ByRef HeaderRow, Optional ExpectedHeaderRow) As Boolean

    On Error GoTo ErrHandler
    Const PermitBaseDifference = True

    WhatDiffers = ""
    TestDescription = "Test " + CStr(TestNo) + " " + TestDescription

    Observed = CSVRead(FileName, ConvertTypes, Delimiter, IgnoreRepeated, DateFormat, Comment, IgnoreEmptyLines, HeaderRowNum, SkipToRow, _
        SkipToCol, NumRows, NumCols, TrueStrings, FalseStrings, MissingStrings, ShowMissingsAs, Encoding, DecimalSeparator, HeaderRow)
        
    If Not IsMissing(ExpectedHeaderRow) Then
        If Not ArraysIdentical(ExpectedHeaderRow, HeaderRow, True, PermitBaseDifference, WhatDiffers) Then
            WhatDiffers = TestDescription + " FAILED. HeaderRow failed to match ExpectedHeaderRow: " & WhatDiffers
            GoTo Failed
        End If
    End If

    If NumRowsExpected <> 0 Or NumColsExpected <> 0 Then
        If NRows(Observed) <> NumRowsExpected Or NCols(Observed) <> NumColsExpected Then
            WhatDiffers = TestDescription + " FAILED, expected dimensions: " + CStr(NumRowsExpected) + _
                ", " + CStr(NumColsExpected) + " observed dimensions: " + CStr(NRows(Observed)) + ", " + CStr(NCols(Observed))
            GoTo Failed
        ElseIf IsEmpty(Expected) Then
            TestCSVRead = True
            Exit Function
        End If
    End If

    If VarType(Observed) = vbString Then
        If VarType(Expected) = vbString Then
            If Observed = Expected Then
                TestCSVRead = True
                Exit Function
            ElseIf RegExSyntaxValid(CStr(Expected)) Then
                If IsRegMatch(CStr(Expected), CStr(Observed)) Then
                    TestCSVRead = True
                    Exit Function
                Else
                    WhatDiffers = TestDescription + " FAILED, CSVRead returned error: '" + Observed + _
                        "' but expected a different error: '" + Expected + "'"
                    GoTo Failed
                End If
            Else
                WhatDiffers = TestDescription + " FAILED, CSVRead returned error: '" + Observed + _
                    "' but expected a different error: '" + Expected + "'"
                GoTo Failed
            End If
        Else
            WhatDiffers = TestDescription + " FAILED, CSVRead returned error: '" + Observed + "'"
            GoTo Failed
        End If
    End If

    If NumDimensions(Observed) = 2 And NumDimensions(Expected) = 2 Then
        If ArraysIdentical(Observed, Expected, True, PermitBaseDifference, WhatDiffers, AbsTol, RelTol) Then
            TestCSVRead = True
            Exit Function
        Else
            WhatDiffers = TestDescription + " FAILED, observed and expected differed: " + WhatDiffers
            GoTo Failed
        End If
    Else
        TestCSVRead = False
        WhatDiffers = TestDescription + " FAILED, observed has " + CStr(NumDimensions(Observed)) + _
            " dimensions, expected has " + CStr(NumDimensions(Expected)) + " dimensions"
    End If

Failed:
    Debug.Print WhatDiffers
    TestCSVRead = False

    Exit Function
ErrHandler:
    Throw "#TestCSVRead (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function NameThatFile(Folder As String, ByVal OS As String, NumRows As Long, NumCols As Long, ExtraInfo As String, Encoding As String, Ragged As Boolean)
    NameThatFile = (Folder & "\" & IIf(ExtraInfo = "", "", ExtraInfo & "_") & IIf(OS = "", "", OS & "_") & Format(NumRows, "0000") & "_x_" & Format(NumCols, "000") & "_" & Encoding & IIf(Ragged, "_Ragged", "") & ".csv")
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : NCols
' Purpose   : Number of columns in an array. Missing has zero rows, 1-dimensional arrays
'             have one row and the number of columns returned by this function.
' -----------------------------------------------------------------------------------------------------------------------
Function NCols(Optional TheArray) As Long
    If TypeName(TheArray) = "Range" Then
        NCols = TheArray.Columns.Count
    ElseIf IsMissing(TheArray) Then
        NCols = 0
    ElseIf VarType(TheArray) < vbArray Then
        NCols = 1
    Else
        Select Case NumDimensions(TheArray)
            Case 1
                NCols = UBound(TheArray, 1) - LBound(TheArray, 1) + 1
            Case Else
                NCols = UBound(TheArray, 2) - LBound(TheArray, 2) + 1
        End Select
    End If
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : NRows
' Purpose   : Number of rows in an array. Missing has zero rows, 1-dimensional arrays have one row.
' -----------------------------------------------------------------------------------------------------------------------
Function NRows(Optional TheArray) As Long
    If TypeName(TheArray) = "Range" Then
        NRows = TheArray.Rows.Count
    ElseIf IsMissing(TheArray) Then
        NRows = 0
    ElseIf VarType(TheArray) < vbArray Then
        NRows = 1
    Else
        Select Case NumDimensions(TheArray)
            Case 1
                NRows = 1
            Case Else
                NRows = UBound(TheArray, 1) - LBound(TheArray, 1) + 1
        End Select
    End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CreatePath
' Purpose   : Creates a folder on disk. FolderPath can be passed in as C:\This\That\TheOther even if the
'             folder C:\This does not yet exist. If successful returns the name of the
'             folder. If not successful returns an error string.
' Arguments
' FolderPath: Path of the folder to be created. For example C:\temp\My_New_Folder. It does not matter if
'             this path has a terminating backslash or not.
' -----------------------------------------------------------------------------------------------------------------------
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

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FolderExists
' Purpose   : Returns True or False. Does not matter if FolderPath has a terminating backslash or not.
' -----------------------------------------------------------------------------------------------------------------------
Private Function FolderExists(ByVal FolderPath As String)
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

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ElapsedTime
' Purpose    : Retrieves the current value of the performance counter, which is a high resolution (<1us)
'              time stamp that can be used for time-interval measurements.
'
'              See http://msdn.microsoft.com/en-us/library/windows/desktop/ms644904(v=vs.85).aspx
' -----------------------------------------------------------------------------------------------------------------------
Function ElapsedTime() As Double
    Dim a As Currency
    Dim b As Currency
    On Error GoTo ErrHandler

    QueryPerformanceCounter a
    QueryPerformanceFrequency b
    ElapsedTime = a / b

    Exit Function
ErrHandler:
    Throw "#ElapsedTime: " & Err.Description & "!"
End Function

'Copy of identical function in modCSVReadWrite
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

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Throw
' Purpose    : Simple error handling.
' -----------------------------------------------------------------------------------------------------------------------
Public Sub Throw(ByVal ErrorString As String)
    Err.Raise vbObjectError + 1, , ErrorString
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CreateMissing
' Purpose   : Returns a variant of type Missing
' -----------------------------------------------------------------------------------------------------------------------
Function CreateMissing()
    CreateMissing = CM2()
End Function
Function CM2(Optional OptionalArg As Variant)
    CM2 = OptionalArg
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Force2DArray
' Purpose   : In-place amendment of singletons and one-dimensional arrays to two dimensions.
'             singletons and 1-d arrays are returned as 2-d 1-based arrays. Leaves two
'             two dimensional arrays untouched (i.e. a zero-based 2-d array will be left as zero-based).
'             See also Force2DArrayR that also handles Range objects.
' -----------------------------------------------------------------------------------------------------------------------
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

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Force2DArrayR
' Purpose   : When writing functions to be called from sheets, we often don't want to process
'             the inputs as Range objects, but instead as Arrays. This method converts the
'             input into a 2-dimensional 1-based array (even if it's a single cell or single row of cells)
' -----------------------------------------------------------------------------------------------------------------------
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

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : IsNumberOrDate
' Purpose   : Is a singleton a number or date
' -----------------------------------------------------------------------------------------------------------------------
Function IsNumberOrDate(x As Variant) As Boolean
    Select Case VarType(x)
        Case vbDouble, vbInteger, vbSingle, vbLong, vbDate
            IsNumberOrDate = True
    End Select
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SafeSubtract
' Purpose   : low-level subtraction with error handling
' -----------------------------------------------------------------------------------------------------------------------
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

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Equals
' Purpose   : Returns TRUE if a is equal to b, FALSE otherwise. a and b may be numbers, strings,
'             Booleans or Excel error values, but not arrays. For testing equality of
'             arrays see ArrayEquals and ArraysIdentical.
'             Examples
'             Equals(1,1) = TRUE
'             Equals(#DIV0!,1) = FALSE
' Arguments
' a         : The first value to compare. Must be a single cell (not an array) but can contain numbers,
'             text, logical value or error values.
' b         : The second value to compare. Must be a single cell (not an array) but can contain numbers,
'             text, logical value or error values.
' CaseSensitive: Determines if comparison of strings is done in a case sensitive manner. If omitted
'             defaults to FALSE (case insensitive matching).
'
'Note:        Avoids VBA booby trap that False = 0 and True = -1
' -----------------------------------------------------------------------------------------------------------------------
Function Equals(a, b, Optional CaseSensitive As Boolean = False) As Variant
    On Error GoTo ErrHandler
    Dim VTA As Long
    Dim VTB As Long

    VTA = VarType(a)
    VTB = VarType(b)
    If VTA >= vbArray Or VTB >= vbArray Then
        Equals = "#Equals: Function does not handle arrays. Use sArrayEquals or ArraysIdentical instead!"
        Exit Function
    End If

    If VTA = VTB Then
        If VTA = vbString And Not CaseSensitive Then
            If Len(a) = Len(b) Then
                Equals = UCase$(a) = UCase$(b)
            Else
                Equals = False
            End If
        Else
            Equals = (a = b)
        End If
    Else
        If VTA = vbBoolean Or VTB = vbBoolean Or VTA = vbString Or VTB = vbString Then
            Equals = False
        Else
            Equals = (a = b)
        End If
    End If
    Exit Function
ErrHandler:
    Equals = False
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : IsApprox
'Purpose:    Inexact equality comparison: for numeric x and y, True if
'            Abs(x-y) <= Max(AbsTol, RelTol*max(Abs(x), Abs(y))).
'            Similar to Julia's function of the same name.
' -----------------------------------------------------------------------------------------------------------------------
Function IsApprox(ByVal x, ByVal y, Optional CaseSensitive As Boolean = False, Optional AbsTol As Double, Optional RelTol As Double)

    Dim CompareTo As Double
    Dim VTA As Long
    Dim VTB As Long

    On Error GoTo ErrHandler

    VTA = VarType(x)
    VTB = VarType(y)
    If VTA >= vbArray Or VTB >= vbArray Then
        IsApprox = "#IsApprox: Function does not handle arrays. Use sArrayNearlyEquals or sArraysNearlyIdentical instead!"
        Exit Function
    End If

    'Both numbers (or dates!)
    If IsNumberOrDate(x) Then
        If IsNumberOrDate(y) Then
            If x = y Then
                IsApprox = True
                Exit Function
            ElseIf AbsTol = 0 Then
                If RelTol = 0 Then
                    IsApprox = False
                    Exit Function
                End If
            End If

            x = CDbl(x): y = CDbl(y)
            CompareTo = Abs(x)
            If Abs(y) > Abs(x) Then
                CompareTo = Abs(y)
            End If
            CompareTo = RelTol * CompareTo
            If AbsTol > CompareTo Then
                CompareTo = AbsTol
            End If
            IsApprox = Abs(x - y) < CompareTo
            Exit Function
        End If
    End If
    'At least one is not x number...
    If VTA = VTB Then
        If VTA = vbString And Not CaseSensitive Then
            If Len(x) = Len(y) Then
                IsApprox = UCase$(x) = UCase$(y)
            Else
                IsApprox = False
            End If
        Else
            IsApprox = (x = y)
        End If
    Else
        If VTA = vbBoolean Or VTB = vbBoolean Or VTA = vbString Or VTB = vbString Then
            IsApprox = False
        Else
            IsApprox = (x = y)
        End If
    End If

    Exit Function
ErrHandler:
    IsApprox = False
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : NonStringToString
' Purpose   : Convert non-string to string in a way that mimics how the non-string would
'             be displayed in an Excel cell. Used by functions such as ConcatenateStrings
'             and Examine (aka g)
' -----------------------------------------------------------------------------------------------------------------------
Private Function NonStringToString(x As Variant, Optional AddSingleQuotesToStings As Boolean = False)
    Dim Res As String
    On Error GoTo ErrHandler
    If IsError(x) Then
        Select Case CStr(x)
            Case "Error 2007"
                Res = "#DIV/0!"
            Case "Error 2029"
                Res = "#NAME?"
            Case "Error 2023"
                Res = "#REF!"
            Case "Error 2036"
                Res = "#NUM!"
            Case "Error 2000"
                Res = "#NULL!"
            Case "Error 2042"
                Res = "#N/A"
            Case "Error 2015"
                Res = "#VALUE!"
            Case "Error 2045"
                Res = "#SPILL!"
            Case "Error 2047"
                Res = "#BLOCKED!"
            Case "Error 2046"
                Res = "#CONNECT!"
            Case "Error 2048"
                Res = "#UNKNOWN!"
            Case "Error 2043"
                Res = "#GETTING_DATA!"
            Case Else
                Res = CStr(x)        'should never hit this line...
        End Select
    ElseIf VarType(x) = vbDate Then
        If CDbl(x) = CLng(x) Then
            Res = Format$(x, "dd-mmm-yyyy")
        Else
            Res = Format$(x, "dd-mmm-yyyy hh:mm:ss")
        End If
    ElseIf IsNull(x) Then
        Res = "null" 'Follow how json represents Null as lower-case null
    ElseIf VarType(x) = vbString And AddSingleQuotesToStings Then
        Res = "'" + x + "'"
    Else
        Res = SafeCStr(x)        'Converts Empty to null string. Prior to 15 Jan 2017 Empty was converted to "Empty"
    End If
    NonStringToString = Res
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

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Transpose
' Purpose   : Returns the transpose of an array.
' Arguments
' TheArray  : An array of arbitrary values.
'             also converts 0-based to 1-based arrays
' -----------------------------------------------------------------------------------------------------------------------
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

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : HStack
' Purpose   : Places arrays horizontally side by side. If the arrays are of unequal height then they will be padded
'             underneath with #NA! values.
'  Notes   1) Input arrays to range can have 0,1, or 2 dimensions
'          2) output array has lower bound 1, whatever the lower bounds of the inputs
'          3) input arrays of 1 dimension are treated as if they were columns, different from SAI equivalent fn.
' -----------------------------------------------------------------------------------------------------------------------
Function HStack(ParamArray Arrays())

    Dim AllC As Long
    Dim AllR As Long
    Dim c As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim r As Long
    Dim ReturnArray()
    Dim Y0 As Long

    On Error GoTo ErrHandler

    Static NA As Variant
    If IsEmpty(NA) Then NA = CVErr(xlErrNA)

    If IsMissing(Arrays) Then
        HStack = CreateMissing()
    Else
        For i = LBound(Arrays) To UBound(Arrays)
            If TypeName(Arrays(i)) = "Range" Then Arrays(i) = Arrays(i).value
            If IsMissing(Arrays(i)) Then
                r = 0: c = 0
            Else
                Select Case NumDimensions(Arrays(i))
                    Case 0
                        r = 1: c = 1
                    Case 1
                        r = UBound(Arrays(i)) - LBound(Arrays(i)) + 1
                        c = 1
                    Case 2
                        r = UBound(Arrays(i), 1) - LBound(Arrays(i), 1) + 1
                        c = UBound(Arrays(i), 2) - LBound(Arrays(i), 2) + 1
                End Select
            End If
            If r > AllR Then AllR = r
            AllC = AllC + c
        Next i

        If AllR = 0 Then
            HStack = CreateMissing()
            Exit Function
        End If

        ReDim ReturnArray(1 To AllR, 1 To AllC)

        Y0 = 1
        For i = LBound(Arrays) To UBound(Arrays)
            If Not IsMissing(Arrays(i)) Then
                Select Case NumDimensions(Arrays(i))
                    Case 0
                        r = 1: c = 1
                        ReturnArray(1, Y0) = Arrays(i)
                    Case 1
                        r = UBound(Arrays(i)) - LBound(Arrays(i)) + 1
                        c = 1
                        For j = 1 To r
                            ReturnArray(j, Y0) = Arrays(i)(j + LBound(Arrays(i)) - 1)
                        Next j
                    Case 2
                        r = UBound(Arrays(i), 1) - LBound(Arrays(i), 1) + 1
                        c = UBound(Arrays(i), 2) - LBound(Arrays(i), 2) + 1

                        For j = 1 To r
                            For k = 1 To c
                                ReturnArray(j, Y0 + k - 1) = Arrays(i)(j + LBound(Arrays(i), 1) - 1, k + LBound(Arrays(i), 2) - 1)
                            Next k
                        Next j

                End Select
                If r < AllR Then        'Pad with #NA! values
                    For j = r + 1 To AllR
                        For k = 1 To c
                            ReturnArray(j, Y0 + k - 1) = NA
                        Next k
                    Next j
                End If

                Y0 = Y0 + c
            End If
        Next i
        HStack = ReturnArray
    End If

    Exit Function
ErrHandler:
    HStack = "#HStack (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ArraysIdentical
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
' AbsTol,RelTol       : Tolerances for inexact equality comparison. See IsApprox.
' -----------------------------------------------------------------------------------------------------------------------
Function ArraysIdentical(ByVal Array1, ByVal Array2, Optional CaseSensitive As Boolean, _
    Optional PermitBaseDifference As Boolean = False, Optional ByRef WhatDiffers As String, _
    Optional AbsTol As Double, Optional RelTol As Double) As Variant
    
    Dim cN As Long
    Dim i As Long
    Dim j As Long
    Dim NumDiff As Long
    Dim NumSame As Long
    Dim rN As Long
    
    On Error GoTo ErrHandler

    'Lazy programming, flips both arrays to 2-d to avoid having to _
     write code for the 1-d case, also handles Range inputs
    Force2DArrayR Array1: Force2DArrayR Array2

    WhatDiffers = ""
    If (UBound(Array1, 1) - LBound(Array1, 1)) <> (UBound(Array2, 1) - LBound(Array2, 1)) Then
        WhatDiffers = "Row count different: " + CStr(1 + (UBound(Array1, 1) - LBound(Array1, 1))) + " vs " _
            + CStr(1 + (UBound(Array2, 1) - LBound(Array2, 1)))
        ArraysIdentical = False
    ElseIf (UBound(Array1, 2) - LBound(Array1, 2)) <> (UBound(Array2, 2) - LBound(Array2, 2)) Then
        WhatDiffers = "Column count different: " + CStr(1 + (UBound(Array1, 2) - LBound(Array1, 2))) + " vs " _
            + CStr(1 + (UBound(Array2, 2) - LBound(Array2, 2)))
        ArraysIdentical = False
    Else
        If Not PermitBaseDifference Then
            If (LBound(Array1, 1) <> LBound(Array2, 1)) Or (LBound(Array1, 2) <> LBound(Array2, 2)) Then
                WhatDiffers = "Lower bounds different"
                ArraysIdentical = False
                Exit Function
            End If
        End If
        rN = LBound(Array2, 1) - LBound(Array1, 1)
        cN = LBound(Array2, 2) - LBound(Array1, 2)
        For i = LBound(Array1, 1) To UBound(Array1, 1)
            For j = LBound(Array1, 2) To UBound(Array1, 2)
                If Not IsApprox(Array1(i, j), Array2(i + rN, j + cN), CaseSensitive, AbsTol, RelTol) Then
                    NumDiff = NumDiff + 1
                    If NumDiff = 1 Then
                        WhatDiffers = "first difference at " + CStr(i) + "," + CStr(j) + ": " + _
                            TypeName(Array1(i, j)) + " '" + CStr(Array1(i, j)) + "' vs " + _
                            TypeName(Array2(i + rN, j + cN)) + " '" + CStr(Array2(i + rN, j + cN)) + "' SafeSubtract = " & SafeSubtract(Array1(i, j), Array2(i + rN, j + cN))
                    End If
                    ArraysIdentical = False
                Else
                    NumSame = NumSame + 1
                End If
            Next j
        Next i
        If NumDiff = 0 Then
            ArraysIdentical = True
        Else
            ArraysIdentical = False
            WhatDiffers = CStr(NumDiff) + " of " + CStr(NumDiff + NumSame) + " elements differ, " + WhatDiffers
        End If

    End If

    Exit Function
ErrHandler:
    ArraysIdentical = "#ArraysIdentical (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : IsNumber
' Purpose   : Is a singleton a number?
' -----------------------------------------------------------------------------------------------------------------------
Function IsNumber(x As Variant) As Boolean
    Select Case VarType(x)
        Case vbDouble, vbInteger, vbSingle, vbLong ', vbCurrency, vbDecimal
            IsNumber = True
    End Select
End Function

Function FileCopy(SourceFile As String, TargetFile As String)
    Dim CopyOfErr As String
    Dim F As Scripting.File
    Dim FSO As Scripting.FileSystemObject
    On Error GoTo ErrHandler
    Set FSO = New FileSystemObject
    Set F = FSO.GetFile(SourceFile)
    F.Copy TargetFile, True
    FileCopy = TargetFile
    Set FSO = Nothing: Set F = Nothing
    Exit Function
ErrHandler:
    CopyOfErr = Err.Description
    Set FSO = Nothing: Set F = Nothing
    Throw "#" + CopyOfErr + "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sStringBetweenStrings
' Purpose   : The function returns the substring of the input TheString which lies between LeftString
'             and RightString.
' Arguments
' TheString : The input string to be searched.
' LeftString: The returned string will start immediately after the first occurrence of LeftString in
'             TheString. If LeftString is not found or is the null string or missing, then
'             the return will start at the first character of TheString.
' RightString: The return will stop immediately before the first subsequent occurrence of RightString. If
'             such occurrrence is not found or if RightString is the null string or
'             missing, then the return will stop at the last character of TheString.
' IncludeLeftString: If TRUE, then if LeftString appears in TheString, the return will include LeftString. This
'             argument is optional and defaults to FALSE.
' IncludeRightString: If TRUE, then if RightString appears in TheString (and appears after the first occurance
'             of LeftString) then the return will include RightString. This argument is
'             optional and defaults to FALSE.
' -----------------------------------------------------------------------------------------------------------------------
Function StringBetweenStrings(TheString, LeftString, RightString, Optional IncludeLeftString As Boolean, Optional IncludeRightString As Boolean)
    Dim MatchPoint1 As Long        ' the position of the first character to return
    Dim MatchPoint2 As Long        ' the position of the last character to return
    Dim FoundLeft As Boolean
    Dim FoundRight As Boolean
    
    On Error GoTo ErrHandler
    
    If VarType(TheString) <> vbString Or VarType(LeftString) <> vbString Or VarType(RightString) <> vbString Then Throw "Inputs must be strings"
    If LeftString = vbNullString Then
        MatchPoint1 = 0
    Else
        MatchPoint1 = InStr(1, TheString, LeftString, vbTextCompare)
    End If

    If MatchPoint1 = 0 Then
        FoundLeft = False
        MatchPoint1 = 1
    Else
        FoundLeft = True
    End If

    If RightString = vbNullString Then
        MatchPoint2 = 0
    ElseIf FoundLeft Then
        MatchPoint2 = InStr(MatchPoint1 + Len(LeftString), TheString, RightString, vbTextCompare)
    Else
        MatchPoint2 = InStr(1, TheString, RightString, vbTextCompare)
    End If

    If MatchPoint2 = 0 Then
        FoundRight = False
        MatchPoint2 = Len(TheString)
    Else
        FoundRight = True
        MatchPoint2 = MatchPoint2 - 1
    End If

    If Not IncludeLeftString Then
        If FoundLeft Then
            MatchPoint1 = MatchPoint1 + Len(LeftString)
        End If
    End If

    If IncludeRightString Then
        If FoundRight Then
            MatchPoint2 = MatchPoint2 + Len(RightString)
        End If
    End If

    StringBetweenStrings = Mid$(TheString, MatchPoint1, MatchPoint2 - MatchPoint1 + 1)

    Exit Function
ErrHandler:
    StringBetweenStrings = "#StringBetweenStrings (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function