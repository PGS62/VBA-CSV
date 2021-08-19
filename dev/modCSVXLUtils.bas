Attribute VB_Name = "modCSVXLUtils"
Option Explicit
'Functions that, inaddition to CSVRead and CSVWrite, are called from the worksheets of this workbook


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

Function RawFileContents(FileName As String)
    Dim FSO As New FileSystemObject, F As Scripting.File, T As Scripting.TextStream
    On Error GoTo ErrHandler
    Set F = FSO.GetFile(FileName)
    Set T = F.OpenAsTextStream()
    RawFileContents = T.ReadAll
    T.Close

    Exit Function
ErrHandler:
    Throw "#RawFileContents (line " & CStr(Erl) + "): " & Err.Description & "!"
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

