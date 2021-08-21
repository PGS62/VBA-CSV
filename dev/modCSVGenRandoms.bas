Attribute VB_Name = "modCSVGenRandoms"
Option Explicit

Private Function RandomString(AllowLineFeed As Boolean, Unicode As Boolean, EOL As String)

    Const maxlen = 20
    Dim i As Long
    Dim length As Long
    Dim res As String
    
    On Error GoTo ErrHandler
    
    length = CLng(1 + Rnd() * maxlen)
    res = String(length, " ")

    For i = 1 To length
        If Unicode Then
            Mid(res, i, 1) = ChrW(33 + Rnd() * 370)
        Else
            Mid(res, i, 1) = Chr(34 + Rnd() * 88)
        End If

        If Not AllowLineFeed Then
            If Mid(res, i, 1) = vbLf Or Mid(res, i, 1) = vbCr Then
                Mid(res, i, 1) = " "
            End If
        End If
    Next

    If AllowLineFeed Then
        If length > 5 Then
            If Rnd() < 0.2 Then
                Mid(res, length / 2, Len(EOL)) = EOL
            End If
        End If
    End If
    
    RandomString = res

    Exit Function
ErrHandler:
    Throw "#RandomString: " & Err.Description & "!"
End Function

Function RandomStrings(NumRows As Long, NumCols As Long, Unicode As Boolean, AllowLineFeed As Boolean, EOL As String)

    Dim i As Long
    Dim j As Long
    Dim Result() As String
    
    On Error GoTo ErrHandler
    
    ReDim Result(1 To NumRows, 1 To NumCols)
    For i = 1 To NumRows
        For j = 1 To NumCols
            Result(i, j) = RandomString(AllowLineFeed, Unicode, EOL)
        Next j
    Next i
    If AllowLineFeed Then
        i = 0.5 + Rnd() * NumRows
        j = 0.5 + Rnd() * NumCols
        Result(i, j) = "Here's a carriage return (ascii 13):" & vbCr & "and here's a line feed (ascii 10):" & vbLf & "and here's both together:" & vbCrLf
    End If
    RandomStrings = Result
    
    Exit Function
ErrHandler:
    Throw "#RandomStrings: " & Err.Description & "!"
End Function

Private Function RandomLong() As Long
    RandomLong = CLng((Rnd() - 0.5) * 2000000)
End Function

Function RandomLongs(NumRows As Long, NumCols As Long)

    Dim i As Long
    Dim j As Long
    Dim Result() As Long
    
    On Error GoTo ErrHandler
    
    ReDim Result(1 To NumRows, 1 To NumCols)
    For i = 1 To NumRows
        For j = 1 To NumCols
            Result(i, j) = RandomLong()
        Next j
    Next i
    RandomLongs = Result
    Exit Function
ErrHandler:
    Throw "#RandomLongs: " & Err.Description & "!"
End Function

Private Function RandomDouble() As Double
    On Error GoTo ErrHandler
    'Trick! - Generate a Double that has an exact representation as String. Avoids rounding errors when we write to disk and read back
    RandomDouble = CDbl(CStr((Rnd() - 0.5) * 2 * 10 ^ ((Rnd() - 0.5) * 20)))
    Exit Function
ErrHandler:
    Throw "#RandomDouble: " & Err.Description & "!"
End Function

Function RandomDoubles(NumRows As Long, NumCols As Long)

    Dim i As Long
    Dim j As Long
    Dim Result() As Double
    
    On Error GoTo ErrHandler
    
    ReDim Result(1 To NumRows, 1 To NumCols)
    For i = 1 To NumRows
        For j = 1 To NumCols
            Result(i, j) = RandomDouble()
        Next j
    Next i
    RandomDoubles = Result
    Exit Function
    
ErrHandler:
    Throw "#RandomDoubles: " & Err.Description & "!"
End Function

Private Function RandomBoolean() As Boolean
    RandomBoolean = Rnd() < 0.5
End Function

Function RandomBooleans(NumRows As Long, NumCols As Long)
    
    Dim i As Long
    Dim j As Long
    Dim Result() As Boolean
    
    On Error GoTo ErrHandler
    
    ReDim Result(1 To NumRows, 1 To NumCols)
    For i = 1 To NumRows
        For j = 1 To NumCols
            Result(i, j) = RandomBoolean()
        Next j
    Next i
    RandomBooleans = Result
    Exit Function
ErrHandler:
    Throw "#RandomBooleans: " & Err.Description & "!"
End Function

Private Function RandomDate()
    On Error GoTo ErrHandler
    RandomDate = CDate(CLng(25569 + Rnd() * 36525)) 'Date in range 1 Jan 1970 to 1 Jan 2070
    Exit Function
ErrHandler:
    Throw "#RandomDate: " & Err.Description & "!"
End Function

Function RandomDates(NumRows As Long, NumCols As Long)
    
    Dim i As Long
    Dim j As Long
    Dim Result() As Date
    
    On Error GoTo ErrHandler
    
    ReDim Result(1 To NumRows, 1 To NumCols)
    For i = 1 To NumRows
        For j = 1 To NumCols
            Result(i, j) = RandomDate()
        Next j
    Next i
    RandomDates = Result
    Exit Function
    
ErrHandler:
    Throw "#RandomDates: " & Err.Description & "!"
End Function

Private Function RandomErrorValue()
    Dim N As Long
    On Error GoTo ErrHandler
    N = CLng(0.5 + Rnd() * 14)
    RandomErrorValue = CVErr(Choose(N, xlErrBlocked, xlErrCalc, xlErrConnect, xlErrDiv0, xlErrField, xlErrGettingData, _
        xlErrNA, xlErrName, xlErrNull, xlErrNum, xlErrRef, xlErrSpill, xlErrUnknown, xlErrValue))
    Exit Function
ErrHandler:
    Throw "#RandomErrorValue: " & Err.Description & "!"
End Function

Function RandomErrorValues(NumRows As Long, NumCols As Long)
    
    Dim i As Long
    Dim j As Long
    Dim Result() As Variant
    
    On Error GoTo ErrHandler
    ReDim Result(1 To NumRows, 1 To NumCols)
    For i = 1 To NumRows
        For j = 1 To NumCols
            Result(i, j) = RandomErrorValue()
        Next j
    Next i
    RandomErrorValues = Result
    Exit Function
    
ErrHandler:
    Throw "#RandomErrorValues: " & Err.Description & "!"
End Function

Private Function RandomVariant(DateFormat As String, AllowLineFeed As Boolean, Unicode As Boolean, EOL As String)

    Dim N As Long
    Const NUMTYPES = 11

    On Error GoTo ErrHandler
    N = CLng(0.5 + NUMTYPES * Rnd())

    Select Case N
        Case 1
            RandomVariant = RandomBoolean()
        Case 2
            RandomVariant = RandomLong()
        Case 3
            RandomVariant = RandomDouble()
        Case 4
            RandomVariant = RandomString(AllowLineFeed, Unicode, EOL)
        Case 5
            RandomVariant = RandomDate()
        Case 6
            RandomVariant = vbNullString
        Case 7
            'String that looks like a number
            RandomVariant = CStr(RandomDouble())
        Case 8
            'String that looks like a date
            RandomVariant = Format(CLng(RandomDate()), DateFormat)
        Case 9
            'String that looks like Boolean
            RandomVariant = UCase(CStr(RandomBoolean()))
        Case 10
            RandomVariant = Empty
        Case 11
            RandomVariant = RandomErrorValue()
    End Select

    Exit Function
ErrHandler:
    Throw "#RandomVariant: " & Err.Description & "!"
End Function

'Copy of identical function in modCVS so that copy there can be Private
Private Function OStoEOL(OS As String, ArgName As String) As String

    Const Err_Invalid = " must be one of ""Windows"", ""Unix"" or ""Mac"", or the associented end of line characters."

    Select Case LCase(OS)
        Case "windows", vbCrLf
            OStoEOL = vbCrLf
        Case "unix", vbLf
            OStoEOL = vbLf
        Case "mac", vbCr
            OStoEOL = vbCr
        Case Else
            Throw ArgName + Err_Invalid
    End Select
    
End Function

Function RandomVariants(NRows As Long, NCols As Long, AllowLineFeed As Boolean, Unicode As Boolean, ByVal EOL As String)

    Application.Volatile

    Const DateFormat = "yyyy-mmm-dd"
    Dim i As Long
    Dim j As Long
    Dim res() As Variant

    On Error GoTo ErrHandler

    EOL = OStoEOL(EOL, "EOL")
    ReDim res(1 To NRows, 1 To NCols)

    For i = 1 To NRows
        For j = 1 To NCols
            res(i, j) = RandomVariant(DateFormat, AllowLineFeed, Unicode, EOL)
        Next j
    Next i
    If AllowLineFeed Then
        i = 0.5 + Rnd() * NRows
        j = 0.5 + Rnd() * NCols
        res(i, j) = "Here's a carriage return (ascii 13):" & vbCr & "and here's a line feed (ascii 10):" & vbLf & "and here's both together:" & vbCrLf
    End If

    RandomVariants = res

    Exit Function
ErrHandler:
    Throw "#RandomVariants: " & Err.Description & "!"
End Function
