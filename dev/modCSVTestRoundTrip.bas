Attribute VB_Name = "modCSVTestRoundTrip"

' VBA-CSV

' Copyright (C) 2021 - Philip Swannell (https://github.com/PGS62/VBA-CSV )
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/VBA-CSV#readme

Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RoundTripTest
' Purpose    : Tests multiple times that CSVRead correctly round-trips data previously saved to disk by CSVWrite.
'              Tests include:
'           *  Embedded line feeds in quoted strings.
'           *  Files with Windows, Unix or (old) Mac line endings.
'           *  Both unicode and ascii files.
'           *  That the delimiter is automatically detected by CSVRead (reliable only if files have all strings quoted).
'           *  That unicode vs ascii is automatically detected.
'           *  That line endings are automatically detected.
'Results are printed to the VBA immediate window, if a difference is detected
' -----------------------------------------------------------------------------------------------------------------------
Public Sub RoundTripTest()

    Dim AllowLineFeed As Variant
    Dim Data As Variant
    Dim DateFormat As Variant
    Dim Delimiter As Variant
    Dim EOL As String
    Dim ExtraInfo As String
    Dim Folder As String
    Dim k As Long
    Dim NCols As Variant
    Dim NRows As Variant
    Dim NumFailed As Long
    Dim NumPassed As Long
    Dim NumTests As Long
    Dim OS As Variant
    Dim Prompt As String
    Dim Unicode As Variant
    Dim WhatDiffers As String
    Const Title = "VBA-CSV Round Trip Tests"
    
    On Error GoTo ErrHandler
    
    Folder = Environ("Temp") & "\VBA-CSV\RoundTripTests"

    Prompt = "Run Round Trip Tests?" + vbLf + vbLf + _
        "Note this will generate 1,620 files in folder" + vbLf + _
        Environ("Temp") & "\VBA-CSV\RoundTripTests"

    If MsgBox(Prompt, vbOKCancel + vbQuestion, Title) <> vbOK Then Exit Sub

    ThrowIfError CreatePath(Folder)

    For Each OS In Array("Windows", "Unix", "Mac")
        EOL = IIf(OS = "Windows", vbCrLf, IIf(OS = "Unix", vbLf, vbCr))
    
        For Each Unicode In Array(True, False)
            For Each Delimiter In Array(",", "::::")
                For Each NRows In Array(1, 2, 20)
                    For Each NCols In Array(1, 2, 10)
        
                        'For Variants we need to vary AllowLineFeed and DateFormat
                        For Each AllowLineFeed In Array(True, False)
                            For Each DateFormat In Array("mmm-dd-yyyy", "dd-mmm-yyyy", "yyyy-mm-dd")
                                Data = RandomVariants(CLng(NRows), CLng(NCols), CBool(AllowLineFeed), CBool(Unicode), EOL)
                                NumTests = NumTests + 1
                                ExtraInfo = "Test " & CStr(NumTests) & " " & "RandomVariants" & IIf(AllowLineFeed, "WithLineFeed", "")
                                RoundTripTestCore Folder, CStr(OS), Data, CStr(DateFormat), CBool(Unicode), CStr(OS), CStr(Delimiter), ExtraInfo, WhatDiffers, NumPassed, NumFailed
                            Next DateFormat
                        Next AllowLineFeed

                        'For Dates, we need to vary DateFormat
                        For Each DateFormat In Array("mmm-dd-yyyy", "dd-mmm-yyyy", "yyyy-mm-dd")
                            Data = RandomDates(CLng(NRows), CLng(NCols))
                            NumTests = NumTests + 1
                            ExtraInfo = "Test " & CStr(NumTests) & " " & "RandomDates"
                            RoundTripTestCore Folder, CStr(OS), Data, CStr(DateFormat), CBool(Unicode), CStr(OS), CStr(Delimiter), ExtraInfo, WhatDiffers, NumPassed, NumFailed
                        Next DateFormat

                        'For Strings, we need to vary AllowLineFeed
                        For Each AllowLineFeed In Array(True, False)
                            Data = RandomStrings(CLng(NRows), CLng(NCols), CBool(Unicode), CBool(AllowLineFeed), EOL)
                            NumTests = NumTests + 1
                            ExtraInfo = "Test " & CStr(NumTests) & " " & IIf(AllowLineFeed, "RandomStringsWithLineFeeds", "RandomStrings")
                            RoundTripTestCore Folder, CStr(OS), Data, CStr(DateFormat), CBool(Unicode), CStr(OS), CStr(Delimiter), ExtraInfo, WhatDiffers, NumPassed, NumFailed
                        Next AllowLineFeed

                        For k = 1 To 4
                            NumTests = NumTests + 1
                            If k = 1 Then
                                Data = RandomBooleans(CLng(NRows), CLng(NCols))
                                ExtraInfo = "Test " & CStr(NumTests) & " " & "RandomBooleans"
                            ElseIf k = 2 Then
                                Data = RandomDoubles(CLng(NRows), CLng(NCols))
                                ExtraInfo = "Test " & CStr(NumTests) & " " & "RandomDoubles"
                            ElseIf k = 3 Then
                                Data = RandomErrorValues(CLng(NRows), CLng(NCols))
                                ExtraInfo = "Test " & CStr(NumTests) & " " & "RandomErrorValues"
                            ElseIf k = 4 Then
                                Data = RandomLongs(CLng(NRows), CLng(NCols))
                                ExtraInfo = "Test " & CStr(NumTests) & " " & "RandomLongs"
                            End If
                            RoundTripTestCore Folder, CStr(OS), Data, CStr(DateFormat), CBool(Unicode), _
                                CStr(OS), CStr(Delimiter), ExtraInfo, WhatDiffers, NumPassed, NumFailed
                        Next k
                        DoEvents 'Kick Immediate window back to life?
                    Next NCols
                Next NRows
            Next Delimiter
        Next Unicode
    Next OS
    Debug.Print "Finished. NumPassed = " + Format(NumPassed, "###,##0") & " NumFailed = " & Format(NumFailed, "###,##0")
    If NumFailed = 0 Then
        Prompt = "Finished, all " + Format(NumPassed, "###,##0") + " tests passed"
        MsgBox Prompt, vbOKOnly + vbInformation, Title
    Else
        Prompt = "Finished, " + Format(NumPassed, "###,##0") + " tests passed, and " & Format(NumFailed, "###,##0") + " tests failed. See VBA immediate window for details."
        MsgBox Prompt, vbOKOnly + vbCritical, Title
    End If

    Exit Sub
ErrHandler:
    MsgBox "#RoundTripTest: " & Err.Description & "!", vbCritical
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RoundTripTestCore
' Purpose    : Test for "round trip" between functions CSVRead and CSVWrite. We write data to a file, then read it back
'              and test that the read-back data is identical to the starting data. For round-tripping to work we must write
'              with QuoteAllStrings being TRUE and read back with ShowMissingsAs being Empty (to be able to distinguish
'              Empty and null string. Also method RandomDoubles only generates doubles that have exact representation as
'              strings (avoid errors of order 10E-15).
' -----------------------------------------------------------------------------------------------------------------------
Private Sub RoundTripTestCore(Folder As String, OS As String, ByVal Data As Variant, DateFormat As String, _
    Unicode As Boolean, EOL As String, Delimiter As String, ExtraInfo As String, ByRef WhatDiffers As String, _
    ByRef NumPassed As Long, ByRef NumFailed As Long)

    Const ConvertTypes = "NDBE" 'must use this for round-tripping to work.
    Const PermitBaseDifference = True
    Dim DataReadBack
    Dim FileName As String
    Dim NC As Long
    Dim NR As Long
    Dim NumDone As Long
    
    On Error GoTo ErrHandler
    
    WhatDiffers = ""

    NR = NRows(Data)
    NC = NCols(Data)

    FileName = NameThatFile(Folder, OS, NR, NC, ExtraInfo, CBool(Unicode), False)

    ThrowIfError CSVWrite(Data, FileName, True, DateFormat, , Delimiter, Unicode, EOL)

    'The Call to CSVRead has to infer both Encoding and EOL
    DataReadBack = CSVRead(FileName, ConvertTypes, Delimiter, DateFormat:=DateFormat, ShowMissingsAs:=Empty)
    
    If ArraysIdentical(Data, DataReadBack, True, PermitBaseDifference, WhatDiffers) Then
        NumPassed = NumPassed + 1
    Else
        Debug.Print String(100, "=")
        Debug.Print "Round trip failed for: '" & FileName & "'"
        Debug.Print WhatDiffers
        NumFailed = NumFailed + 1
    End If

    NumDone = NumPassed + NumFailed
    If NumDone Mod 50 = 0 Then Debug.Print Format(NumDone, "###,##0")
    
    Exit Sub
ErrHandler:
    Throw "#RoundTripTestCore: " & Err.Description & "!"
End Sub

Private Function RandomString(AllowLineFeed As Boolean, Unicode As Boolean, EOL As String)

    Const maxlen = 20
    Dim i As Long
    Dim length As Long
    Dim Res As String
    
    On Error GoTo ErrHandler
    
    length = CLng(1 + Rnd() * maxlen)
    Res = String(length, " ")

    For i = 1 To length
        If Unicode Then
            Mid(Res, i, 1) = ChrW(33 + Rnd() * 370)
        Else
            Mid(Res, i, 1) = Chr(34 + Rnd() * 88)
        End If

        If Not AllowLineFeed Then
            If Mid(Res, i, 1) = vbLf Or Mid(Res, i, 1) = vbCr Then
                Mid(Res, i, 1) = " "
            End If
        End If
    Next

    If AllowLineFeed Then
        If length > 5 Then
            If Rnd() < 0.2 Then
                Mid(Res, length / 2, Len(EOL)) = EOL
            End If
        End If
    End If
    
    RandomString = Res

    Exit Function
ErrHandler:
    Throw "#RandomString: " & Err.Description & "!"
End Function

Private Function RandomStrings(NumRows As Long, NumCols As Long, Unicode As Boolean, AllowLineFeed As Boolean, EOL As String)

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

Private Function RandomLongs(NumRows As Long, NumCols As Long)

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

Private Function RandomDoubles(NumRows As Long, NumCols As Long)

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

Private Function RandomBooleans(NumRows As Long, NumCols As Long)
    
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

Private Function RandomDates(NumRows As Long, NumCols As Long)
    
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

Private Function RandomErrorValues(NumRows As Long, NumCols As Long)
    
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

    Const Err_Invalid = " must be one of ""Windows"", ""Unix"" or ""Mac"", or the associated end of line characters."

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

'Public because called from worksheet "Demo"
Public Function RandomVariants(NRows As Long, NCols As Long, AllowLineFeed As Boolean, Unicode As Boolean, ByVal EOL As String)

    Application.Volatile

    Const DateFormat = "yyyy-mmm-dd"
    Dim i As Long
    Dim j As Long
    Dim Res() As Variant

    On Error GoTo ErrHandler

    EOL = OStoEOL(EOL, "EOL")
    ReDim Res(1 To NRows, 1 To NCols)

    For i = 1 To NRows
        For j = 1 To NCols
            Res(i, j) = RandomVariant(DateFormat, AllowLineFeed, Unicode, EOL)
        Next j
    Next i
    If AllowLineFeed Then
        i = 0.5 + Rnd() * NRows
        j = 0.5 + Rnd() * NCols
        Res(i, j) = "Here's a carriage return (ascii 13):" & vbCr & "and here's a line feed (ascii 10):" & vbLf & "and here's both together:" & vbCrLf
    End If

    RandomVariants = Res

    Exit Function
ErrHandler:
    Throw "#RandomVariants: " & Err.Description & "!"
End Function

