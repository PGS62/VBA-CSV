Attribute VB_Name = "modCSVISO8601"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SpeedTestISO8601
' Purpose    : Testing speed of CastISO8601
'Example output:
'Running SpeedTestISO8601 2021-08-24T18:59:30
'Calls per second = 2,650,189              strIn = "Foo" Check = 1900-01-00T00:00:00.000
'Calls per second = 2,660,280              strIn = "Foo" Check = 1900-01-00T00:00:00.000
'Calls per second = 2,681,348              strIn = "Foo" Check = 1900-01-00T00:00:00.000
'Calls per second = 2,229,503              strIn = "xxxxxxxxxxxx"      Check = 1900-01-00T00:00:00.000
'Calls per second = 2,144,496              strIn = "xxxxxxxxxxxx"      Check = 1900-01-00T00:00:00.000
'Calls per second = 2,135,019              strIn = "xxxxxxxxxxxx"      Check = 1900-01-00T00:00:00.000
'Calls per second = 1,581,522              strIn = "xxxx-xxxxxxx"      Check = 1900-01-00T00:00:00.000
'Calls per second = 1,574,575              strIn = "xxxx-xxxxxxx"      Check = 1900-01-00T00:00:00.000
'Calls per second = 1,594,879              strIn = "xxxx-xxxxxxx"      Check = 1900-01-00T00:00:00.000
'Calls per second = 632,374  strIn = "2021-08-24T15:18:01.123+05:0x"   Check = 1900-01-00T00:00:00.000
'Calls per second = 633,033  strIn = "2021-08-24T15:18:01.123+05:0x"   Check = 1900-01-00T00:00:00.000
'Calls per second = 643,524  strIn = "2021-08-24T15:18:01.123+05:0x"   Check = 1900-01-00T00:00:00.000
'Calls per second = 469,506  strIn = "2021-08-23"        Check = 2021-08-23T00:00:00.000
'Calls per second = 471,985  strIn = "2021-08-23"        Check = 2021-08-23T00:00:00.000
'Calls per second = 465,819  strIn = "2021-08-23"        Check = 2021-08-23T00:00:00.000
'Calls per second = 368,133  strIn = "2021-08-24T15:18:01"             Check = 2021-08-24T15:18:01.000
'Calls per second = 367,303  strIn = "2021-08-24T15:18:01"             Check = 2021-08-24T15:18:01.000
'Calls per second = 350,648  strIn = "2021-08-24T15:18:01"             Check = 2021-08-24T15:18:01.000
'Calls per second = 230,480  strIn = "2021-08-23T08:47:21.123"         Check = 2021-08-23T08:47:21.000
'Calls per second = 225,836  strIn = "2021-08-23T08:47:21.123"         Check = 2021-08-23T08:47:21.000
'Calls per second = 246,625  strIn = "2021-08-23T08:47:21.123"         Check = 2021-08-23T08:47:21.000
'Calls per second = 235,492  strIn = "2021-08-24T15:18:01+05:00"       Check = 2021-08-24T10:18:01.000
'Calls per second = 230,097  strIn = "2021-08-24T15:18:01+05:00"       Check = 2021-08-24T10:18:01.000
'Calls per second = 231,845  strIn = "2021-08-24T15:18:01+05:00"       Check = 2021-08-24T10:18:01.000
'Calls per second = 215,734  strIn = "2021-08-24T15:18:01.123+05:00"   Check = 2021-08-24T10:18:01.000
'Calls per second = 208,329  strIn = "2021-08-24T15:18:01.123+05:00"   Check = 2021-08-24T10:18:01.000
'Calls per second = 211,625  strIn = "2021-08-24T15:18:01.123+05:00"   Check = 2021-08-24T10:18:01.000

' -----------------------------------------------------------------------------------------------------------------------
Function SpeedTestISO8601()

    Const N = 1000
    Dim Converted As Boolean
    Dim dtOut As Date
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim strIn As String
    Dim SysDateOrder As Long
    Dim t1 As Double
    Dim t2 As Double
    Dim res

    SysDateOrder = Application.International(xlDateOrder)

    Debug.Print "Running SpeedTestISO8601 " + Format(Now(), "yyyy-mm-ddThh:mm:ss")
    For k = 1 To 9
        For j = 1 To 3
            dtOut = 0
            Select Case k
                Case 1
                    strIn = "Foo" ' less than 10 in length
                Case 2
                    strIn = "xxxxxxxxxxxx" '5th character not "-"
                Case 3
                    strIn = "xxxx-xxxxxxx" 'rejected by RegEx
                Case 4
                    strIn = "2021-08-24T15:18:01.123+05:0x" ' rejected by regex
                Case 5
                    strIn = "2021-08-23"
                Case 6
                    strIn = "2021-08-24T15:18:01"
                Case 7
                    strIn = "2021-08-23T08:47:21.123"
                Case 8
                    strIn = "2021-08-24T15:18:01+05:00"
                Case 9
                    strIn = "2021-08-24T15:18:01.123+05:00"
            End Select

            t1 = sElapsedTime()
            For i = 1 To N
                Call CastISO8601(strIn, dtOut, Converted, SysDateOrder)
            Next i
            t2 = sElapsedTime
            Debug.Print "Calls per second = " & Format(N / (t2 - t1), "###,###"), "strIn = """ & strIn & """", "Check = " & Application.WorksheetFunction.text(dtOut, "yyyy-mm-ddThh:mm:ss.000")
            DoEvents 'kick Immediate window to life
        Next j
    Next k

End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseISO8601
' Purpose    : Test harness for calling from spreadsheets
' Parameters :
'  strIn:
' -----------------------------------------------------------------------------------------------------------------------
Function ParseISO8601(strIn As String)
          Dim dtOut As Date
          Dim Converted As Boolean

1         On Error GoTo ErrHandler
2         CastISO8601 strIn, dtOut, Converted, Application.International(xlDateOrder)

3         If Converted Then
4             ParseISO8601 = dtOut
5         Else
6             ParseISO8601 = "#Not recognised as ISO8601 date!"
7         End If
8         Exit Function
ErrHandler:
9         ParseISO8601 = "#ParseISO8601 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CastToTime
' Purpose    : Cast strings that represent a time to a date, no handling of TimeZone.
'Example inputs:
'                13:30
'                13:30:10
'                13:30:10.123
' Parameters :
'  strIn    :
'  dtOut    :
'  Converted:
' -----------------------------------------------------------------------------------------------------------------------
Sub CastToTime(strIn As String, dtOut As Date, ByRef Converted As Boolean)

    Static rx As VBScript_RegExp_55.RegExp

    Dim L As Long

    On Error GoTo ErrHandler
    If rx Is Nothing Then
        Set rx = New RegExp
        With rx
            .IgnoreCase = False
            .Pattern = "^[0-9][0-9]:[0-9][0-9](:[0-9][0-9](\.[0-9]+)?)?$"
            .Global = False        'Find first match only
        End With
    End If

    L = Len(strIn)

    If L < 5 Then Exit Sub
    If Not rx.Test(strIn) Then Exit Sub
    If L <= 8 Then
        dtOut = CDate(strIn)
        Converted = True
    Else
        dtOut = CDate(Left(strIn, 8)) + CDbl(Mid(strIn, 9)) / 86400
        Converted = True
    End If
    Exit Sub
ErrHandler:
    'Do nothing, was not a valid time (e.g. h,m or s out of range)
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CastISO8601
' Purpose    : Convert ISO8601 formatted datestrings to UTC date. Handles the following formats:
'Format                        Example                        Comment
'yyyy-mm-dd                    2021-08-23                     Date only
'yyyy-mm-ddThh:mm:ss           2021-08-23T08:47:21            Date and time, no time zone given
'yyyy-mm-ddThh:mm:ssZ          2021-08-23T08:47:21Z           Date and time, UTC
'yyyy-mm-ddThh:mm:ss.000       2021-08-23T08:47:20.920        Date and time with fractional second, no time zone given
'yyyy-mm-ddThh:mm:ss.000Z      2021-08-23T08:47:20.920Z       Date and time with fractional second, UTC
'yyyy-mm-ddThh:mm:ss+hh:mm     2021-08-23T08:47:21+05:00      Date and time, time zone given
'yyyy-mm-ddThh:mm:ss.000+hh:mm 2021-08-23T08:47:20.920+05:00  Date and time with fractional second, time zone given

'https://xkcd.com/1179/

' Parameters :
'  StrIn       : The string to be converted
'  DtOut       : The date that the string converts to. If strIn specifies a time zone then the output is UTC time.
'  Converted   : Did the function convert (true) or reject as not a correctly formatted date (false)
'  SysDateOrder: The Windows system date order. 0 = M-D-Y, 1= D-M-Y, 2 = Y-M-D
' -----------------------------------------------------------------------------------------------------------------------
Sub CastISO8601(ByVal strIn As String, dtOut As Date, ByRef Converted As Boolean, SysDateOrder As Long)

    Dim D As String
    Dim HaveMilliPart As Boolean
    Dim L As Long
    Dim LocalTime As Double
    Dim Mask As String
    Dim MilliPart As Double
    Dim MinusPos As Long
    Dim Mo As String
    Dim PlusPos As Long
    Dim Sign
    Dim Y As String
    Dim ZAtEnd As Boolean
    
    Static rx As VBScript_RegExp_55.RegExp
    Static rxExists As Boolean

    On Error GoTo ErrHandler

    'We use CDbl and CDate but these have some not-so-helpful behaviours, such as:
    'CDbl("4,2") = 42
    'CDate("12:30:1A") = "00:30:01"
    'CDate("12:30:1P") = "12:30:01"
    'CDbl("1£") = 1
     
    'The regex below will reject all strings that do not match the seven patterns given in the header to this function, _
     but do not reject some out-of range input such as "2021-13-32". Instead, out of range inputs will be rejected by the _
     calls to CDate.

    If Not rxExists Then
        Set rx = New RegExp
        With rx
            .IgnoreCase = False
            .Pattern = "^[0-9][0-9][0-9][0-9]\-[[0-1][0-9]\-[0-3][0-9](T[0-2][0-9]:[0-5][0-9]:[0-5][0-9](\.[0-9]+)?((Z|((\+|\-)[0-2][0-9]:[0-5][0-9])))?)?$"
            .Global = False        'Find first match only
        End With
        rxExists = True
    End If

    Converted = False

    L = Len(strIn)

    If L < 10 Then Exit Sub
    If Mid(strIn, 5, 1) <> "-" Then Exit Sub
    If Not rx.Test(strIn) Then Exit Sub
    Y = Left$(strIn, 4)
    Mo = Mid$(strIn, 6, 2)
    D = Mid$(strIn, 9, 2)
    
    If L = 10 Then 'e.g. 2021-08-23, date only
        Select Case SysDateOrder
            Case 2 ' Y-M-D
                dtOut = CDate(Left(strIn, 10))
            Case 0 'M-D-Y
                Mask = String(10, "-")
                Mid(Mask, 1, 2) = Mo
                Mid(Mask, 4, 2) = D
                Mid(Mask, 7, 4) = Y
                dtOut = CDate(Mask)
            Case 1 'D-M-Y
                Mask = String(10, "-")
                Mid(Mask, 1, 2) = D
                Mid(Mask, 4, 2) = Mo
                Mid(Mask, 7, 4) = Y
                dtOut = CDate(Mask)
        End Select
        Converted = True
        Exit Sub
    End If
    
    If L = 19 Then 'Example: "2021-08-23T08:47:21" i.e. "Local time" with no time zone indicated
        Select Case SysDateOrder
            Case 2 ' Y-M-D
                dtOut = CDate(strIn)
            Case 0 'M-D-Y
                'Change strIn in-place, we don't need to edit the time part
                Mid(strIn, 1, 2) = Mo
                Mid(strIn, 3, 1) = "-"
                Mid(strIn, 4, 2) = D
                Mid(strIn, 6, 1) = "-"
                Mid(strIn, 7, 4) = Y
                Mid(strIn, 11, 1) = " "
                dtOut = CDate(strIn)
            Case 1 'D-M-Y
                Mid(strIn, 1, 2) = D
                Mid(strIn, 3, 1) = "-"
                Mid(strIn, 4, 2) = Mo
                Mid(strIn, 6, 1) = "-"
                Mid(strIn, 7, 4) = Y
                Mid(strIn, 11, 1) = " "
                dtOut = CDate(strIn)
        End Select
        Converted = True
        Exit Sub
    End If

    If Right(strIn, 1) = "Z" Then
        Sign = 0
        ZAtEnd = True
    Else
        PlusPos = InStr(20, strIn, "+")
        If PlusPos > 0 Then
            Sign = 1
        Else
            MinusPos = InStr(20, strIn, "-")
            If MinusPos > 0 Then
                Sign = -1
            End If
        End If
    End If

    If Mid(strIn, 20, 1) = "." Then 'Have fraction of a second
        Select Case Sign
            Case 0
                'Example: "2021-08-23T08:47:20.920Z"
                MilliPart = CDbl(Mid(strIn, 20, IIf(ZAtEnd, L - 20, L - 19)))
            Case 1
                'Example: "2021-08-23T08:47:20.920+05:00"
                MilliPart = CDbl(Mid(strIn, 20, PlusPos - 20))
            Case -1
                'Example: "2021-08-23T08:47:20.920-05:00"
                MilliPart = CDbl(Mid(strIn, 20, MinusPos - 20))
        End Select
    End If
    
    Select Case SysDateOrder
        Case 2 ' Y-M-D
            LocalTime = CDate(Left(strIn, 19)) + MilliPart / 86400
        Case 0 'M-D-Y
            Mask = String(19, " ")
            Mid(Mask, 1, 2) = Mo
            Mid(Mask, 3, 1) = "-"
            Mid(Mask, 4, 2) = D
            Mid(Mask, 6, 1) = "-"
            Mid(Mask, 7, 4) = Y
            Mid(Mask, 11, 1) = " "
            Mid(Mask, 12, 8) = Mid(strIn, 12, 8)
            LocalTime = CDate(Mask) + MilliPart / 86400
        Case 1 'D-M-Y
            Mask = String(19, " ")
            Mid(Mask, 1, 2) = D
            Mid(Mask, 3, 1) = "-"
            Mid(Mask, 4, 2) = Mo
            Mid(Mask, 6, 1) = "-"
            Mid(Mask, 7, 4) = Y
            Mid(Mask, 11, 1) = " "
            Mid(Mask, 12, 8) = Mid(strIn, 12, 8)
            LocalTime = CDate(Mask) + MilliPart / 86400
    End Select

    Dim Adjust As Date
    Select Case Sign
        Case 0
            dtOut = LocalTime
            Converted = True
            Exit Sub
        Case 1
            If L <> PlusPos + 5 Then Exit Sub
            Adjust = CDate(Right(strIn, 5))
            dtOut = LocalTime - Adjust
            Converted = True
        Case -1
            If L <> MinusPos + 5 Then Exit Sub
            Adjust = CDate(Right(strIn, 5))
            dtOut = LocalTime + Adjust
            Converted = True
    End Select

    Exit Sub
ErrHandler:
    'Was not recognised as ISO8601 date
End Sub

'See "gogeek"'s post at https://stackoverflow.com/questions/1600875/how-to-get-the-current-datetime-in-utc-from-an-excel-vba-macro
' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : GetLocalOffsetToUTC
' Author     : Philip Swannell
' Date       : 25-Aug-2021
' Purpose    :
' Parameters :
' -----------------------------------------------------------------------------------------------------------------------
Function GetLocalOffsetToUTC()
    Dim dt As Object, UTC As Date
    Dim TimeNow As Date
    On Error GoTo ErrHandler
        TimeNow = Now()

    Set dt = CreateObject("WbemScripting.SWbemDateTime")
    dt.SetVarDate TimeNow
    UTC = dt.GetVarDate(False)
    GetLocalOffsetToUTC = (TimeNow - UTC)


    Exit Function
ErrHandler:
    Throw "#GetLocalOffsetToUTC (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function ISO8601FormatString()
    Dim TimeZone
    Dim RightChars

    TimeZone = GetLocalOffsetToUTC()

    If TimeZone = 0 Then
        RightChars = "Z"
    ElseIf TimeZone > 0 Then
        RightChars = "+" & Format(TimeZone, "hh:mm")
    Else
        RightChars = "-" & Format(Abs(TimeZone), "hh:mm")
    End If
    ISO8601FormatString = "yyyy-mm-ddT:hh:mm:ss.000" & RightChars

End Function
