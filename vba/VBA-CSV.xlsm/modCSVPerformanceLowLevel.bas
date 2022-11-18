Attribute VB_Name = "modCSVPerformanceLowLevel"
Option Explicit

'Performance testing for three methods of modCSVReadWrite: CastToDate, CastISO8601 and use of Sentinels dictionary

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SpeedTest_CastToDate
' Purpose    : Helps tune CastToDate for speed.
' Usage      : Before running this method, two edits are required:
'              a) Make method modCSVReadWrite.CastToDate Public instead of Private
'              b) Uncomment call to CastToDate
'              Don't forget to revert those changes!

'Example output:
'Running SpeedTest_CastToDate 2021-Sep-07 13:36:55
'SysDateOrder = 0
'SysDateSeparator = /
'N = 1,000,000
'Calls per second = 4,733,596  strIn = "foo"                      DateOrder = 2  Result as expected? True
'Calls per second = 4,005,577  strIn = "foo-bar"                  DateOrder = 2  Result as expected? True
'Calls per second = 771,609    strIn = "09-07-2021"               DateOrder = 0  Result as expected? True
'Calls per second = 500,064    strIn = "07-09-2021"               DateOrder = 1  Result as expected? True
'Calls per second = 729,058    strIn = "2021-09-07"               DateOrder = 2  Result as expected? True
'Calls per second = 379,716    strIn = "08-24-2021 15:18:01"      DateOrder = 0  Result as expected? True
'Calls per second = 202,723    strIn = "08-24-2021 15:18:01.123"  DateOrder = 0  Result as expected? True
'Calls per second = 321,445    strIn = "24-08-2021 15:18:01"      DateOrder = 1  Result as expected? True
'Calls per second = 200,997    strIn = "24-08-2021 15:18:01.123"  DateOrder = 1  Result as expected? True
'Calls per second = 375,057    strIn = "2021-08-24 15:18:01"      DateOrder = 2  Result as expected? True
'Calls per second = 207,064    strIn = "2021-08-24 15:18:01.123"  DateOrder = 2  Result as expected? True
'Calls per second = 475,397    strIn = "2021-08-24 15:18:01.123x" DateOrder = 2  Result as expected? True
' -----------------------------------------------------------------------------------------------------------------------
Private Sub SpeedTest_CastToDate()

    Const N As Long = 1000000
    Dim Converted As Boolean
    Dim DateOrder As Long
    Dim DateSeparator As String
    Dim DtOut As Date
    Dim Expected As Date
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim strIn As String
    Dim SysDateOrder As Long
    Dim SysDateSeparator As String
    Dim t1 As Double
    Dim t2 As Double

    '0 = month-day-year, 1 = day-month-year, 2 = year-month-day
    SysDateOrder = Application.International(xlDateOrder)
    SysDateSeparator = Application.International(xlDateSeparator)

    Debug.Print String(100, "=")
    Debug.Print "Running SpeedTest_CastToDate " & Format$(Now(), "yyyy-mmm-dd hh:mm:ss")
    Debug.Print "SysDateOrder = " & SysDateOrder
    Debug.Print "SysDateSeparator = " & SysDateSeparator
    Debug.Print "N = " & Format$(N, "###,###")
    
    For k = 1 To 12
        For j = 1 To 1 'Maybe do multiple times to test for variability or results.
            DtOut = 0
            Converted = False
            Select Case k
                Case 1
                    DateOrder = 2
                    DateSeparator = "-"
                    strIn = "foo" 'Contains no date separator, so rejected quickly by CastToDate
                    Expected = CDate(0)
                Case 2
                    DateOrder = 2
                    DateSeparator = "-"
                    strIn = "foo-bar" 'Contains only one date separator, so rejected quickly by CastToDate
                    Expected = CDate(0)
                Case 3
                    DateOrder = 0 'month-day-year
                    DateSeparator = "-"
                    strIn = "09-07-2021"
                    Expected = CDate("2021-Sep-07")
                Case 4
                    DateOrder = 1 'day-month-year
                    DateSeparator = "-"
                    strIn = "07-09-2021"
                    Expected = CDate("2021-Sep-07")
                Case 5
                    DateOrder = 2   'year-month-day
                    DateSeparator = "-"
                    strIn = "2021-09-07"
                    Expected = CDate("2021-Sep-07")
                Case 6
                    DateOrder = 0
                    DateSeparator = "-"
                    strIn = "08-24-2021 15:18:01" 'date with time, no fractions of second
                    Expected = CDate("2021-Aug-24 15:18:01")
                Case 7
                    DateOrder = 0
                    DateSeparator = "-"
                    strIn = "08-24-2021 15:18:01.123" 'date with time, with fractions of second
                    Expected = CDate("2021-Aug-24 15:18:01") + 0.123 / 86400
                Case 8
                    DateOrder = 1
                    DateSeparator = "-"
                    strIn = "24-08-2021 15:18:01" 'date with time, no fractions of second
                    Expected = CDate("2021-Aug-24 15:18:01")
                Case 9
                    DateOrder = 1
                    DateSeparator = "-"
                    strIn = "24-08-2021 15:18:01.123" 'date with time, with fractions of second
                    Expected = CDate("2021-Aug-24 15:18:01") + 0.123 / 86400
                Case 10
                    DateOrder = 2
                    DateSeparator = "-"
                    strIn = "2021-08-24 15:18:01" 'date with time, no fractions of second
                    Expected = CDate("2021-Aug-24 15:18:01")
                Case 11
                    DateOrder = 2
                    DateSeparator = "-"
                    strIn = "2021-08-24 15:18:01.123" 'date with time, with fractions of second
                    Expected = CDate("2021-Aug-24 15:18:01") + 0.123 / 86400
                Case 12
                    DateOrder = 2
                    DateSeparator = "-"
                    strIn = "2021-08-24 15:18:01.123x" 'Nearly a date, but final "x" stops it being so
                    Expected = CDate(0)
            End Select

            t1 = ElapsedTime()
            For i = 1 To N
                'CastToDate strIn, dtOut, DateOrder, DateSeparator, SysDateOrder, SysDateSeparator, Converted
            Next i
            t2 = ElapsedTime()
            Dim PrintThis As String
            PrintThis = "Calls per second = " & Format$(N / (t2 - t1), "###,###")
            If Len(PrintThis) < 30 Then PrintThis = PrintThis & String(30 - Len(PrintThis), " ")
            PrintThis = PrintThis & "strIn = """ & strIn & """"
            If Len(PrintThis) < 65 Then PrintThis = PrintThis & String(65 - Len(PrintThis), " ")
            PrintThis = PrintThis & "DateOrder = " & DateOrder & "  Result as expected? " & (Expected = DtOut)
            
            Debug.Print PrintThis
            DoEvents 'kick Immediate window to life
        Next j
    Next k
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SpeedTest_Sentinels
' Purpose    : Test speed of accessing the sentinels dictionary, using similar approach to that employed in method
'              ConvertField.
' Usage      : Before running this method, two edits are required:
'              a) Make method modCSVReadWrite.MakeSentinels Public instead of Private
'              b) Uncomment call to MakeSentinels
'              Don't forget to revert those changes!

'
' Results:  On Surface Book 2, Intel(R) Core(TM) i7-8650U CPU @ 1.90GHz   2.11 GHz, 16GB RAM
'
'Running SpeedTest_Sentinels 2021-08-25T15:01:33
'Conversions per second = 90,346,968       Field = "This string is longer than the longest sentinel, which is 14"
'Conversions per second = 20,976,150       Field = "mini" (Not a sentinel, but shorter than the longest sentinel)
'Conversions per second = 9,295,050        Field = "True" (A sentinel, one of the elements of TrueStrings)
' -----------------------------------------------------------------------------------------------------------------------
Private Sub SpeedTest_Sentinels()
    
    Const N As Long = 10000000
    Dim Comment As String
    Dim Field As String
    Dim i As Long
    Dim j As Long
    Dim MaxLength As Long
    Dim Res As Variant
    Dim Sentinels As Scripting.Dictionary
    Dim t1 As Double
    Dim t2 As Double

    On Error GoTo ErrHandler
    
        Set Sentinels = New Scripting.Dictionary
      '  MakeSentinels Sentinels, ConvertQuoted, MaxLength, AnySentinels, _
        ShowBooleansAsBooleans:=True, _
        ShowErrorsAsErrors:=True, _
        ShowMissingsAs:=Empty, _
        TrueStrings:=Array("True", "T"), _
        FalseStrings:=Array("False", "F"), _
        MissingStrings:=Array("NA", "-999")
    
    Dim Converted As Boolean
    
    Debug.Print "Running SpeedTest_Sentinels " & Format$(Now(), "yyyy-mm-ddThh:mm:ss")
    
    For j = 1 To 3
    
        Select Case j
            Case 1
                Field = "This string is longer than the longest sentinel, which is 14"
            Case 2
                Field = "mini"
                Comment = "Not a sentinel, but shorter than the longest sentinel"
            Case 3
                Field = "True"
                Comment = "A sentinel, one of the elements of TrueStrings"
        End Select

        t1 = ElapsedTime()
        For i = 1 To N
            If Len(Field) <= MaxLength Then
                If Sentinels.Exists(Field) Then
                    Res = Sentinels.item(Field)
                    Converted = True
                End If
            End If
        Next i
        t2 = ElapsedTime()

        Debug.Print "Conversions per second = " & Format$(N / (t2 - t1), "###,###"), _
            "Field = """ & Field & """" & IIf(Comment = vbNullString, vbNullString, " (" & Comment & ")")

    Next j

    Exit Sub
ErrHandler:
    MsgBox ReThrow("SpeedTest_Sentinels", Err, True), vbCritical
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SpeedTest_CastISO8601
' Purpose    : Testing speed of CastISO8601

' Usage      : Before running this method, two edits are required:
'              a) Make method modCSVReadWrite.CastISO8601 Public instead of Private
'              b) Uncomment call to CastISO8601
'              Don't forget to revert those changes!

'Example output: (Surface Book 2, Intel(R) Core(TM) i7-8650U CPU @ 1.90GHz   2.11 GHz, 16GB RAM)

'Running SpeedTest_CastISO8601 2021-Sep-07 22:03:27
'N = 5,000,000
'Calls per second = 1,052,769  strIn = "xxxxxxxxxxxxxxxxxxxxxxxxxxx..."  Result as expected? True
'Calls per second = 2,436,414  strIn = "Foo"                             Result as expected? True
'Calls per second = 1,718,279  strIn = "xxxxxxxxxxxx"                    Result as expected? True
'Calls per second = 1,718,023  strIn = "xxxx-xxxxxxx"                    Result as expected? True
'Calls per second = 587,754    strIn = "2021-08-24T15:18:01.123+05:0x"   Result as expected? True
'Calls per second = 574,610    strIn = "2021-08-23"                      Result as expected? True
'Calls per second = 348,325    strIn = "2021-08-24T15:18:01"             Result as expected? True
'Calls per second = 247,093    strIn = "2021-08-23T08:47:21.123"         Result as expected? True
'Calls per second = 221,942    strIn = "2021-08-24T15:18:01+05:00"       Result as expected? True
'Calls per second = 191,331    strIn = "2021-08-24T15:18:01.123+05:00"   Result as expected? True
' -----------------------------------------------------------------------------------------------------------------------
Private Sub SpeedTest_CastISO8601()

    Const N As Long = 5000000
    Dim DtOut As Date
    Dim Expected As Date
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim PrintThis As String
    Dim strIn As String
    Dim t1 As Double
    Dim t2 As Double

    Debug.Print String(100, "=")
    Debug.Print "Running SpeedTest_CastISO8601 " & Format$(Now(), "yyyy-mmm-dd hh:mm:ss")
    Debug.Print "N = " & Format$(N, "###,###")
    For k = 0 To 9
        For j = 1 To 1
            DtOut = 0
            Select Case k
                Case 0
                    strIn = String(10000, "x")
                    Expected = CDate(0)
                Case 1
                    strIn = "Foo" ' less than 10 in length
                    Expected = CDate(0)
                Case 2
                    strIn = "xxxxxxxxxxxx" '5th character not "-"
                    Expected = CDate(0)
                Case 3
                    strIn = "xxxx-xxxxxxx" 'rejected by RegEx
                    Expected = CDate(0)
                Case 4
                    strIn = "2021-08-24T15:18:01.123+05:0x" ' rejected by regex
                    Expected = CDate(0)
                Case 5
                    strIn = "2021-08-23"
                    Expected = CDate("2021-Aug-23")
                Case 6
                    strIn = "2021-08-24T15:18:01"
                    Expected = CDate("2021-Aug-24 15:18:01")
                Case 7
                    strIn = "2021-08-23T08:47:21.123"
                    Expected = CDate("2021-Aug-23 08:47:21") + 0.123 / 86400
                Case 8
                    strIn = "2021-08-24T15:18:01+05:00"
                    Expected = CDate("2021-Aug-24 15:18:01") - 5 / 24
                Case 9
                    strIn = "2021-08-24T15:18:01.123+05:00"
                    Expected = CDate("2021-Aug-24 15:18:01") + 0.123 / 86400 - 5 / 24
            End Select

            t1 = ElapsedTime()
            For i = 1 To N
                'CastISO8601 strIn, dtOut, Converted, True, True
            Next i
            t2 = ElapsedTime()
            
            PrintThis = "Calls per second = " & Format$(N / (t2 - t1), "###,###")
            If Len(PrintThis) < 30 Then PrintThis = PrintThis & String(30 - Len(PrintThis), " ")
            If Len(strIn) > 30 Then
                PrintThis = PrintThis & "strIn = """ & Left$(strIn, 27) & "..."""
            Else
                PrintThis = PrintThis & "strIn = """ & strIn & """"
            End If
            If Len(PrintThis) < 70 Then PrintThis = PrintThis & String(70 - Len(PrintThis), " ")
            PrintThis = PrintThis & "  Result as expected? " & (Expected = DtOut)
            
            Debug.Print PrintThis
            DoEvents 'kick Immediate window to life
        Next j
    Next k

End Sub

