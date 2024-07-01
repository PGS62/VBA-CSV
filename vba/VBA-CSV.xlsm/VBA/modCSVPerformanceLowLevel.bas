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

'This after improved rejection of datestrings that conform to the "wrong" date format. Speed up to case "foo"
'(case of cell does not contain date separator) due to moving error handler a little later within CastToDate
'====================================================================================================
'Running SpeedTest_CastToDate 2023-Feb-24 16:36:56
'SysDateSeparator = /
'N = 1,000,000
'ComputerName = PHILIP - LAPTOP
'Calls per second = 6,047,584  strIn = "foo"                      DateOrder = 2  Result as expected? True
'Calls per second = 4,894,329  strIn = "foo-bar"                  DateOrder = 2  Result as expected? True
'Calls per second = 428,692    strIn = "09-07-2021"               DateOrder = 0  Result as expected? True
'Calls per second = 480,587    strIn = "07-09-2021"               DateOrder = 1  Result as expected? True
'Calls per second = 776,257    strIn = "2021-09-07"               DateOrder = 2  Result as expected? True
'Calls per second = 357,532    strIn = "08-24-2021 15:18:01"      DateOrder = 0  Result as expected? True
'Calls per second = 225,028    strIn = "08-24-2021 15:18:01.123"  DateOrder = 0  Result as expected? True
'Calls per second = 333,270    strIn = "24-08-2021 15:18:01"      DateOrder = 1  Result as expected? True
'Calls per second = 237,835    strIn = "24-08-2021 15:18:01.123"  DateOrder = 1  Result as expected? True
'Calls per second = 511,861    strIn = "2021-08-24 15:18:01"      DateOrder = 2  Result as expected? True
'Calls per second = 245,579    strIn = "2021-08-24 15:18:01.123"  DateOrder = 2  Result as expected? True
'Calls per second = 848,108    strIn = "2021-08-24 15:18:01.123x" DateOrder = 2  Result as expected? True

'====================================================================================================
'Running SpeedTest_CastToDate 2023-Feb-27 11:42:47
'SysDateSeparator = /
'N = 1,000,000
'ComputerName = DESKTOP - HSGAM5S
'Calls per second = 10,123,722  strIn = "foo"                     DateOrder = 2  Result as expected? True
'Calls per second = 8,909,257   strIn = "foo-bar"                 DateOrder = 2  Result as expected? True
'Calls per second = 1,462,615   strIn = "09-07-2021"              DateOrder = 0  Result as expected? True
'Calls per second = 1,466,331   strIn = "07-09-2021"              DateOrder = 1  Result as expected? True
'Calls per second = 1,706,292   strIn = "2021-09-07"              DateOrder = 2  Result as expected? True
'Calls per second = 645,595     strIn = "08-24-2021 15:18:01"     DateOrder = 0  Result as expected? True
'Calls per second = 422,385     strIn = "08-24-2021 15:18:01.123" DateOrder = 0  Result as expected? True
'Calls per second = 654,780     strIn = "24-08-2021 15:18:01"     DateOrder = 1  Result as expected? True
'Calls per second = 420,353     strIn = "24-08-2021 15:18:01.123" DateOrder = 1  Result as expected? True
'Calls per second = 850,998     strIn = "2021-08-24 15:18:01"     DateOrder = 2  Result as expected? True
'Calls per second = 420,610     strIn = "2021-08-24 15:18:01.123" DateOrder = 2  Result as expected? True
'Calls per second = 1,440,736   strIn = "2021-08-24 15:18:01.123x"DateOrder = 2  Result as expected? True

' -----------------------------------------------------------------------------------------------------------------------
Private Sub SpeedTest_CastToDate()

    Const n As Long = 1000000
    Dim Converted As Boolean
    Dim DateOrder As Long
    Dim DateSeparator As String
    Dim DtOut As Date
    Dim Expected As Date
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim strIn As String
    Dim SysDateSeparator As String
    Dim t1 As Double
    Dim t2 As Double

    SysDateSeparator = Application.International(xlDateSeparator)

    Debug.Print "'" & String(100, "=")
    Debug.Print "'Running SpeedTest_CastToDate " & Format$(Now(), "yyyy-mmm-dd hh:mm:ss")
    Debug.Print "'SysDateSeparator = " & SysDateSeparator
    Debug.Print "'N = " & Format$(n, "###,###")
    Debug.Print "'ComputerName = " & Environ("ComputerName")
    Debug.Print "'VBA-CSV Audit Sheet Version = " & shAudit.Range("Headers").Cells(2, 1).value
    
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
            For i = 1 To n
                'CastToDate strIn, DtOut, DateOrder, DateSeparator, SysDateSeparator, Converted
            Next i
            t2 = ElapsedTime()

            Dim PrintThis As String
            PrintThis = "'Calls per second = " & Format$(n / (t2 - t1), "###,###")
            If Len(PrintThis) < 30 Then PrintThis = PrintThis & String(30 - Len(PrintThis), " ")
            PrintThis = PrintThis & " strIn = """ & strIn & """"
            If Len(PrintThis) < 65 Then PrintThis = PrintThis & String(65 - Len(PrintThis), " ")
            PrintThis = PrintThis & "DateOrder = " & DateOrder & "  Result as expected? " & (Expected = DtOut)
            
            Debug.Print PrintThis
            DoEvents 'kick Immediate window to life
        Next j
    Next k
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SpeedTestCDateVDateSerial
' Author     : Philip Swannell
' Date       : 27-Feb-2023
' Purpose    : Test that shows that parsing an Iso date using DateSerial & Mid$ is faster than parsing using CDate.

'Example:
'--------------------------------------------------------------------------------
'Time:                       27/02/2023 10:03:22
'ComputerName:               DESKTOP -HSGAM5S
'VersionNumber:               229
'Elapsed time for 10,000,000 calls to CDate: 4.3078174000002 seconds
'Elapsed time for 10,000,000 calls to DateSerial: 2.51074429999971 seconds
' -----------------------------------------------------------------------------------------------------------------------
Sub SpeedTestCDateVDateSerial()

          Const TheInput As String = "2023-02-13"
          Dim TheOutput As Date
          Const n As Long = 10000000
          Dim i As Long

1         Debug.Print String(80, "-")
2         Debug.Print "Time:         ", Now
3         Debug.Print "ComputerName: ", Environ$("ComputerName")
4         Debug.Print "'VBA-CSV Audit Sheet Version = " & shAudit.Range("Headers").Cells(2, 1).value

5         tic
6         For i = 1 To n
7             TheOutput = CDate(TheInput)
8         Next
9         toc Format(n, "#,###") & " calls to CDate"
10        tic
11        For i = 1 To n
12            TheOutput = DateSerial(Mid$(TheInput, 1, 4), Mid$(TheInput, 6, 2), Mid$(TheInput, 9, 2))
13        Next
14        toc Format(n, "#,###") & " calls to DateSerial"

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
    
    Const n As Long = 10000000
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
        For i = 1 To n
            If Len(Field) <= MaxLength Then
                If Sentinels.Exists(Field) Then
                    Res = Sentinels.item(Field)
                    Converted = True
                End If
            End If
        Next i
        t2 = ElapsedTime()

        Debug.Print "Conversions per second = " & Format$(n / (t2 - t1), "###,###"), _
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

'====================================================================================================
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

'====================================================================================================
'Running SpeedTest_CastISO8601 2023-Feb-27 12:48:41
'N = 5,000,000
'ComputerName = DESKTOP-HSGAM5S
'Calls per second = 3,770,571 strIn = "xxxxxxxxxxxxxxxxxxxxxxxxxxx..."  Result as expected? True
'Calls per second = 6,544,999 strIn = "Foo"                             Result as expected? True
'Calls per second = 4,355,957 strIn = "xxxxxxxxxxxx"                    Result as expected? True
'Calls per second = 4,400,869 strIn = "xxxx-xxxxxxx"                    Result as expected? True
'Calls per second = 1,381,468 strIn = "2021-08-24T15:18:01.123+05:0x"   Result as expected? True
'Calls per second = 1,895,714 strIn = "2021-08-23"                      Result as expected? True
'Calls per second = 812,601   strIn = "2021-08-24T15:18:01"             Result as expected? True
'Calls per second = 565,381   strIn = "2021-08-23T08:47:21.123"         Result as expected? True
'Calls per second = 513,451   strIn = "2021-08-24T15:18:01+05:00"       Result as expected? True
'Calls per second = 472,148   strIn = "2021-08-24T15:18:01.123+05:00"   Result as expected? True
' -----------------------------------------------------------------------------------------------------------------------
Private Sub SpeedTest_CastISO8601()

    Const n As Long = 5000000
    Dim DtOut As Date
    Dim Expected As Date
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim PrintThis As String
    Dim strIn As String
    Dim t1 As Double
    Dim t2 As Double

    Debug.Print "'" & String(100, "=")
    Debug.Print "'Running SpeedTest_CastISO8601 " & Format$(Now(), "yyyy-mmm-dd hh:mm:ss")
    Debug.Print "'N = " & Format$(n, "###,###")
    Debug.Print "'ComputerName = " & Environ("ComputerName")
    Debug.Print "'VBA-CSV Audit Sheet Version = " & shAudit.Range("Headers").Cells(2, 1).value
    
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
            For i = 1 To n
                'CastISO8601 strIn, DtOut, Converted, True, True
            Next i
            t2 = ElapsedTime()
            
            PrintThis = "'Calls per second = " & Format$(n / (t2 - t1), "###,###")
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

'Before changes to CastToDouble
'====================================================================================================
'Running SpeedTest_CastToDouble 2023-Feb-28 09:53:32
'N = 10,000,000
'ComputerName = DESKTOP-HSGAM5S
'Calls per second = 1,299,440  strIn = "a random string", Separator = "." Result as expected? True
'Calls per second = 12,704,398 strIn = "1", Separator = "."      Result as expected? True
'Calls per second = 12,775,171 strIn = "9", Separator = "."      Result as expected? True
'Calls per second = 11,256,508 strIn = "-4", Separator = "."     Result as expected? True
'Calls per second = 12,164,803 strIn = "1e6", Separator = "."    Result as expected? True
'Calls per second = 11,258,459 strIn = ".5", Separator = "."     Result as expected? True
'Calls per second = 2,445,940  strIn = "123,4", Separator = ","  Result as expected? True
'Calls per second = 2,444,934  strIn = ",4", Separator = ","     Result as expected? True

'After changes - about 11 times faster at rejecting bad input, about 20% slower at accepting good input
'====================================================================================================
'Running SpeedTest_CastToDouble 2023-Feb-28 09:55:40
'N = 10,000,000
'ComputerName = DESKTOP-HSGAM5S
'Calls per second = 15,301,814 strIn = "a random string", Separator = "." Result as expected? True
'Calls per second = 10,090,514 strIn = "1", Separator = "."      Result as expected? True
'Calls per second = 10,169,321 strIn = "9", Separator = "."      Result as expected? True
'Calls per second = 9,706,174  strIn = "-4", Separator = "."     Result as expected? True
'Calls per second = 9,670,801  strIn = "1e6", Separator = "."    Result as expected? True
'Calls per second = 9,356,131  strIn = ".5", Separator = "."     Result as expected? True
'Calls per second = 2,367,018  strIn = "123,4", Separator = ","  Result as expected? True
'Calls per second = 2,362,177  strIn = ",4", Separator = ","     Result as expected? True
Private Sub SpeedTest_CastToDouble()

          Const n As Long = 10000000
          Dim Converted As Boolean
          Dim DblOut As Double
          Dim Expected As Date
          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim strIn As String
          Dim t1 As Double
          Dim t2 As Double
          Dim SepStandard As Boolean
          Dim DecimalSeparator As String
          Dim SysDecimalSeparator As String
          Dim AscSeparator As Long

1         Debug.Print "'" & String(100, "=")
2         Debug.Print "'Running SpeedTest_CastToDouble " & Format$(Now(), "yyyy-mmm-dd hh:mm:ss")
3         Debug.Print "'N = " & Format$(n, "###,###")
4         Debug.Print "'ComputerName = " & Environ("ComputerName")
5         Debug.Print "'VBA-CSV Audit Sheet Version = " & shAudit.Range("Headers").Cells(2, 1).value
6         SysDecimalSeparator = Application.DecimalSeparator
          
7         For k = 1 To 8
8             For j = 1 To 1
9                 DblOut = 0
10                Converted = False
11                Select Case k
                      Case 1
12                        strIn = "a random string"
13                        DecimalSeparator = "."
14                        SepStandard = DecimalSeparator = SysDecimalSeparator
15                        AscSeparator = Asc(DecimalSeparator)
16                        Expected = CDbl(0)
17                    Case 2
18                        strIn = "1"
19                        DecimalSeparator = "."
20                        SepStandard = DecimalSeparator = SysDecimalSeparator
21                        AscSeparator = Asc(DecimalSeparator)
22                        Expected = CDbl(1)
23                    Case 3
24                        strIn = "9"
25                        DecimalSeparator = "."
26                        SepStandard = DecimalSeparator = SysDecimalSeparator
27                        AscSeparator = Asc(DecimalSeparator)
28                        Expected = CDbl(9)
29                    Case 4
30                        strIn = "-4"
31                        DecimalSeparator = "."
32                        SepStandard = DecimalSeparator = SysDecimalSeparator
33                        AscSeparator = Asc(DecimalSeparator)
34                        Expected = CDbl(-4)
35                    Case 5
36                        strIn = "1e6"
37                        DecimalSeparator = "."
38                        SepStandard = DecimalSeparator = SysDecimalSeparator
39                        AscSeparator = Asc(DecimalSeparator)
40                        Expected = CDbl(1000000)
41                    Case 6
42                        strIn = ".5"
43                        DecimalSeparator = "."
44                        SepStandard = DecimalSeparator = SysDecimalSeparator
45                        AscSeparator = Asc(DecimalSeparator)
46                        Expected = 0.5
47                    Case 7
48                        strIn = "123,4"
49                        DecimalSeparator = ","
50                        SepStandard = DecimalSeparator = SysDecimalSeparator
51                        AscSeparator = Asc(DecimalSeparator)
52                        Expected = 123.4
53                    Case 8
54                        strIn = ",4"
55                        DecimalSeparator = ","
56                        SepStandard = DecimalSeparator = SysDecimalSeparator
57                        AscSeparator = Asc(DecimalSeparator)
58                        Expected = 0.4
59                End Select

60                t1 = ElapsedTime()
61                For i = 1 To n
62                    CastToDouble strIn, DblOut, SepStandard, DecimalSeparator, AscSeparator, SysDecimalSeparator, Converted
63                Next i
64                t2 = ElapsedTime()

                  Dim PrintThis As String
65                PrintThis = "'Calls per second = " & Format$(n / (t2 - t1), "###,###")
66                If Len(PrintThis) < 30 Then PrintThis = PrintThis & String(30 - Len(PrintThis), " ")
67                PrintThis = PrintThis & " strIn = """ & strIn & """, Separator = " & """" & DecimalSeparator & """ "
68                If Len(PrintThis) < 65 Then PrintThis = PrintThis & String(65 - Len(PrintThis), " ")
69                PrintThis = PrintThis & "Result as expected? " & (Expected = DblOut)
                  
70                Debug.Print PrintThis
71                DoEvents 'kick Immediate window to life
72            Next j
73        Next k
End Sub

