Attribute VB_Name = "modCompareAgainstAlternatives"
Option Explicit
Const m_FolderSpeedTest = "C:\Temp\CSVTest\CompareAgainstAlternatives"

'====================================================================================================
'Time of test = 04-Aug-2021 15:56:29 Computer = PHILIP-LAPTOP
'2.15576220001094 CSVRead           seconds to read file containing random doubles 100,000 rows, 10 cols.
'2.26041069999337 CSVRead_sdkn104   seconds to read file containing random doubles 100,000 rows, 10 cols.
'3.34533129999181 CSVRead_ws_garcia seconds to read file containing random doubles 100,000 rows, 10 cols.
'v sdk104       1.04854361950585           >1 = I'm faster
'v garcia       1.55180905388119           >1 = I'm faster
'----------
'1.65061310000601 CSVRead           seconds to read file containing 10 char Strings unquoted 100,000 rows, 10 cols.
'1.74892220000038 CSVRead_sdkn104   seconds to read file containing 10 char Strings unquoted 100,000 rows, 10 cols.
'3.27794169998378 CSVRead_ws_garcia seconds to read file containing 10 char Strings unquoted 100,000 rows, 10 cols.
'v sdk104       1.05955914199034           >1 = I'm faster
'v garcia       1.98589342346298           >1 = I'm faster
'----------
'2.51833069999702 CSVRead           seconds to read file containing 5 char strings quoted 100,000 rows, 10 cols.
'2.26482949999627 CSVRead_sdkn104   seconds to read file containing 5 char strings quoted 100,000 rows, 10 cols.
'3.81022660000599 CSVRead_ws_garcia seconds to read file containing 5 char strings quoted 100,000 rows, 10 cols.
'v sdk104       0.89933760486617           >1 = I'm faster
'v garcia       1.51299692292616           >1 = I'm faster
'----------
'2.65262430001167 CSVRead           seconds to read file containing 6 char strings quoted with line feeds 100,000 rows, 10 cols.
'2.79761830001371 CSVRead_sdkn104   seconds to read file containing 6 char strings quoted with line feeds 100,000 rows, 10 cols.
'33.6131897000014 CSVRead_ws_garcia seconds to read file containing 6 char strings quoted with line feeds 100,000 rows, 10 cols.
'v sdk104       1.05466058649972           >1 = I'm faster
'v garcia       12.6716737458273           >1 = I'm faster
'----------
'Done

Private Sub CompareAgainstAlternatives()

    Dim data As Variant
    Dim DataReread1, DataReread2, DataReread3, DataReread4
    Dim FileName As String
    Dim i As Long
    Dim j As Long
    Dim NumCols As Long
    Dim NumRows As Long
    Dim OS As String
    Dim SmallFileName As String
    Dim t1 As Double, t2 As Double, t3 As Double, t4 As Double, tstart As Double, tend As Double
    Const Unicode = False
    Dim QuoteAllStrings As Boolean
    Dim ExtraInfo As String
    Dim StringLength As Double
    Dim FnName As String

    On Error GoTo ErrHandler

    NumRows = 100000
    NumCols = 10
    StringLength = 20
    OS = "Windows"

    ThrowIfError CreatePath(m_FolderSpeedTest)
    Debug.Print String(100, "=")
    Debug.Print "Time of test = " + _
        Format(Now, "dd-mmm-yyyy hh:mm:ss") + " Computer = " + Environ("COMPUTERNAME")

    For i = 1 To 4
        Select Case i
            Case 1
                data = RandomDoubles(NumRows, NumCols)
                ExtraInfo = "random doubles"
                QuoteAllStrings = False
            Case 2
                data = sFill("abcdefghij", NumRows, NumCols)
                ExtraInfo = "10 char Strings unquoted"
                QuoteAllStrings = False
            Case 3
                data = sFill(String(StringLength, "x"), NumRows, NumCols)
                ExtraInfo = CStr(Len(data(1, 1))) & " char strings quoted"
                QuoteAllStrings = True
            Case 4
                data = sFill(String(StringLength / 2, "x") + vbCrLf + String((StringLength / 2) - 1, "y"), NumRows, NumCols)
                ExtraInfo = CStr(Len(data(1, 1))) & " char strings quoted with line feeds"
                QuoteAllStrings = True
        End Select

        FileName = NameThatFile(m_FolderSpeedTest, OS, NumRows, NumCols, Replace(ExtraInfo, " ", "-"), Unicode, False)
        ThrowIfError CSVWrite(FileName, data, QuoteAllStrings, , , , Unicode, OS, False)
        
        Debug.Print "FileName = " & FileName
        Debug.Print "Contains " + ExtraInfo + " " + _
            Format(NumRows, "###,##0") + " rows, " + Format(NumCols, "###,##0") + " cols. " '+ _
            "File size = " + Format(sFileInfo(FileName, "size"), "###,##0") + " bytes."
        For j = 1 To 4
            tstart = sElapsedTime
            Select Case j
                Case 1
                    DataReread1 = ThrowIfError(CSVRead_V3(FileName, False, ",", , , , , , False))
                    FnName = "CSVRead_V3       "
                Case 2
                    DataReread2 = ThrowIfError(CSVRead_sdkn104(FileName, Unicode))
                    FnName = "CSVRead_sdkn104  "
                Case 3
                    DataReread3 = ThrowIfError(CSVRead_ws_garcia(FileName, ",", vbCrLf))
                    FnName = "CSVRead_ws_garcia"
                Case 4
                    DataReread4 = ThrowIfError(sFileShow(FileName, ",", False, , False, vbCrLf, , , , False, , , , , False))
                    FnName = "sFileShow        "
            End Select
            tend = sElapsedTime()
            Select Case j
                Case 1
                    t1 = tend - tstart
                Case 2
                    t2 = tend - tstart
                Case 3
                    t3 = tend - tstart
                Case 4
                    t4 = tend - tstart
            End Select
            
            Debug.Print FnName + " " + CStr(tend - tstart)
        Next j
        Debug.Print "v sdk104          " & CStr(t2 / t1) & "           >1 = CSVRead_V3 faster"
        Debug.Print "v garcia          " & CStr(t3 / t1) & "           >1 = CSVRead_V3 faster"
        Debug.Print "v sFileShow       " & CStr(t4 / t1) & "           >1 = CSVRead_V3 faster"

        'Hook in to SolumAddin
        If Not Application.Run("sArraysIdentical", DataReread1, DataReread2) Then
            Debug.Print "WARNING RETURNS NOT IDENTICAL (1<>2)"
        End If
        'Comparing arrays but allowing for different lower bounds
        If Not Application.Run("sArraysIdentical", DataReread1, DataReread3, True, True) Then
            Debug.Print "WARNING RETURNS NOT IDENTICAL (1<>3)"
        End If
        Debug.Print String(10, "-")
    Next i
    Debug.Print "Done"

    Exit Sub
ErrHandler:
    MsgBox "#CompareAgainstAlternatives: " & Err.Description & "!", vbCritical
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : TimeSixParsers
' Author     : Philip Swannell
' Date       : 07-Aug-2021
' Purpose    : For use from sheet TimingResults - compares speed of 5 CSV parsing functions
' -----------------------------------------------------------------------------------------------------------------------
Function TimeSixParsers(EachFieldContains As Variant, NumRows As Long, NumCols As Long, CheckReturnsIdentical As Boolean, Optional WithHeaders As Boolean)

          Dim data As Variant
          Dim FileName As String
          Dim i As Long
          Dim j As Long
          Dim OS As String
          Dim t1 As Double, t2 As Double, t3 As Double, t4 As Double, t5 As Double, t6 As Double, tstart As Double, tend As Double
          Dim FnName1 As String
          Dim FnName2 As String
          Dim FnName3 As String
          Dim FnName4 As String
          Dim FnName5 As String
          Dim FnName6 As String
          
          Const Unicode = False
          Dim ExtraInfo As String
          Dim FnName As String
          Dim DataReread1, DataReread2, DataReread3, DataReread4, DataReread5, DataReread6

1         On Error GoTo ErrHandler
2         OS = "Windows"
          
3         If VarType(EachFieldContains) = vbDouble Then
4             ExtraInfo = "Doubles"
5         ElseIf VarType(EachFieldContains) = vbString Then
6             If Left(EachFieldContains, 1) = """" & Right(EachFieldContains, 1) = """" Then
7                 ExtraInfo = "Quoted_Strings_length_" & Len(EachFieldContains)
8             Else
9                 ExtraInfo = "Strings_length_" & Len(EachFieldContains)
10            End If
11        Else
12            ExtraInfo = "Unknown"
13        End If

14        ThrowIfError CreatePath(m_FolderSpeedTest)

15        data = sFill(EachFieldContains, NumRows, NumCols)
16        FileName = NameThatFile(m_FolderSpeedTest, OS, NumRows, NumCols, Replace(ExtraInfo, " ", "-"), Unicode, False)
17        ThrowIfError Application.Run("sFileSave", FileName, data, ",", , , , True)
              
18        For j = 1 To 6
19            tstart = sElapsedTime
20            Select Case j
                  Case 1
21                    FnName1 = "CSVRead_V1"
22                    DataReread1 = ThrowIfError(CSVRead_V1(FileName, False, ",", , , , , , , Unicode))
23                Case 2
24                    FnName2 = "CSVRead_V2"
25                    DataReread2 = ThrowIfError(CSVRead_V2(FileName, False, ",", , , , , , Unicode))
26                Case 3
27                    FnName3 = "CSVRead_V3"
28                    DataReread3 = ThrowIfError(CSVRead_V3(FileName, False, ",", , , , , , Unicode))
29                Case 4
30                    FnName4 = "CSVRead_sdkn104"
31                    DataReread4 = ThrowIfError(CSVRead_sdkn104(FileName, Unicode))
32                Case 5
33                    FnName5 = "CSVRead_ws_garcia"
34                    DataReread5 = ThrowIfError(CSVRead_ws_garcia(FileName, ",", vbCrLf))
35                Case 6
36                    FnName6 = "sFileShow"
37                    DataReread6 = ThrowIfError(sFileShow(FileName, ",", False, False, False, vbCrLf))
38            End Select
39            tend = sElapsedTime()
40            Select Case j
                  Case 1
41                    t1 = tend - tstart
42                Case 2
43                    t2 = tend - tstart
44                Case 3
45                    t3 = tend - tstart
46                Case 4
47                    t4 = tend - tstart
48                Case 5
49                    t5 = tend - tstart
50                Case 6
51                    t6 = tend - tstart
52            End Select
53        Next j

          Dim OneEqTwo, OneEqThree, OneEqFour, OneEqFive, OneEqSix

          'Hook in to SolumAddin
54        If CheckReturnsIdentical Then
55            OneEqTwo = Application.Run("sArraysIdentical", DataReread1, DataReread2)
              'Comparing arrays but allowing for different lower bounds
56            OneEqThree = Application.Run("sArraysIdentical", DataReread1, DataReread3, True, True)
57            OneEqFour = Application.Run("sArraysIdentical", DataReread1, DataReread4)
58            OneEqFive = Application.Run("sArraysIdentical", DataReread1, DataReread5)
59            OneEqSix = Application.Run("sArraysIdentical", DataReread1, DataReread6)
60        Else
61            OneEqTwo = "-"
62            OneEqThree = "-"
63            OneEqFour = "-"
64            OneEqFive = "-"
65            OneEqSix = "-"
66        End If

                      
          Dim Ret As Variant
67        ReDim Ret(1 To IIf(WithHeaders, 2, 1), 1 To 13) As Variant
          Dim DataRow As Long
68        DataRow = IIf(WithHeaders, 2, 1)
          
69        Ret(DataRow, 1) = t1: If WithHeaders Then Ret(1, 1) = FnName1
70        Ret(DataRow, 2) = t2: If WithHeaders Then Ret(1, 2) = FnName2
71        Ret(DataRow, 3) = t3: If WithHeaders Then Ret(1, 3) = FnName3
72        Ret(DataRow, 4) = t4: If WithHeaders Then Ret(1, 4) = FnName4
73        Ret(DataRow, 5) = t5: If WithHeaders Then Ret(1, 5) = FnName5
74        Ret(DataRow, 6) = t6: If WithHeaders Then Ret(1, 6) = FnName6
75        Ret(DataRow, 7) = OneEqTwo: If WithHeaders Then Ret(1, 7) = "1 = 2?"
76        Ret(DataRow, 8) = OneEqThree: If WithHeaders Then Ret(1, 8) = "1 = 3?"
77        Ret(DataRow, 9) = OneEqFour: If WithHeaders Then Ret(1, 9) = "1 = 4?"
78        Ret(DataRow, 10) = OneEqFive: If WithHeaders Then Ret(1, 10) = "1 = 5?"
79        Ret(DataRow, 11) = OneEqFive: If WithHeaders Then Ret(1, 11) = "1 = 6?"
80        Ret(DataRow, 12) = FileName: If WithHeaders Then Ret(1, 12) = "File"
81        Ret(DataRow, 13) = sFileInfo(FileName, "Size"): If WithHeaders Then Ret(1, 13) = "Size"
          

82        TimeSixParsers = Ret

83        Exit Function
ErrHandler:
84        TimeSixParsers = "#TimeSixParsers (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
