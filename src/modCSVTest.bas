Attribute VB_Name = "modCSVTest"
Option Explicit
Private Const m_FolderOriginals = "c:\temp\CSVTest\Originals"
Private Const m_FolderReadAndRewrite = "c:\temp\CSVTest\ReadAndWritten"
Private Const m_FolderSpeedTest = "C:\Temp\CSVTest\SpeedTest"

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CSVSpeedTest
' Author     : Philip Swannell
' Date       : 19-Jul-2021
' Purpose    : Testing speed of CSVRead - record results below...
'2021-07-20 16:00:13.150   ====================================================================================================
'2021-07-20 16:00:13.152   SolumAddin version 2,166. Time of test = 20-Jul-2021 16:00:13
'2021-07-20 16:00:13.152   Time to read random doubles 10,000 rows, 100 cols = 3.68973119999998 seconds. File size = 18,180,900 bytes.
'2021-07-20 16:00:20.735   Time to read 10-char strings 10,000 rows, 100 cols = 3.03961860000004 seconds. File size = 11,010,000 bytes.
'2021-07-20 16:00:37.664   Time to read 10-char strings, all with line-feeds 10,000 rows, 100 cols = 11.2207711 seconds. File size = 14,010,000 bytes.
'2021-07-20 16:00:42.174   Time to read 1000 files = 2.80944360000001 seconds.
'2021-07-20 16:00:42.174   ====================================================================================================
'2021-07-20 19:47:32.791   ====================================================================================================
'2021-07-20 19:47:32.791   SolumAddin version 2,170. Time of test = 20-Jul-2021 19:47:32 Computer = PHILIP-LAPTOP
'2021-07-20 19:47:38.880   Time to read 1 file containing random doubles 10,000 rows, 100 cols = 2.04967829999987 seconds. File size = 18,180,900 bytes.
'2021-07-20 19:47:43.493   Time to read 1 file containing 10-char strings 10,000 rows, 100 cols = 1.96609960000023 seconds. File size = 11,010,000 bytes.
'2021-07-20 19:47:52.498   Time to read 1 file containing 10-char strings, all with line-feeds 10,000 rows, 100 cols = 5.65375040000072 seconds. File size = 14,010,000 bytes.
'2021-07-20 19:47:58.406   Time to write 1000 files = 4.86368840000068 seconds. Each file has 70 rows and 6 columns
'2021-07-20 19:48:06.752   Time to read 1000 files = 8.34504259999994 seconds. Each file has 70 rows and 6 columns
'2021-07-20 19:48:06.752   ====================================================================================================
'2021-07-27 11:15:51.958   ====================================================================================================
'2021-07-27 11:15:51.958   SolumAddin version 2,188. Time of test = 27-Jul-2021 11:15:51 Computer = PHILIP-LAPTOP
'2021-07-27 11:15:56.201   1.81936810002662 seconds to read 1 file containing random doubles 10,000 rows, 100 cols. File size = 18,180,900 bytes.
'2021-07-27 11:15:59.938   1.53920500003733 seconds to read 1 file containing UNquoted 10-char strings 10,000 rows, 100 cols. File size = 11,010,000 bytes.
'2021-07-27 11:16:08.306   5.71575269999448 seconds to read 1 file containing QUOTED 10-char strings 10,000 rows, 100 cols. File size = 13,010,000 bytes.
'2021-07-27 11:16:17.049   6.22214199998416 seconds to read 1 file containing 10-char strings, all with line-feeds 10,000 rows, 100 cols. File size = 15,010,000 bytes.
'2021-07-27 11:16:21.069   2.97768730006646 seconds to write 1000 files. Each file has 70 rows and 6 columns.
'2021-07-27 11:16:22.194   1.12387880007736 seconds to read 1000 files. Each file has 70 rows and 6 columns.
'2021-07-27 11:16:22.194   ====================================================================================================
'====================================================================================================
'Time of test = 02-Aug-2021 17:41:42 Computer = PHILIP-LAPTOP
'1.83393520000027 seconds to read 1 file containing random doubles 10,000 rows, 100 cols.
'1.6036654999989 seconds to read 1 file containing UNquoted 10-char strings 10,000 rows, 100 cols. File size =
'4.82088320000184 seconds to read 1 file containing QUOTED 10-char strings 10,000 rows, 100 cols. File size =
'5.52146999999968 seconds to read 1 file containing 10-char strings, all with line-feeds 10,000 rows, 100 cols. File size =
'3.71276379999836 seconds to write 1000 files. Each file has 70 rows and 6 columns.
'10.0182820000009 seconds to read 1000 files. Each file has 70 rows and 6 columns.
'====================================================================================================

' -----------------------------------------------------------------------------------------------------------------------
Private Sub CSVSpeedTest()

          Const NumColsSmall = 6
          Const NumFilesToReadAndWrite = 1000
          Const NumRowsSmall = 70
          Dim Data, DataReread
          Dim FileName As String
          Dim i As Long
          Dim NumCols As Long
          Dim NumRows As Long
          Dim OS As String
          Dim SmallFileName As String
          Dim T1 As Double, t2 As Double

1         On Error GoTo ErrHandler

3         NumRows = 10000
4         NumCols = 100
5         OS = "Windows"

6         ThrowIfError CreatePath(m_FolderSpeedTest)
7         Debug.Print String(100, "=")
8         Debug.Print "Time of test = " + _
              Format(Now, "dd-mmm-yyyy hh:mm:ss") + " Computer = " + Environ("COMPUTERNAME")

          'Doubles only, cast back to doubles
10        Data = RandomDoubles(NumRows, NumCols)
11        FileName = NameThatFile(m_FolderSpeedTest, OS, NumRows, NumCols, "Doubles", False, False)
12        ThrowIfError CSVWrite(FileName, Data, True, , , , False, OS, False)
13        T1 = sElapsedTime
14        DataReread = ThrowIfError(CSVRead(FileName, True))
15        t2 = sElapsedTime
16        Debug.Print CStr(t2 - T1) + " seconds to read 1 file containing random doubles " + _
              Format(NumRows, "###,##0") + " rows, " + Format(NumCols, "###,##0") + " cols. " '+ _
              "File size = " + Format(sFileInfo(FileName, "size"), "###,##0") + " bytes."

          '10-character strings, unquoted
20        Data = sFill("abcdefghij", NumRows, NumCols)
21        FileName = NameThatFile(m_FolderSpeedTest, OS, NumRows, NumCols, "10-char-strings-unquoted", False, False)
22        ThrowIfError CSVWrite(FileName, Data, False, , , , , OS)
23        T1 = sElapsedTime
24        DataReread = ThrowIfError(CSVRead(FileName, False))
25        t2 = sElapsedTime
26        Debug.Print CStr(t2 - T1) + " seconds to read 1 file containing UNquoted 10-char strings " + _
              Format(NumRows, "###,##0") + " rows, " + _
              Format(NumCols, "###,##0") + " cols. File size = " ' + _
              Format(sFileInfo(FileName, "size"), "###,##0") + " bytes."

          '10-character strings...
30        Data = sFill("abcdefghij", NumRows, NumCols)
31        FileName = NameThatFile(m_FolderSpeedTest, OS, NumRows, NumCols, "10-char-strings", False, False)
32        ThrowIfError CSVWrite(FileName, Data, , , , , , OS, False)
33        T1 = sElapsedTime
34        DataReread = ThrowIfError(CSVRead(FileName, False))
35        t2 = sElapsedTime
36        Debug.Print CStr(t2 - T1) + " seconds to read 1 file containing QUOTED 10-char strings " + _
              Format(NumRows, "###,##0") + " rows, " + _
              Format(NumCols, "###,##0") + " cols. File size = " ' + _
              Format(sFileInfo(FileName, "size"), "###,##0") + " bytes."

          '10-character strings ALL with linefeeds
40        Data = sFill("abcd+" + vbCrLf + "efghi", NumRows, NumCols)
41        FileName = NameThatFile(m_FolderSpeedTest, OS, NumRows, NumCols, "10-char-strings-with-line-feeds", False, False)
42        ThrowIfError CSVWrite(FileName, Data, , , , , , OS, False)
43        T1 = sElapsedTime
44        DataReread = ThrowIfError(CSVRead(FileName))
45        t2 = sElapsedTime
46        Debug.Print CStr(t2 - T1) + " seconds to read 1 file containing 10-char strings, all with line-feeds " + _
              Format(NumRows, "###,##0") + " rows, " + Format(NumCols, "###,##0") + " cols. File size = " ' + _
              Format(sFileInfo(FileName, "size"), "###,##0") + " bytes."

          'Write and read many files
          
          'Create files
50        T1 = sElapsedTime()
52        For i = 1 To NumFilesToReadAndWrite
53            SmallFileName = NameThatFile(m_FolderSpeedTest, OS, NumRowsSmall, NumColsSmall, Format(i, "0000"), False, False)
54            Data = RandomDoubles(NumRowsSmall, NumColsSmall)
55            ThrowIfError CSVWrite(SmallFileName, Data)
56        Next i
57        t2 = sElapsedTime()
58        Debug.Print CStr(t2 - T1) + " seconds to write " + CStr(NumFilesToReadAndWrite) + " files. " + _
              "Each file has " + CStr(NumRowsSmall) + " rows and " + CStr(NumColsSmall) + " columns."
          
          'Read them back
59        T1 = sElapsedTime()
60        For i = 1 To NumFilesToReadAndWrite
61            SmallFileName = NameThatFile(m_FolderSpeedTest, OS, NumRowsSmall, NumColsSmall, Format(i, "0000"), False, False)
62            Data = ThrowIfError(CSVRead(SmallFileName, True, ","))
63        Next i
64        t2 = sElapsedTime()
65        Debug.Print CStr(t2 - T1) + " seconds to read " + CStr(NumFilesToReadAndWrite) + " files. " + _
              "Each file has " + CStr(NumRowsSmall) + " rows and " + CStr(NumColsSmall) + " columns."
66        Debug.Print String(100, "=")

67        Exit Sub
ErrHandler:
68        MsgBox "#CSVSpeedTest (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical
End Sub

Private Function NameThatFile(Folder As String, ByVal OS As String, NumRows As Long, NumCols As Long, ExtraInfo As String, Unicode As Boolean, Ragged As Boolean)
1         NameThatFile = (Folder & "\" & OS & "_" & Format(NumRows, "0000") & "_x_" & Format(NumCols, "000") & "_" & ExtraInfo & IIf(Unicode, "_Unicode", "_Ascii") & IIf(Ragged, "_Ragged", "_NotRagged") & ".csv")
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CSVRoundTripTestMulti
' Author     : Philip Swannell
' Date       : 22-Jul-2021
' Purpose    : Tests multiple times that CSVRead correctly round-trips data previously saved to disk by CSVWrite.
'              Tests include:
'           *  Embedded line feeds in quoted strings.
'           *  Files with Windows, Unix or (old) Mac line endings.
'           *  Both unicode and ascii files.
'           *  Files with varying number of fields in each line (tricky since CSVWrite does not support creating such files).
'           *  That the delimiter is automatically detected by CSVRead (reliable only if files have all strings quoted).
'           *  That unicode vs ascii is automatically detected.
'           *  That line endings are automatically detected.
' -----------------------------------------------------------------------------------------------------------------------
Sub AnotherWay()
          Dim OS As Variant
          Dim Unicode As Variant
          Dim NRows As Variant
          Dim NCols As Variant
          Dim Data As Variant
          Dim k As Long
          Dim DateFormat As Variant
          Dim AllowLineFeed As Variant
          Const WholeFile = True
          Dim Ragged As Variant
          Dim ExtraInfo As String
          Dim EOL As String

1         On Error GoTo ErrHandler

2         ThrowIfError CreatePath(m_FolderOriginals)
3         ThrowIfError CreatePath(m_FolderReadAndRewrite)

4         For Each OS In Array("Windows", "Unix", "Mac")
5             EOL = IIf(OS = "Windows", vbCrLf, IIf(OS = "Unix", vbLf, vbCr))
          
6             For Each Unicode In Array(True, False)
7                 For Each Ragged In Array(True, False)
8                     For Each NRows In Array(1, 2, 20)
9                         For Each NCols In Array(1, 2, 10)
                    
                              'For Variants we need to vary AllowLineFeed and DateFormat
10                            For Each AllowLineFeed In Array(True, False)
11                                For Each DateFormat In Array("mmm-dd-yyyy", "dd-mmm-yyyy", "yyyy-mm-dd")
12                                    Data = RandomVariants(CLng(NRows), CLng(NCols), CBool(AllowLineFeed), CBool(Unicode), EOL)
13                                    ExtraInfo = "RandomVariants" & IIf(AllowLineFeed, "WithLineFeed", "")
14                                    CSVRoundTripTest CStr(OS), Data, CStr(DateFormat), CBool(Unicode), CStr(OS), vbTab, CBool(Ragged), ExtraInfo
15                                Next DateFormat
16                            Next AllowLineFeed

                              'For Dates, we need to vary DateFormat
17                            For Each DateFormat In Array("mmm-dd-yyyy", "dd-mmm-yyyy", "yyyy-mm-dd")
18                                Data = RandomDates(CLng(NRows), CLng(NCols))
19                                ExtraInfo = "RandomDates"
20                                CSVRoundTripTest CStr(OS), Data, CStr(DateFormat), CBool(Unicode), CStr(OS), vbTab, CBool(Ragged), ExtraInfo
21                            Next DateFormat

                              'For Strings, we need to vary AllowLineFeed
22                            For Each AllowLineFeed In Array(True, False)
23                                Data = RandomStrings(CLng(NRows), CLng(NCols), CBool(Unicode), CBool(AllowLineFeed), EOL)
24                                ExtraInfo = IIf(AllowLineFeed, "RandomStringsWithLineFeeds", "RandomStrings")
25                                CSVRoundTripTest CStr(OS), Data, CStr(DateFormat), CBool(Unicode), CStr(OS), vbTab, CBool(Ragged), ExtraInfo
26                            Next AllowLineFeed

27                            For k = 1 To 4
28                                If k = 1 Then
29                                    Data = RandomBooleans(CLng(NRows), CLng(NCols))
30                                    ExtraInfo = "RandomBooleans"
31                                ElseIf k = 2 Then
32                                    Data = RandomDoubles(CLng(NRows), CLng(NCols))
33                                    ExtraInfo = "RandomDoubles"
34                                ElseIf k = 3 Then
35                                    Data = RandomErrorValues(CLng(NRows), CLng(NCols))
36                                    ExtraInfo = "RandomErrorValues"
37                                ElseIf k = 4 Then
38                                    Data = RandomLongs(CLng(NRows), CLng(NCols))
39                                    ExtraInfo = "RandomLongs"
40                                End If
41                                CSVRoundTripTest CStr(OS), Data, CStr(DateFormat), CBool(Unicode), CStr(OS), vbTab, CBool(Ragged), ExtraInfo
42                            Next k
43                        Next NCols
44                    Next NRows
45                Next Ragged
46            Next Unicode
47        Next OS

48        Exit Sub
ErrHandler:
49        Throw "#AnotherWay (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub


' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CSVRoundTripTest
' Author     : Philip Swannell
' Date       : 30-Jul-2021
' Purpose    : Test for "round trip" between functions CSVRead and CSVWrite. We write data to a file, then read it back
'              then write the read-back data to a second file and test if the two files are identical.
' Parameters :
'  OS        :
'  Data      :
'  DateFormat:
'  Unicode   :
'  EOL       :
'  Delimiter :
'  Ragged    :
'  ExtraInfo :
' -----------------------------------------------------------------------------------------------------------------------
Function CSVRoundTripTest(OS As String, ByVal Data As Variant, DateFormat As String, Unicode As Boolean, EOL As String, Delimiter As String, Ragged As Boolean, ExtraInfo As String)

          Dim DataReadBack

1         On Error GoTo ErrHandler
          Dim FileName1 As String
          Dim FileName2 As String
          Dim NR As Long
          Dim NC As Long

2         NR = sNRows(Data)
3         NC = sNCols(Data)

4         FileName1 = NameThatFile(m_FolderOriginals, OS, NR, NC, ExtraInfo, CBool(Unicode), CBool(Ragged))
5         FileName2 = NameThatFile(m_FolderReadAndRewrite, OS, NR, NC, ExtraInfo, CBool(Unicode), CBool(Ragged))

6         If Ragged Then Data = MakeArrayRagged(Data)

7         ThrowIfError CSVWrite(FileName1, Data, True, DateFormat, , Delimiter, Unicode, EOL, Ragged)

          'The Call to CSVRead has to infer both Unicode and EOL
8         DataReadBack = CSVRead(FileName1, True, , DateFormat, , , , , , , Empty)

9         ThrowIfError CSVWrite(FileName2, Data, True, DateFormat, , Delimiter, Unicode, EOL, Ragged)

10        If Not TextFilesIdentical(FileName1, FileName2, IIf(Unicode, TristateTrue, TristateFalse)) Then
11            Stop
12        End If

13        Exit Function
ErrHandler:
14        Throw "#CSVRoundTripTest (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : MakeArrayRagged
' Author     : Philip Swannell
' Date       : 28-Jul-2021
' Purpose    : For each row of an array choose random number n less than number of cols and make the n right most columns empty
'              also guarantee that one row will not have an empty right-most column.
' -----------------------------------------------------------------------------------------------------------------------
Private Function MakeArrayRagged(Data)

          Dim NR As Long, NC As Long
          Dim i As Long, j As Long
          Dim RowToLeaveUnchanged As Long

1         On Error GoTo ErrHandler
2         NR = sNRows(Data)
3         NC = sNCols(Data)
4         RowToLeaveUnchanged = CLng(0.5 + Rnd() * (NR))

5         For i = 1 To NR
6             If i = RowToLeaveUnchanged Then
7                 If IsEmpty(Data(i, NC)) Then
8                     Data(i, NC) = "Not empty!"
9                 End If
10            Else
11                For j = CLng(0.5 + Rnd() * NC) To NC
12                    Data(i, j) = Empty
13                Next
14            End If
15        Next
16        MakeArrayRagged = Data

17        Exit Function
ErrHandler:
18        Throw "#MakeArrayRagged (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function



Function TextFilesIdentical(File1 As String, File2 As String, Format As Scripting.Tristate) As Boolean
          Static FSO As FileSystemObject

          Dim Contents1 As String
          Dim Contents2 As String
          Dim T As Scripting.TextStream
          Dim CopyOfErr As String

1         On Error GoTo ErrHandler
2         If FSO Is Nothing Then Set FSO = New FileSystemObject
3         Set T = FSO.GetFile(File1).OpenAsTextStream(ForReading, Format)
4         Contents1 = T.ReadAll
5         T.Close
6         Set T = FSO.GetFile(File2).OpenAsTextStream(ForReading, Format)
7         Contents2 = T.ReadAll
8         T.Close
9         TextFilesIdentical = Contents1 = Contents2
10        Exit Function
ErrHandler:
11        CopyOfErr = "#TextFilesIdentical (line " & CStr(Erl) + "): " & Err.Description & "!"
12        If Not T Is Nothing Then T.Close
13        Throw CopyOfErr
End Function


