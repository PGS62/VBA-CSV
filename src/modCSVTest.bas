Attribute VB_Name = "modCSVTest"
Option Explicit

Private Const m_FolderOriginals = "c:\temp\CSVTest\Originals"
Private Const m_FolderReadAndRewrite = "c:\temp\CSVTest\ReadAndWritten"
Private Const m_FolderSpeedTest = "C:\Temp\CSVTest\SpeedTest"

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CSVSpeedTest
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
    Dim data, DataReread
    Dim FileName As String
    Dim i As Long
    Dim NumCols As Long
    Dim NumRows As Long
    Dim OS As String
    Dim SmallFileName As String
    Dim t1 As Double, t2 As Double

    On Error GoTo ErrHandler

    NumRows = 10000
    NumCols = 100
    OS = "Windows"

    ThrowIfError CreatePath(m_FolderSpeedTest)
    Debug.Print String(100, "=")
    Debug.Print "Time of test = " + _
        Format(Now, "dd-mmm-yyyy hh:mm:ss") + " Computer = " + Environ("COMPUTERNAME")

    'Doubles only, cast back to doubles
    data = RandomDoubles(NumRows, NumCols)
    FileName = NameThatFile(m_FolderSpeedTest, OS, NumRows, NumCols, "Doubles", False, False)
    ThrowIfError CSVWrite(FileName, data, True, , , , False, OS, False)
    t1 = sElapsedTime
    DataReread = ThrowIfError(CSVRead_V2(FileName, True))
    t2 = sElapsedTime
    Debug.Print CStr(t2 - t1) + " seconds to read 1 file containing random doubles " + _
        Format(NumRows, "###,##0") + " rows, " + Format(NumCols, "###,##0") + " cols. " '+ _
        "File size = " + Format(sFileInfo(FileName, "size"), "###,##0") + " bytes."

    '10-character strings, unquoted
    data = sFill("abcdefghij", NumRows, NumCols)
    FileName = NameThatFile(m_FolderSpeedTest, OS, NumRows, NumCols, "10-char-strings-unquoted", False, False)
    ThrowIfError CSVWrite(FileName, data, False, , , , , OS)
    t1 = sElapsedTime
    DataReread = ThrowIfError(CSVRead_V2(FileName, False))
    t2 = sElapsedTime
    Debug.Print CStr(t2 - t1) + " seconds to read 1 file containing UNquoted 10-char strings " + _
        Format(NumRows, "###,##0") + " rows, " + _
        Format(NumCols, "###,##0") + " cols. File size = " ' + _
        Format(sFileInfo(FileName, "size"), "###,##0") + " bytes."

    '10-character strings...
    data = sFill("abcdefghij", NumRows, NumCols)
    FileName = NameThatFile(m_FolderSpeedTest, OS, NumRows, NumCols, "10-char-strings", False, False)
    ThrowIfError CSVWrite(FileName, data, , , , , , OS, False)
    t1 = sElapsedTime
    DataReread = ThrowIfError(CSVRead_V2(FileName, False))
    t2 = sElapsedTime
    Debug.Print CStr(t2 - t1) + " seconds to read 1 file containing QUOTED 10-char strings " + _
        Format(NumRows, "###,##0") + " rows, " + _
        Format(NumCols, "###,##0") + " cols. File size = " ' + _
        Format(sFileInfo(FileName, "size"), "###,##0") + " bytes."

    '10-character strings ALL with linefeeds
    data = sFill("abcd+" + vbCrLf + "efghi", NumRows, NumCols)
    FileName = NameThatFile(m_FolderSpeedTest, OS, NumRows, NumCols, "10-char-strings-with-line-feeds", False, False)
    ThrowIfError CSVWrite(FileName, data, , , , , , OS, False)
    t1 = sElapsedTime
    DataReread = ThrowIfError(CSVRead_V2(FileName))
    t2 = sElapsedTime
    Debug.Print CStr(t2 - t1) + " seconds to read 1 file containing 10-char strings, all with line-feeds " + _
        Format(NumRows, "###,##0") + " rows, " + Format(NumCols, "###,##0") + " cols. File size = " ' + _
        Format(sFileInfo(FileName, "size"), "###,##0") + " bytes."

    'Write and read many files
    
    'Create files
    t1 = sElapsedTime()
    For i = 1 To NumFilesToReadAndWrite
        SmallFileName = NameThatFile(m_FolderSpeedTest, OS, NumRowsSmall, NumColsSmall, Format(i, "0000"), False, False)
        data = RandomDoubles(NumRowsSmall, NumColsSmall)
        ThrowIfError CSVWrite(SmallFileName, data)
    Next i
    t2 = sElapsedTime()
    Debug.Print CStr(t2 - t1) + " seconds to write " + CStr(NumFilesToReadAndWrite) + " files. " + _
        "Each file has " + CStr(NumRowsSmall) + " rows and " + CStr(NumColsSmall) + " columns."
    
    'Read them back
    t1 = sElapsedTime()
    For i = 1 To NumFilesToReadAndWrite
        SmallFileName = NameThatFile(m_FolderSpeedTest, OS, NumRowsSmall, NumColsSmall, Format(i, "0000"), False, False)
        data = ThrowIfError(CSVRead_V2(SmallFileName, True, ","))
    Next i
    t2 = sElapsedTime()
    Debug.Print CStr(t2 - t1) + " seconds to read " + CStr(NumFilesToReadAndWrite) + " files. " + _
        "Each file has " + CStr(NumRowsSmall) + " rows and " + CStr(NumColsSmall) + " columns."
    Debug.Print String(100, "=")

    Exit Sub
ErrHandler:
    MsgBox "#CSVSpeedTest: " & Err.Description & "!", vbCritical
End Sub

Function NameThatFile(Folder As String, ByVal OS As String, NumRows As Long, NumCols As Long, ExtraInfo As String, Unicode As Boolean, Ragged As Boolean)
    NameThatFile = (Folder & "\" & OS & "_" & Format(NumRows, "0000") & "_x_" & Format(NumCols, "000") & "_" & ExtraInfo & IIf(Unicode, "_Unicode", "_Ascii") & IIf(Ragged, "_Ragged", "_NotRagged") & ".csv")
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CSVRoundTripTestMulti
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
    Dim data As Variant
    Dim k As Long
    Dim DateFormat As Variant
    Dim AllowLineFeed As Variant
    Dim Ragged As Variant
    Dim ExtraInfo As String
    Dim EOL As String

    On Error GoTo ErrHandler

    ThrowIfError CreatePath(m_FolderOriginals)
    ThrowIfError CreatePath(m_FolderReadAndRewrite)

    For Each OS In Array("Windows", "Unix", "Mac")
        EOL = IIf(OS = "Windows", vbCrLf, IIf(OS = "Unix", vbLf, vbCr))
    
        For Each Unicode In Array(True, False)
            For Each Ragged In Array(True, False)
                For Each NRows In Array(1, 2, 20)
                    For Each NCols In Array(1, 2, 10)
              
                        'For Variants we need to vary AllowLineFeed and DateFormat
                        For Each AllowLineFeed In Array(True, False)
                            For Each DateFormat In Array("mmm-dd-yyyy", "dd-mmm-yyyy", "yyyy-mm-dd")
                                data = RandomVariants(CLng(NRows), CLng(NCols), CBool(AllowLineFeed), CBool(Unicode), EOL)
                                ExtraInfo = "RandomVariants" & IIf(AllowLineFeed, "WithLineFeed", "")
                                CSVRoundTripTest CStr(OS), data, CStr(DateFormat), CBool(Unicode), CStr(OS), vbTab, CBool(Ragged), ExtraInfo
                            Next DateFormat
                        Next AllowLineFeed

                        'For Dates, we need to vary DateFormat
                        For Each DateFormat In Array("mmm-dd-yyyy", "dd-mmm-yyyy", "yyyy-mm-dd")
                            data = RandomDates(CLng(NRows), CLng(NCols))
                            ExtraInfo = "RandomDates"
                            CSVRoundTripTest CStr(OS), data, CStr(DateFormat), CBool(Unicode), CStr(OS), vbTab, CBool(Ragged), ExtraInfo
                        Next DateFormat

                        'For Strings, we need to vary AllowLineFeed
                        For Each AllowLineFeed In Array(True, False)
                            data = RandomStrings(CLng(NRows), CLng(NCols), CBool(Unicode), CBool(AllowLineFeed), EOL)
                            ExtraInfo = IIf(AllowLineFeed, "RandomStringsWithLineFeeds", "RandomStrings")
                            CSVRoundTripTest CStr(OS), data, CStr(DateFormat), CBool(Unicode), CStr(OS), vbTab, CBool(Ragged), ExtraInfo
                        Next AllowLineFeed

                        For k = 1 To 4
                            If k = 1 Then
                                data = RandomBooleans(CLng(NRows), CLng(NCols))
                                ExtraInfo = "RandomBooleans"
                            ElseIf k = 2 Then
                                data = RandomDoubles(CLng(NRows), CLng(NCols))
                                ExtraInfo = "RandomDoubles"
                            ElseIf k = 3 Then
                                data = RandomErrorValues(CLng(NRows), CLng(NCols))
                                ExtraInfo = "RandomErrorValues"
                            ElseIf k = 4 Then
                                data = RandomLongs(CLng(NRows), CLng(NCols))
                                ExtraInfo = "RandomLongs"
                            End If
                            CSVRoundTripTest CStr(OS), data, CStr(DateFormat), CBool(Unicode), CStr(OS), vbTab, CBool(Ragged), ExtraInfo
                        Next k
                    Next NCols
                Next NRows
            Next Ragged
        Next Unicode
    Next OS

    Exit Sub
ErrHandler:
    Throw "#AnotherWay: " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CSVRoundTripTest
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
Function CSVRoundTripTest(OS As String, ByVal data As Variant, DateFormat As String, Unicode As Boolean, EOL As String, Delimiter As String, Ragged As Boolean, ExtraInfo As String)

    Dim DataReadBack

    On Error GoTo ErrHandler
    Dim FileName1 As String
    Dim FileName2 As String
    Dim NR As Long
    Dim NC As Long

    NR = sNRows(data)
    NC = sNCols(data)

    FileName1 = NameThatFile(m_FolderOriginals, OS, NR, NC, ExtraInfo, CBool(Unicode), CBool(Ragged))
    FileName2 = NameThatFile(m_FolderReadAndRewrite, OS, NR, NC, ExtraInfo, CBool(Unicode), CBool(Ragged))

    If Ragged Then data = MakeArrayRagged(data)

    ThrowIfError CSVWrite(FileName1, data, True, DateFormat, , Delimiter, Unicode, EOL, Ragged)

    'The Call to CSVRead has to infer both Unicode and EOL
    DataReadBack = CSVRead_V2(FileName1, True, , DateFormat, , , , , , Empty)

    ThrowIfError CSVWrite(FileName2, data, True, DateFormat, , Delimiter, Unicode, EOL, Ragged)

    If Not TextFilesIdentical(FileName1, FileName2, IIf(Unicode, TristateTrue, TristateFalse)) Then
        Stop
    End If

    Exit Function
ErrHandler:
    Throw "#CSVRoundTripTest: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : MakeArrayRagged
' Purpose    : For each row of an array choose random number n less than number of cols and make the n right most columns empty
'              also guarantee that one row will not have an empty right-most column.
' -----------------------------------------------------------------------------------------------------------------------
Private Function MakeArrayRagged(data)

    Dim NR As Long, NC As Long
    Dim i As Long, j As Long
    Dim RowToLeaveUnchanged As Long

    On Error GoTo ErrHandler
    NR = sNRows(data)
    NC = sNCols(data)
    RowToLeaveUnchanged = CLng(0.5 + Rnd() * (NR))

    For i = 1 To NR
        If i = RowToLeaveUnchanged Then
            If IsEmpty(data(i, NC)) Then
                data(i, NC) = "Not empty!"
            End If
        Else
            For j = CLng(0.5 + Rnd() * NC) To NC
                data(i, j) = Empty
            Next
        End If
    Next
    MakeArrayRagged = data

    Exit Function
ErrHandler:
    Throw "#MakeArrayRagged: " & Err.Description & "!"
End Function

Function TextFilesIdentical(File1 As String, File2 As String, Format As Scripting.Tristate) As Boolean
    Static FSO As FileSystemObject

    Dim Contents1 As String
    Dim Contents2 As String
    Dim T As Scripting.TextStream
    Dim CopyOfErr As String

    On Error GoTo ErrHandler
    If FSO Is Nothing Then Set FSO = New FileSystemObject
    Set T = FSO.GetFile(File1).OpenAsTextStream(ForReading, Format)
    Contents1 = T.ReadAll
    T.Close
    Set T = FSO.GetFile(File2).OpenAsTextStream(ForReading, Format)
    Contents2 = T.ReadAll
    T.Close
    TextFilesIdentical = Contents1 = Contents2
    Exit Function
ErrHandler:
    CopyOfErr = "#TextFilesIdentical: " & Err.Description & "!"
    If Not T Is Nothing Then T.Close
    Throw CopyOfErr
End Function
