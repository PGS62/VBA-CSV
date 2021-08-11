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
    'DataReread = ThrowIfError(CSVRead_V2(FileName, True))
    t2 = sElapsedTime
    Debug.Print CStr(t2 - t1) + " seconds to read 1 file containing random doubles " + _
        Format(NumRows, "###,##0") + " rows, " + Format(NumCols, "###,##0") + " cols. " '+ _
        "File size = " + Format(sFileInfo(FileName, "size"), "###,##0") + " bytes."

    '10-character strings, unquoted
    data = sFill("abcdefghij", NumRows, NumCols)
    FileName = NameThatFile(m_FolderSpeedTest, OS, NumRows, NumCols, "10-char-strings-unquoted", False, False)
    ThrowIfError CSVWrite(FileName, data, False, , , , , OS)
    t1 = sElapsedTime
    'DataReread = ThrowIfError(CSVRead_V2(FileName, False))
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
    'DataReread = ThrowIfError(CSVRead_V2(FileName, False))
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
    'DataReread = ThrowIfError(CSVRead_V2(FileName))
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
        'data = ThrowIfError(CSVRead_V2(SmallFileName, True, ","))
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
    'DataReadBack = CSVRead_V2(FileName1, True, , DateFormat, , , , , , Empty)

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




Function CSVWriteMaybeRagged(FileName As String, ByVal data As Variant, Optional QuoteAllStrings As Boolean = True, _
        Optional DateFormat As String = "yyyy-mm-dd", Optional DateTimeFormat As String = "yyyy-mm-dd hh:mm:ss", _
        Optional Delimiter As String = ",", Optional Unicode As Boolean, Optional ByVal EOL As String = vbCrLf, Optional Ragged As Boolean = False)

          Dim FSO As Scripting.FileSystemObject
          Dim i As Long
          Dim j As Long
          Dim k As Long
          
          Dim OneLine() As String
          Dim OneLineJoined As String
          Dim T As TextStream
          Dim EOLIsWindows As Boolean
          Const DQ = """"
          
          'Const Err_Delimiter = "Delimiter must be one character, and cannot be double quote or line feed characters"
          Const Err_Delimiter = "Delimiter must not contain double quote or line feed characters"

1         On Error GoTo ErrHandler

2         EOL = OStoEOL(EOL, "EOL")
3         EOLIsWindows = EOL = vbCrLf

          'If Len(Delimiter) <> 1 Or Delimiter = """" Or Delimiter = vbCr Or Delimiter = vbLf Then
4         If InStr(Delimiter, DQ) > 0 Or InStr(Delimiter, vbLf) > 0 Or InStr(Delimiter, vbCr) > 0 Then
5             Throw Err_Delimiter
6         End If

7         If TypeName(data) = "Range" Then
              'Preserve elements of type Date by using .Value, not .Value2
8             data = data.value
9         End If
10        Force2DArray data 'Coerce 0-dim & 1-dim to 2-dims.

11        Set FSO = New FileSystemObject
12        Set T = FSO.CreateTextFile(FileName, True, Unicode)

13        ReDim OneLine(LBound(data, 2) To UBound(data, 2))

14        For i = LBound(data) To UBound(data)
15            For j = LBound(data, 2) To UBound(data, 2)
16                OneLine(j) = Encode(data(i, j), QuoteAllStrings, DateFormat, DateTimeFormat)
17            Next j
18            OneLineJoined = VBA.Join(OneLine, Delimiter)

              'If writing in "Ragged" style, remove terminating delimiters
19            If Ragged Then
20                For k = Len(OneLineJoined) To 1 Step -1
21                    If Mid(OneLineJoined, k, 1) <> Delimiter Then Exit For
22                Next k
23                If k < Len(OneLineJoined) Then
24                    OneLineJoined = Left(OneLineJoined, k)
25                End If
26            End If
              
27            WriteLineWrap T, OneLineJoined, EOLIsWindows, EOL, Unicode
28        Next i

          'Quote from https://tools.ietf.org/html/rfc4180#section-2 : _
          "The last record in the file may or may not have an ending line break." _
           We follow Excel (File save as CSV) and *do* put a line break after the last line.

29        T.Close: Set T = Nothing: Set FSO = Nothing
30        CSVWriteMaybeRagged = FileName
31        Exit Function
ErrHandler:
32        CSVWriteMaybeRagged = "#CSVWriteMaybeRagged (line " & CStr(Erl) + "): " & Err.Description & "!"
33        If Not T Is Nothing Then Set T = Nothing: Set FSO = Nothing

End Function


' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : WriteLineWrap
' Purpose    : Wrapper to TextStream.Write(Line) to give more informative error message than "invalid procedure call or
'              argument" if the error is caused by attempting to write Unicode characters to ascii file
' -----------------------------------------------------------------------------------------------------------------------
Private Sub WriteLineWrap(T As TextStream, text As String, EOLIsWindows As Boolean, EOL As String, Unicode As Boolean)
          Dim ErrDesc As String
          Dim ErrNum As Long
          Dim i As Long
          Dim ErrLN As String

1         On Error GoTo ErrHandler
2         If EOLIsWindows Then
3             T.WriteLine text
4         Else
5             T.Write text
6             T.Write EOL
7         End If

8         Exit Sub

ErrHandler:
9         ErrNum = Err.Number
10        ErrDesc = Err.Description
11        ErrLN = CStr(Erl)
12        If Not Unicode Then
13            If ErrNum = 5 Then
14                For i = 1 To Len(text)
15                    If AscW(Mid(text, i, 1)) > 255 Then
16                        ErrDesc = "Data contains unicode characters (first found has code " & CStr(AscW(Mid(text, i, 1))) & ") which cannot be written to an ascii file. Set argument Unicode to True"
17                        Exit For
18                    End If
19                Next i
20            End If
21        End If
22        Throw "#WriteLineWrap (line " & ErrLN + "): " & ErrDesc & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Encode
' Purpose    : Encode arbitrary value as a string, sub-routine of CSVWrite.
' -----------------------------------------------------------------------------------------------------------------------
Private Function Encode(x As Variant, QuoteAllStrings As Boolean, DateFormat As String, DateTimeFormat As String) As String
    Const DQ = """"
    Const DQ2 = """"""

    On Error GoTo ErrHandler
    Select Case VarType(x)

        Case vbString
            If InStr(x, DQ) > 0 Then
                Encode = DQ + Replace(x, DQ, DQ2) + DQ
            ElseIf QuoteAllStrings Then
                Encode = DQ + x + DQ
            ElseIf InStr(x, vbCr) > 0 Then
                Encode = DQ + x + DQ
            ElseIf InStr(x, vbLf) > 0 Then
                Encode = DQ + x + DQ
            ElseIf InStr(x, ",") > 0 Then
                Encode = DQ + x + DQ
            Else
                Encode = x
            End If
        Case vbBoolean, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbLongLong, vbEmpty
            Encode = CStr(x)
        Case vbDate
            If CLng(x) = CDbl(x) Then
                Encode = Format$(x, DateFormat)
            Else
                Encode = Format$(x, DateTimeFormat)
            End If
        Case vbNull
            Encode = "NULL"
        Case vbError
            Select Case CStr(x) 'Editing this case statement? Edit also its inverse - CastToError
                Case "Error 2000"
                    Encode = "#NULL!"
                Case "Error 2007"
                    Encode = "#DIV/0!"
                Case "Error 2015"
                    Encode = "#VALUE!"
                Case "Error 2023"
                    Encode = "#REF!"
                Case "Error 2029"
                    Encode = "#NAME?"
                Case "Error 2036"
                    Encode = "#NUM!"
                Case "Error 2042"
                    Encode = "#N/A"
                Case "Error 2043"
                    Encode = "#GETTING_DATA!"
                Case "Error 2045"
                    Encode = "#SPILL!"
                Case "Error 2046"
                    Encode = "#CONNECT!"
                Case "Error 2047"
                    Encode = "#BLOCKED!"
                Case "Error 2048"
                    Encode = "#UNKNOWN!"
                Case "Error 2049"
                    Encode = "#FIELD!"
                Case "Error 2050"
                    Encode = "#CALC!"
                Case Else
                    Encode = CStr(x)        'should never hit this line...
            End Select
        Case Else
            Throw "Cannot convert variant of type " + TypeName(x) + " to String"
    End Select
    Exit Function
ErrHandler:
    Throw "#Encode: " & Err.Description & "!"
End Function


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

