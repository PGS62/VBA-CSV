Attribute VB_Name = "modCSVReadWrite"

' VBA-CSV

' Copyright (C) 2021 - PGS62 (https://github.com/PGS62/VBA-CSV )
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/VBA-CSV#readme

Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RegisterCSVRead
' Purpose    : Register the function CSVRead with the Excel function wizard. Suggest this function is called from a
'              WorkBook_Open event.
' -----------------------------------------------------------------------------------------------------------------------
Sub RegisterCSVRead()
    Const FnDesc = "Returns the contents of a comma-separated file on disk as an array."
    Dim ArgDescs() As String
    ReDim ArgDescs(1 To 10)
    ArgDescs(1) = "The full name of the file, including the path."
    ArgDescs(2) = "TRUE to convert Numbers, Dates, Booleans and Excel Errors into their typed values, or FALSE to leave as strings. For more control enter a string containing the letters N, D, B, E e.g. ""NB"" to convert just numbers and Booleans, not dates or errors."
    ArgDescs(3) = "Delimiter string. Defaults to the first instance of comma, tab, semi-colon, colon or pipe found outside quoted regions. Enter FALSE to  see the file's raw contents as would be displayed in a text editor. Delimiter may have more than one character."
    ArgDescs(4) = "The format of dates in the file such as D-M-Y, M-D-Y or Y/M/D. If omitted a value is read from Windows regional settings. Repeated D's (or M's or Y's) are equivalent to single instances, so that d-m-y and DD-MMM-YYYY are equivalent."
    ArgDescs(5) = "The row in the file at which reading starts. Optional and defaults to 1 to read from the first row."
    ArgDescs(6) = "The column in the file at which reading starts. Optional and defaults to 1 to read from the first column."
    ArgDescs(7) = "The number of rows to read from the file. If omitted (or zero), all rows from SkipToRow to the end of the file are read."
    ArgDescs(8) = "The number of columns to read from the file. If omitted (or zero), all columns from SkipToCol are read."
    ArgDescs(9) = "Enter TRUE if the file is unicode, FALSE if the file is ascii. Omit to infer from the file's contents."
    ArgDescs(10) = "The character that represents a decimal point. If omitted, then the value from Windows regional settings is used."
    Application.MacroOptions "CSVRead", FnDesc, , , , , , , , , ArgDescs
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RegisterCSVWrite
' Purpose    : Register the function CSVWrite with the Excel function wizard. Suggest this function is called from a
'              WorkBook_Open event.
' -----------------------------------------------------------------------------------------------------------------------
Sub RegisterCSVWrite()
    Const FnDesc = "Creates a comma-separated file on disk containing Data. Any existing file of the same name is overwritten. If successful, the function returns FileName, otherwise an ""error string"" (starts with #, ends with !) describing what went wrong."
    Dim ArgDescs() As String
    ReDim ArgDescs(1 To 8)
    ArgDescs(1) = "The full name of the file, including the path."
    ArgDescs(2) = "An array of data. Elements may be strings, numbers, dates, Booleans, empty, Excel errors or null values."
    ArgDescs(3) = "If TRUE (the default) then all strings in Data are quoted before being written to file. If FALSE only strings containing Delimiter, line feed, carriage return or double quote are quoted. Double quotes are always escaped by another double quote."
    ArgDescs(4) = "A format string that determine how dates, including cells formatted as dates, appear in the file. If omitted, defaults to ""yyyy-mm-dd""."
    ArgDescs(5) = "A format string that determines how dates with non-zero time part appear in the file. If omitted defaults to ""yyyy-mm-dd hh:mm:ss"".The companion function CVSRead is not currently capable of interpreting fields written in DateTime format."
    ArgDescs(6) = "The delimiter string, if omitted defaults to a comma. Delimiter may have more than one character."
    ArgDescs(7) = "If FALSE (the default) the file written will be ascii. If TRUE the file written will be unicode. An error will result if Unicode is FALSE but Data contains strings with unicode characters."
    ArgDescs(8) = "Enter the required line ending character as ""Windows"" (or ascii 13 plus ascii 10), or ""Unix"" (or ascii 10) or ""Mac"" (or ascii 13). If omitted defaults to ""Windows""."
    Application.MacroOptions "CSVWrite", FnDesc, , , , , , , , , ArgDescs
End Sub

'---------------------------------------------------------------------------------------------------------
' Procedure : CSVRead
' Purpose   : Returns the contents of a comma-separated file on disk as an array.
' Arguments
' FileName  : The full name of the file, including the path.
' ConvertTypes: TRUE to convert Numbers, Dates, Booleans and Errors into their typed values, or FALSE to
'             leave as strings. Or letters N, D, B, E, Q. E.g. "BN" to convert just
'             Booleans and numbers. Q indicates quoted strings should retain their quotes.
' Delimiter : Delimiter string. Defaults to the first instance of comma, tab, semi-colon, colon or pipe
'             found outside quoted regions. Enter FALSE to  see the file's raw contents as
'             would be displayed in a text editor. Delimiter may have more than one
'             character.
' DateFormat: The format of dates in the file such as D-M-Y, M-D-Y or Y/M/D. If omitted a value is read
'             from Windows regional settings. Repeated D's (or M's or Y's) are equivalent
'             to single instances, so that d-m-y and DD-MMM-YYYY are equivalent.
' SkipToRow : The row in the file at which reading starts. Optional and defaults to 1 to read from the
'             first row.
' SkipToCol : The column in the file at which reading starts. Optional and defaults to 1 to read from
'             the first column.
' NumRows   : The number of rows to read from the file. If omitted (or zero), all rows from SkipToRow to
'             the end of the file are read.
' NumCols   : The number of columns to read from the file. If omitted (or zero), all columns from
'             SkipToCol are read.
' Unicode   : Enter TRUE if the file is unicode, FALSE if the file is ascii. Omit to infer from the
'             file's contents.
' DecimalSeparator: The character that represents a decimal point. If omitted, then the value from Windows
'             regional settings is used.
'
' Notes     : See also companion function CSVWrite.
'
'             The function handles all csv files that conform to the standards described in
'             RFC4180  https://www.rfc-editor.org/rfc/rfc4180.txt including files with
'             quoted fields.
'
'             In addition the function handles files which break some of those standards:
'             * Not all lines of the file need have the same number of fields. The function
'             "pads" with Empty values.
'             * Fields which start with a double quote but do not end with a double quote
'             are handled by being returned unchanged. Necessarily such fields have an even
'             number of double quotes, or otherwise the field will be treated as the last
'             field in the file.
'             * The standard states that csv files should have Windows-style line endings,
'             but the function supports files with Windows, Unix and (old) Mac line
'             endings. Files may also have mixed line endings.
'---------------------------------------------------------------------------------------------------------
Public Function CSVRead(FileName As String, Optional ConvertTypes As Variant = False, Optional ByVal Delimiter As Variant, _
    Optional DateFormat As String, Optional ByVal SkipToRow As Long = 1, Optional ByVal SkipToCol As Long = 1, _
    Optional ByVal NumRows As Long = 0, Optional ByVal NumCols As Long = 0, _
    Optional ByVal Unicode As Variant, Optional DecimalSeparator As String = vbNullString)
Attribute CSVRead.VB_Description = "Returns the contents of a comma-separated file on disk as an array."
Attribute CSVRead.VB_ProcData.VB_Invoke_Func = " \n14"

    Const DQ = """"
    Const DQ2 = """"""
    Const Err_Delimiter = "Delimiter character must be passed as a string, FALSE for no delimiter. Omit to guess from file contents"
    Const Err_FileEmpty = "File is empty"
    Const Err_FileIsUniCode = "Unicode must be passed as TRUE or FALSE. Omit to infer from file contents"
    Const Err_NumCols = "NumCols must be positive to read a given number of columns, or zero or omitted to read all columns from SkipToCol to the maximum column encountered."
    Const Err_NumRows = "NumRows must be positive to read a given number of rows, or zero or omitted to read all rows from SkipToRow to the end of the file."
    Const Err_Seps = "DecimalSeparator must be different from Delimiter"
    Const Err_SkipToCol = "SkipToCol must be at least 1."
    Const Err_SkipToRow = "SkipToRow must be at least 1."
    Dim AnyConversion As Boolean
    Dim ColIndexes() As Long
    Dim CSVContents As String
    Dim DateOrder As Long
    Dim DateSeparator As String
    Dim F As Scripting.File
    Dim FSO As New Scripting.FileSystemObject
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim Lengths() As Long
    Dim NotDelimited As Boolean
    Dim NumColsFound As Long
    Dim NumColsInReturn As Long
    Dim NumFields As Long
    Dim NumRowsFound As Long
    Dim NumRowsInReturn As Long
    Dim QuoteCounts() As Long
    Dim RemoveQuotes As Boolean
    Dim ReturnArray() As Variant
    Dim RowIndexes() As Long
    Dim SepsStandard As Boolean
    Dim ShowBooleansAsBooleans As Boolean
    Dim ShowDatesAsDates As Boolean
    Dim ShowErrorsAsErrors As Boolean
    Dim ShowMissingAsNullString As Boolean
    Dim ShowNumbersAsNumbers As Boolean
    Dim Starts() As Long
    Dim strDelimiter As String
    Dim SysDateOrder As Long
    Dim SysDateSeparator As String
    Dim SysDecimalSeparator As String
    Dim T As Scripting.TextStream
    Dim ThisField As String
    
    On Error GoTo ErrHandler

    'Parse and validate inputs...
    If IsEmpty(Unicode) Or IsMissing(Unicode) Then
        Unicode = IsUnicodeFile(FileName)
    ElseIf VarType(Unicode) <> vbBoolean Then
        Throw Err_FileIsUniCode
    End If

    If VarType(Delimiter) = vbBoolean Then
        If Not Delimiter Then
            NotDelimited = True
        Else
            Throw Err_Delimiter
        End If
    ElseIf VarType(Delimiter) = vbString Then
        strDelimiter = Delimiter
    ElseIf IsEmpty(Delimiter) Or IsMissing(Delimiter) Then
        strDelimiter = InferDelimiter(FileName, CBool(Unicode))
    Else
        Throw Err_Delimiter
    End If

    ParseConvertTypes ConvertTypes, ShowNumbersAsNumbers, _
        ShowDatesAsDates, ShowBooleansAsBooleans, ShowErrorsAsErrors, RemoveQuotes

    If ShowNumbersAsNumbers Then
        If ((DecimalSeparator = Application.DecimalSeparator) Or DecimalSeparator = vbNullString) Then
            SepsStandard = True
        ElseIf DecimalSeparator = strDelimiter Then
            Throw Err_Seps
        End If
    End If

    If ShowDatesAsDates Then
        ParseDateFormat DateFormat, DateOrder, DateSeparator
        SysDateOrder = Application.International(xlDateOrder)
        SysDateSeparator = Application.International(xlDateSeparator)
    End If

    If SkipToRow < 1 Then Throw Err_SkipToRow
    If SkipToCol < 1 Then Throw Err_SkipToCol
    If NumRows < 0 Then Throw Err_NumRows
    If NumCols < 0 Then Throw Err_NumCols
    'End of input validation
          
    If NotDelimited Then
        CSVRead = ShowTextFile(FileName, SkipToRow, NumRows, CBool(Unicode))
        Exit Function
    End If
          
    Set F = FSO.GetFile(FileName)
    Set T = F.OpenAsTextStream(ForReading, IIf(Unicode, TristateTrue, TristateFalse))

    If T.AtEndOfStream Then
        T.Close: Set T = Nothing: Set F = Nothing
        Throw Err_FileEmpty
    End If
          
    If SkipToRow = 1 And NumRows = 0 Then
        CSVContents = T.ReadAll
        T.Close: Set T = Nothing: Set F = Nothing
        Call ParseCSVContents(CSVContents, DQ, strDelimiter, SkipToRow, NumRows, NumRowsFound, NumColsFound, NumFields, _
            Starts, Lengths, RowIndexes, ColIndexes, QuoteCounts)
    Else
        CSVContents = ParseCSVContents(T, DQ, strDelimiter, SkipToRow, NumRows, NumRowsFound, NumColsFound, NumFields, _
            Starts, Lengths, RowIndexes, ColIndexes, QuoteCounts)
        T.Close
    End If
    
    'Useful for debugging...
    '   CSVRead = sArrayRange(sArrayStack(NumRowsFound, NumColsFound, NumFields, strDelimiter), sArrayTranspose(Starts), _
    sArrayTranspose(Lengths), sArrayTranspose(RowIndexes), sArrayTranspose(ColIndexes), sArrayTranspose(QuoteCounts))
    '   Exit Function
    
    AnyConversion = ShowNumbersAsNumbers Or ShowDatesAsDates Or _
        ShowBooleansAsBooleans Or ShowErrorsAsErrors Or (Not ShowMissingAsNullString)
        
    If NumCols = 0 Then
        NumColsInReturn = NumColsFound - SkipToCol + 1
        If NumColsInReturn <= 0 Then
            Throw "SkipToCol (" + CStr(SkipToCol) + ") exceeds the number of columns in the file (" + CStr(NumColsFound) + ")"
        End If
    Else
        NumColsInReturn = NumCols
    End If
    If NumRows = 0 Then
        NumRowsInReturn = NumRowsFound
    Else
        NumRowsInReturn = NumRows
    End If
        
    ReDim ReturnArray(1 To NumRowsInReturn, 1 To NumColsInReturn)
        
    For k = 1 To NumFields
        i = RowIndexes(k)
        j = ColIndexes(k) - SkipToCol + 1
        If j >= 1 And j <= NumColsInReturn Then
            If QuoteCounts(k) = 0 Or Not RemoveQuotes Then
                ThisField = Mid(CSVContents, Starts(k), Lengths(k))
            ElseIf Mid(CSVContents, Starts(k), 1) = DQ And Mid(CSVContents, Starts(k) + Lengths(k) - 1, 1) = DQ Then
                ThisField = Mid(CSVContents, Starts(k) + 1, Lengths(k) - 2)
                If QuoteCounts(k) > 2 Then
                    ThisField = Replace(ThisField, DQ2, DQ)
                End If
            Else 'Field which does not start with quote but contains quotes, so not RFC4180 compliant. We do not replace DQ2 by DQ in this case.
                ThisField = Mid(CSVContents, Starts(k), Lengths(k))
            End If
        
            If AnyConversion And QuoteCounts(k) = 0 Then
                ReturnArray(i, j) = CastToVariant(ThisField, _
                    ShowNumbersAsNumbers, SepsStandard, DecimalSeparator, SysDecimalSeparator, _
                    ShowDatesAsDates, DateOrder, DateSeparator, SysDateOrder, SysDateSeparator, _
                    ShowBooleansAsBooleans, ShowErrorsAsErrors)
            Else
                ReturnArray(i, j) = ThisField
            End If
        End If
    Next k

    CSVRead = ReturnArray

    Exit Function

ErrHandler:
    CSVRead = "#CSVRead: " & Err.Description & "!"
    If Not T Is Nothing Then T.Close
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseConvertTypes
' Purpose    : Parse the input ConvertTypes to set five Boolean flags which are passed by reference
' Parameters :
'  ConvertTypes        :
'  ShowNumbersAsNumbers  : Should fields in the file that look like numbers be returned as Numbers? (Doubles)
'  ShowDatesAsDates      : Should fields in the file that look like dates with the specified DateFormat be returned as Dates?
'  ShowBooleansAsBooleans: Should fields in the file that are TRUE or FALSE (case insensitive) be returned as Booleans?
'  ShowErrorsAsErrors    : Should fields in the file that look like Excel errors (#N/A #REF! etc) be returned as errors?
'  RemoveQuotes          : Should quoted fields be unquoted?
' -----------------------------------------------------------------------------------------------------------------------
Private Sub ParseConvertTypes(ByVal ConvertTypes As Variant, ByRef ShowNumbersAsNumbers As Boolean, _
    ByRef ShowDatesAsDates As Boolean, ByRef ShowBooleansAsBooleans As Boolean, _
    ByRef ShowErrorsAsErrors As Boolean, ByRef RemoveQuotes As Boolean)

    Const Err_ConvertTypes = "ConvertTypes must be TRUE (convert all types), FALSE (no conversion) or a string " & _
        "containing letters: 'N' to show numbers as numbers, 'D' to show dates as dates, 'B' to show Booleans " & _
        "as Booleans, `E` to show Excel errors as errors, Q for quoted fields to retain their quotes."
    Dim i As Long

    On Error GoTo ErrHandler
    
    If TypeName(ConvertTypes) = "Range" Then
        ConvertTypes = ConvertTypes.value
    End If

    If VarType(ConvertTypes) = vbBoolean Then
        If ConvertTypes Then
            ShowNumbersAsNumbers = True
            ShowDatesAsDates = True
            ShowBooleansAsBooleans = True
            ShowErrorsAsErrors = True
            RemoveQuotes = True
        Else
            ShowNumbersAsNumbers = False
            ShowDatesAsDates = False
            ShowBooleansAsBooleans = False
            ShowErrorsAsErrors = False
            RemoveQuotes = True
        End If
    ElseIf VarType(ConvertTypes) = vbString Then
        ShowNumbersAsNumbers = False
        ShowDatesAsDates = False
        ShowBooleansAsBooleans = False
        ShowErrorsAsErrors = False
        RemoveQuotes = True
        For i = 1 To Len(ConvertTypes)
            Select Case UCase(Mid(ConvertTypes, i, 1))
                Case "N"
                    ShowNumbersAsNumbers = True
                Case "D"
                    ShowDatesAsDates = True
                Case "B", "L" 'Booleans aka Logicals
                    ShowBooleansAsBooleans = True
                Case "E"
                    ShowErrorsAsErrors = True
                Case "Q"
                    RemoveQuotes = False
                Case Else
                    Throw Err_ConvertTypes + " Found unrecognised character '" + Mid(ConvertTypes, i, 1) + "'"
            End Select
        Next i
    Else
        Throw Err_ConvertTypes
    End If

    Exit Sub
ErrHandler:
    Throw "#ParseConvertTypes: " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Min4
' Purpose    : Returns the minimum of four inputs and an indicator of which of the four was the minimum
' -----------------------------------------------------------------------------------------------------------------------
Private Function Min4(N1 As Long, N2 As Long, N3 As Long, N4 As Long, ByRef Which As Long) As Long

    If N1 < N2 Then
        Min4 = N1
        Which = 1
    Else
        Min4 = N2
        Which = 2
    End If

    If N3 < Min4 Then
        Min4 = N3
        Which = 3
    End If

    If N4 < Min4 Then
        Min4 = N4
        Which = 4
    End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : IsUnicodeFile
' Purpose    : Tests if a file is unicode by reading the byte-order-mark. Return is True, False or an error is raised.
'              Adapted from
'              https://stackoverflow.com/questions/36188224/vba-test-encoding-of-a-text-file
' -----------------------------------------------------------------------------------------------------------------------
Private Function IsUnicodeFile(FilePath As String)

    Dim intAsc1Chr As Long
    Dim intAsc2Chr As Long
    Dim T As Scripting.TextStream
    Static FSO As Scripting.FileSystemObject

    On Error GoTo ErrHandler
    If FSO Is Nothing Then Set FSO = New Scripting.FileSystemObject
    If (FSO.FileExists(FilePath) = False) Then
        IsUnicodeFile = "#File not found!"
        Exit Function
    End If

    ' 1=Read-only, False=do not create if not exist, -1=Unicode 0=ASCII
    Set T = FSO.OpenTextFile(FilePath, 1, False, 0)
    If T.AtEndOfStream Then
        T.Close: Set T = Nothing
        IsUnicodeFile = False
        Exit Function
    End If
    intAsc1Chr = Asc(T.Read(1))
    If T.AtEndOfStream Then
        T.Close: Set T = Nothing
        IsUnicodeFile = False
        Exit Function
    End If
    intAsc2Chr = Asc(T.Read(1))
    T.Close
    If (intAsc1Chr = 255) And (intAsc2Chr = 254) Then
        IsUnicodeFile = True
    Else
        IsUnicodeFile = False
    End If

    Exit Function
ErrHandler:
    Throw "#IsUnicodeFile: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : InferDelimiter
' Purpose    : Infer the delimiter in a file by looking for first occurrence outside quoted regions of comma, tab,
'              semi-colon, colon or pipe (|).
' -----------------------------------------------------------------------------------------------------------------------
Private Function InferDelimiter(FileName As String, Unicode As Boolean)
    
    Const CHUNK_SIZE = 1000
    Const QuoteChar As String = """"
    Dim Contents As String
    Dim CopyOfErr As String
    Dim EvenQuotes As Boolean
    Dim F As Scripting.File
    Dim FoundInEven As Boolean
    Dim FSO As Scripting.FileSystemObject
    Dim i As Long
    Dim T As TextStream

    On Error GoTo ErrHandler

    Set FSO = New FileSystemObject
    Set F = FSO.GetFile(FileName)
    Set T = F.OpenAsTextStream(ForReading, IIf(Unicode, TristateTrue, TristateFalse))

    If T.AtEndOfStream Then
        T.Close: Set T = Nothing: Set F = Nothing
        Throw "File is empty"
    End If

    EvenQuotes = True
    While Not T.AtEndOfStream
        Contents = T.Read(CHUNK_SIZE)
        For i = 1 To Len(Contents)
            Select Case Mid$(Contents, i, 1)
                Case QuoteChar
                    EvenQuotes = Not EvenQuotes
                Case ",", vbTab, "|", ";", ":"
                    If EvenQuotes Then
                        InferDelimiter = Mid$(Contents, i, 1)
                        T.Close: Set T = Nothing: Set F = Nothing
                        Exit Function
                    Else
                        FoundInEven = True
                    End If
            End Select
        Next i
    Wend

    'No commonly-used delimiter found in the file outside quoted regions. There are two possibilities: _
    either the file has only one column or some other character(s) has been used, returning comma is _
        equivalent to assuming the former.

    InferDelimiter = ","

    Exit Function
ErrHandler:
    CopyOfErr = "#InferDelimiter: " & Err.Description & "!"
    If Not T Is Nothing Then
        T.Close
        Set T = Nothing: Set F = Nothing: Set FSO = Nothing
    End If
    Throw CopyOfErr
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseDateFormat
' Purpose    : Populate DateOrder and DateSeparator by parsing DateFormat.
' Parameters :
'  DateFormat   : String such as D/M/Y or Y-M-D
'  DateOrder    : ByRef argument is set to DateFormat using same convention as Application.International(xlDateOrder)
'                 (0 = MDY, 1 = DMY, 2 = YMD)
'  DateSeparator: ByRef argument is set to the DateSeparator, typically "-" or "/"
' -----------------------------------------------------------------------------------------------------------------------
Private Sub ParseDateFormat(ByVal DateFormat As String, ByRef DateOrder As Long, ByRef DateSeparator As String)

    Const Err_DateFormat = "DateFormat should be 'M-D-Y', 'D-M-Y' or 'Y-M-D'. A character other " + _
        "than '-' is allowed as the separator. Omit to use the Windows default, which on this PC is "

    On Error GoTo ErrHandler

    'Replace repeated D's with a single D, etc since sParseDateCore only needs _
     to know the order in which the three parts of the date appear.
    If Len(DateFormat) > 5 Then
        DateFormat = UCase(DateFormat)
        ReplaceRepeats DateFormat, "D"
        ReplaceRepeats DateFormat, "M"
        ReplaceRepeats DateFormat, "Y"
    End If

    If Len(DateFormat) = 0 Then
        DateOrder = Application.International(xlDateOrder)
        DateSeparator = Application.International(xlDateSeparator)
    ElseIf Len(DateFormat) <> 5 Then
        Throw Err_DateFormat + WindowsDefaultDateFormat
    ElseIf Mid$(DateFormat, 2, 1) <> Mid$(DateFormat, 4, 1) Then
        Throw Err_DateFormat + WindowsDefaultDateFormat
    Else
        DateSeparator = Mid$(DateFormat, 2, 1)
        Select Case UCase$(Left$(DateFormat, 1) + Mid$(DateFormat, 3, 1) + Right$(DateFormat, 1))
            Case "MDY"
                DateOrder = 0
            Case "DMY"
                DateOrder = 1
            Case "YMD"
                DateOrder = 2
            Case Else
                Throw Err_DateFormat + WindowsDefaultDateFormat
        End Select
    End If

    Exit Sub
ErrHandler:
    Throw "#ParseDateFormat: " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ReplaceRepeats
' Purpose    : Replace repeated instances of a character in a string with a single instance.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub ReplaceRepeats(ByRef TheString As String, TheChar As String)
    Dim ChCh As String
    ChCh = TheChar & TheChar
    While InStr(TheString, ChCh) > 0
        TheString = Replace(TheString, ChCh, TheChar)
    Wend
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : WindowsDefaultDateFormat
' Purpose    : Returns a description of the system date format, used only for error string generation.
' -----------------------------------------------------------------------------------------------------------------------
Private Function WindowsDefaultDateFormat() As String
    Dim DS As String
    On Error GoTo ErrHandler
    DS = Application.International(xlDateSeparator)
    Select Case Application.International(xlDateOrder)
        Case 0
            WindowsDefaultDateFormat = "M" + DS + "D" + DS + "Y"
        Case 1
            WindowsDefaultDateFormat = "D" + DS + "M" + DS + "Y"
        Case 2
            WindowsDefaultDateFormat = "Y" + DS + "M" + DS + "D"
    End Select

    Exit Function
ErrHandler:
    WindowsDefaultDateFormat = "Cannot determine!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ShowTextFile
' Purpose    : Parse any text file to a 1-column two-dimensional array of strings. No splitting into columns and no
'              casting.
' -----------------------------------------------------------------------------------------------------------------------
Private Function ShowTextFile(FileName, StartRow As Long, NumRows As Long, _
    FileIsUnicode As Boolean)

    Dim Contents1D() As String
    Dim Contents2D() As String
    Dim F As Scripting.File
    Dim FSO As Scripting.FileSystemObject
    Dim i As Long
    Dim ReadAll As String
    Dim T As Scripting.TextStream

    On Error GoTo ErrHandler
    Set FSO = New FileSystemObject
    Set F = FSO.GetFile(FileName)

    Set T = F.OpenAsTextStream(ForReading, IIf(FileIsUnicode, TristateTrue, TristateFalse))
    For i = 1 To StartRow - 1
        T.SkipLine
    Next

    If NumRows = 0 Then
        ReadAll = T.ReadAll
        T.Close: Set T = Nothing: Set F = Nothing: Set FSO = Nothing

        ReadAll = Replace(ReadAll, vbCrLf, vbLf)
        ReadAll = Replace(ReadAll, vbCr, vbLf)

        'Text files may or may not be terminated by EOL characters...
        If Right$(ReadAll, 1) = vbLf Then
            ReadAll = Left$(ReadAll, Len(ReadAll) - 1)
        End If

        If Len(ReadAll) = 0 Then
            ReDim Contents1D(0 To 0)
        Else
            Contents1D = VBA.Split(ReadAll, vbLf)
        End If
        ReDim Contents2D(1 To UBound(Contents1D) - LBound(Contents1D) + 1, 1 To 1)
        For i = LBound(Contents1D) To UBound(Contents1D)
            Contents2D(i + 1, 1) = Contents1D(i)
        Next i
    Else
        ReDim Contents2D(1 To NumRows, 1 To 1)

        For i = 1 To NumRows 'BUG, won't work for Mac files. TODO Fix this?
            Contents2D(i, 1) = T.ReadLine
        Next i

        T.Close: Set T = Nothing: Set F = Nothing: Set FSO = Nothing
    End If

    ShowTextFile = Contents2D

    Exit Function
ErrHandler:
    Throw "#ShowTextFile: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseCSVContents
' Purpose    : Parse the contents of a CSV file. Returns a string Buffer together with arrays which assist splitting Buffer
'              into a two-dimensional array.
' Parameters :
'  ContentsOrStream: The contents of a CSV file as a string, or else a Scripting.TextStream.
'  QuoteChar       : The quote character, usually ascii 34 ("), which allow fields to contain characters that would
'                    otherwise be significant to parsing, such as delimiters or new line characters.
'  Delimiter       : The string that separates fields within each line. Typically a single character, but needn't be.
'  SkipToRow       : Rows in the file prior to SkipToRow are ignored.
'  NumRows         : The number of rows to parse. 0 for all rows from SkipToRow to the end of the file.
'  NumRowsFound    : Set to the number of rows in the file that are on or after SkipToRow.
'  NumColsFound    : Set to the number of columns in the file, i.e. the maximum number of fields in any single line.
'  NumFields       : Set to the number of fields in the file that are on or after SkipToRow.  May be less than
'                    NumRowsFound times NumColsFound if not all lines have the same number of fields.
'  Starts          : Set to an array of size at least NumFields. Element k gives the point in Buffer at which the
'                    kth field starts.
'  Lengths         : Set to an array of size at least NumFields. Element k gives the length of the kth field.
'  RowIndexes      : Set to an array of size at least NumFields. Element k gives the row at which the kth field should
'                    appear in the return from CSVRead.
'  ColIndexes      : Set to an array of size at least NumFields. Element k gives the column at which the kth field would
'                    appear in the return from CSVRead under the assumption that argument SkipToCol is 1.
'  QuoteCounts     : Set to an array of size at least NumFields. Element k gives the number of QuoteChars that appear in the
'                    kth field.
' -----------------------------------------------------------------------------------------------------------------------
Private Function ParseCSVContents(ContentsOrStream As Variant, QuoteChar As String, Delimiter As String, SkipToRow As Long, _
    NumRows As Long, ByRef NumRowsFound As Long, ByRef NumColsFound As Long, ByRef NumFields As Long, ByRef Starts() As Long, _
    ByRef Lengths() As Long, RowIndexes() As Long, ColIndexes() As Long, QuoteCounts() As Long) As String

    Const Err_ContentsOrStream = "ContentsOrStream must either be a string or a TextStream"
    Dim Buffer As String
    Dim BufferUpdatedTo As Long
    Dim ColNum As Long
    Dim EvenQuotes As Boolean
    Dim HaveReachedSkipToRow As Boolean
    Dim i As Long 'Index to read from Buffer
    Dim j As Long 'Index to write to Starts, Lengths, RowIndexes and ColIndexes
    Dim LDlm As Long
    Dim NumRowsInFile As Long
    Dim OrigLen As Long
    Dim PosCR As Long
    Dim PosDL As Long
    Dim PosLF As Long
    Dim PosQC As Long
    Dim QuoteArray() As String
    Dim QuoteCount As Long
    Dim RowNum As Long
    Dim SearchFor() As String
    Dim Streaming As Boolean
    Dim T As Scripting.TextStream
    Dim tmp As Long
    Dim Which As Long

    On Error GoTo ErrHandler
    
    If VarType(ContentsOrStream) = vbString Then
        Buffer = ContentsOrStream
        Streaming = False
    ElseIf TypeName(ContentsOrStream) = "TextStream" Then
        Set T = ContentsOrStream
        If NumRows = 0 Then
            Buffer = T.ReadAll
            T.Close
            Streaming = False
        Else
            Call GetMoreFromStream(T, Delimiter, QuoteChar, Buffer, BufferUpdatedTo)
            Streaming = True
        End If
    Else
        Throw Err_ContentsOrStream
    End If
       
    If Streaming Then
        ReDim SearchFor(1 To 4)
        SearchFor(1) = Delimiter
        SearchFor(2) = vbLf
        SearchFor(3) = vbCr
        SearchFor(4) = QuoteChar
        ReDim QuoteArray(1 To 1)
        QuoteArray(1) = QuoteChar
    End If

    ReDim Starts(1 To 8)
    ReDim Lengths(1 To 8)
    ReDim RowIndexes(1 To 8)
    ReDim ColIndexes(1 To 8)
    ReDim QuoteCounts(1 To 8)
    
    LDlm = Len(Delimiter)
    OrigLen = Len(Buffer)
    If Not Streaming Then
        'Ensure Buffer terminates with vbCrLf
        If Right(Buffer, 1) <> vbCr And Right(Buffer, 1) <> vbLf Then
            Buffer = Buffer + vbCrLf
        ElseIf Right(Buffer, 1) = vbCr Then
            Buffer = Buffer + vbLf
        End If
        BufferUpdatedTo = Len(Buffer)
    End If
    
    j = 1
    ColNum = 1: RowNum = 1
    EvenQuotes = True
    Starts(1) = 1
    If SkipToRow = 1 Then HaveReachedSkipToRow = True

    Do
        If EvenQuotes Then
            If Not Streaming Then
                If PosDL <= i Then PosDL = InStr(i + 1, Buffer, Delimiter): If PosDL = 0 Then PosDL = BufferUpdatedTo + 1
                If PosLF <= i Then PosLF = InStr(i + 1, Buffer, vbLf): If PosLF = 0 Then PosLF = BufferUpdatedTo + 1
                If PosCR <= i Then PosCR = InStr(i + 1, Buffer, vbCr): If PosCR = 0 Then PosCR = BufferUpdatedTo + 1
                If PosQC <= i Then PosQC = InStr(i + 1, Buffer, QuoteChar): If PosQC = 0 Then PosQC = BufferUpdatedTo + 1
                i = Min4(PosDL, PosLF, PosCR, PosQC, Which)
            Else
                i = SearchInBuffer(SearchFor, i + 1, T, Delimiter, QuoteChar, Which, Buffer, BufferUpdatedTo)
            End If

            If i = BufferUpdatedTo + 1 Then
                Exit Do
            End If

            If j + 1 > UBound(Starts) Then
                ReDim Preserve Starts(1 To UBound(Starts) * 2)
                ReDim Preserve Lengths(1 To UBound(Lengths) * 2)
                ReDim Preserve RowIndexes(1 To UBound(RowIndexes) * 2)
                ReDim Preserve ColIndexes(1 To UBound(ColIndexes) * 2)
                ReDim Preserve QuoteCounts(1 To UBound(QuoteCounts) * 2)
            End If

            Select Case Which
                Case 1
                    'Found Delimiter
                    Lengths(j) = i - Starts(j)
                    Starts(j + 1) = i + LDlm
                    ColIndexes(j) = ColNum: RowIndexes(j) = RowNum
                    ColNum = ColNum + 1
                    QuoteCounts(j) = QuoteCount: QuoteCount = 0
                    j = j + 1
                    NumFields = NumFields + 1
                    i = i + LDlm - 1
                Case 2, 3
                    Lengths(j) = i - Starts(j)
                    If Which = 2 Then
                        'Unix line ending
                        Starts(j + 1) = i + 1
                    ElseIf Mid(Buffer, i + 1, 1) = vbLf Then
                        'Windows line ending. - it is safe to look one character ahead since Buffer terminates with vbCrLf
                        Starts(j + 1) = i + 2
                        i = i + 1
                    Else
                        'Mac line ending (Mac pre OSX)
                        Starts(j + 1) = i + 1
                    End If

                    If ColNum > NumColsFound Then NumColsFound = ColNum
                    ColIndexes(j) = ColNum: RowIndexes(j) = RowNum
                    ColNum = 1: RowNum = RowNum + 1
                    
                    QuoteCounts(j) = QuoteCount: QuoteCount = 0
                    j = j + 1
                    NumFields = NumFields + 1
                    
                    If HaveReachedSkipToRow Then
                        If RowNum = NumRows + 1 Then
                            Exit Do
                        End If
                    Else
                        If RowNum = SkipToRow Then
                            HaveReachedSkipToRow = True
                            tmp = Starts(j)
                            ReDim Starts(1 To 8): ReDim Lengths(1 To 8): ReDim RowIndexes(1 To 8)
                            ReDim ColIndexes(1 To 8): ReDim QuoteCounts(1 To 8)
                            RowNum = 1: j = 1: NumFields = 0
                            Starts(1) = tmp
                        End If
                    End If
                Case 4
                    'Found QuoteChar
                    EvenQuotes = False
                    QuoteCount = QuoteCount + 1
            End Select
        Else
            If Not Streaming Then
                PosQC = InStr(i + 1, Buffer, QuoteChar)
            Else
                If PosQC <= i Then PosQC = SearchInBuffer(QuoteArray(), i + 1, T, Delimiter, QuoteChar, 0, Buffer, BufferUpdatedTo)
            End If
            
            If PosQC = 0 Then
                'Malformed Buffer (not RFC4180 compliant). There should always be an even number of double quotes. _
                 If there are an odd number then all text after the last double quote in the file will be (part of) _
                 the last field in the last line.
                Lengths(j) = OrigLen - Starts(j) + 1
                ColIndexes(j) = ColNum: RowIndexes(j) = RowNum
                
                RowNum = RowNum + 1
                If ColNum > NumColsFound Then NumColsFound = ColNum
                NumFields = NumFields + 1
                Exit Do
            Else
                i = PosQC
                EvenQuotes = True
                QuoteCount = QuoteCount + 1
            End If
        End If
    Loop

    NumRowsFound = RowNum - 1
    
    If HaveReachedSkipToRow Then
        NumRowsInFile = SkipToRow - 1 + RowNum - 1
    Else
        NumRowsInFile = RowNum - 1
    End If
    
    If SkipToRow > NumRowsInFile Then
        If NumRows = 0 Then 'Attempting to read from SkipToRow to the end of the file, but that would be zero or _
                             a negative number of rows. So throw an error.
            Throw "SkipToRow (" + CStr(SkipToRow) + ") exceeds the number of rows in the file (" + CStr(NumRowsInFile) + ")"
        Else
            'Attempting to read a set number of rows, function CSVRead will return an array of Empty values.
            NumFields = 0
            NumRowsFound = 0
        End If
    End If

    ParseCSVContents = Buffer

    Exit Function
ErrHandler:
    Throw "#ParseCSVContents: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SearchInBuffer
' Purpose    : Returns the location in the buffer of the first-encountered string amongst the elements of SearchFor,
'              starting the search at point SearchFrom and finishing the search at point BufferUpdatedTo. If none found in
'              that region returns BufferUpdatedTo + 1. Otherwise returns the location of the first found and sets the
'              by-reference argument Which to indicate which element of SearchFor was the first to be found.
' -----------------------------------------------------------------------------------------------------------------------
Private Function SearchInBuffer(SearchFor() As String, StartingAt As Long, T As Scripting.TextStream, Delimiter As String, _
    QuoteChar As String, ByRef Which As Long, ByRef Buffer As String, ByRef BufferUpdatedTo As Long)

    Dim InstrRes As Long
    Dim PrevBufferUpdatedTo As Long

    On Error GoTo ErrHandler

    InstrRes = InStrMulti(SearchFor, Buffer, StartingAt, BufferUpdatedTo, Which)
    If (InstrRes > 0 And InstrRes <= BufferUpdatedTo) Then
        SearchInBuffer = InstrRes
        Exit Function
    ElseIf T.AtEndOfStream Then
        SearchInBuffer = BufferUpdatedTo + 1
        Exit Function
    End If

    Do
        PrevBufferUpdatedTo = BufferUpdatedTo
        GetMoreFromStream T, Delimiter, QuoteChar, Buffer, BufferUpdatedTo
        InstrRes = InStrMulti(SearchFor, Buffer, PrevBufferUpdatedTo + 1, BufferUpdatedTo, Which)
        If (InstrRes > 0 And InstrRes <= BufferUpdatedTo) Then
            SearchInBuffer = InstrRes
            Exit Function
        ElseIf T.AtEndOfStream Then
            SearchInBuffer = BufferUpdatedTo + 1
            Exit Function
        End If
    Loop
    Exit Function
ErrHandler:
    Throw "#SearchInBuffer: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : InStrMulti
' Purpose    : Returns the first point in SearchWithin at which one of the elements of SearchFor is found, search is
'              restricted to region [StartingAt, EndingAt] and Which is updated with the index into SearchFor of the
'              first string found.
' -----------------------------------------------------------------------------------------------------------------------
Private Function InStrMulti(SearchFor() As String, SearchWithin As String, StartingAt As Long, EndingAt As Long, _
    ByRef Which As Long)

    Const Inf = 2147483647
    Dim i As Long
    Dim InstrRes() As Long
    Dim LB As Long, UB As Long
    Dim Result As Long

    On Error GoTo ErrHandler
    LB = LBound(SearchFor): UB = UBound(SearchFor)

    Result = Inf

    ReDim InstrRes(LB To UB)
    For i = LB To UB
        InstrRes(i) = InStr(StartingAt, SearchWithin, SearchFor(i))
        If InstrRes(i) > 0 Then
            If InstrRes(i) <= EndingAt Then
                If InstrRes(i) < Result Then
                    Result = InstrRes(i)
                    Which = i
                End If
            End If
        End If
    Next
    InStrMulti = IIf(Result = Inf, 0, Result)

    Exit Function
ErrHandler:
    Throw "#InStrMulti: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : GetMoreFromStream
' Purpose    : Write CHUNKSIZE characters from the TextStream T into the buffer, modifying the passed-by-reference arguments
'              Buffer, BufferUpdatedTo and Streaming.
'              Complexities:
'           a) We have to be careful not to update the buffer to a point part-way through a two-character end-of-line or a
'              multi-character delimiter, otherwise calling method SearchInBuffer might give the wrong result.
'           b) We update a few characters of the buffer beyond the BufferUpdatedTo point with the delimiter, the QuoteChar
'              and vbCrLf. This ensures that the calls to Instr that search the buffer for these strings do not needlessly
'              scan the unupdated part of the buffer.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub GetMoreFromStream(T As Scripting.TextStream, Delimiter As String, QuoteChar As String, ByRef Buffer As String, _
    ByRef BufferUpdatedTo As Long)
    Const CHUNKSIZE = 5000  ' The number of characters to read from the stream on each call. _
                              Set to a small number for testing logic and a bigger number for _
                              performance, but not too high since a common use case is reading _
                              just the first line of a file. Suggest 5000? Note that when reading _
                              an entire file (NumRows argument to CSVRead is zero) function _
                              GetMoreFromStream is not called, instead the entire file is read _
                              with a single call to T.ReadAll.
    Dim ExpandBufferBy As Long
    Dim FirstPass As Boolean
    Dim i As Long
    Dim NCharsToWriteToBuffer As Long
    Dim NewChars
    Dim OKToExit

    On Error GoTo ErrHandler
    FirstPass = True
    Do
        NewChars = T.Read(IIf(FirstPass, CHUNKSIZE, 1))
        FirstPass = False
        If T.AtEndOfStream Then
            'Ensure NewChars terminates with vbCrLf
            If Right(NewChars, 1) <> vbCr And Right(NewChars, 1) <> vbLf Then
                NewChars = NewChars + vbCrLf
            ElseIf Right(NewChars, 1) = vbCr Then
                NewChars = NewChars + vbLf
            End If
        End If

        NCharsToWriteToBuffer = Len(NewChars) + Len(Delimiter) + 3

        If BufferUpdatedTo + NCharsToWriteToBuffer > Len(Buffer) Then
            ExpandBufferBy = MaxLngs(Len(Buffer), NCharsToWriteToBuffer)
            Buffer = Buffer & String(ExpandBufferBy, "?")
        End If
        
        Mid(Buffer, BufferUpdatedTo + 1, Len(NewChars)) = NewChars
        BufferUpdatedTo = BufferUpdatedTo + Len(NewChars)

        OKToExit = True
        'Ensure we don't leave the buffer updated to part way through a two-character end of line marker.
        If Right(NewChars, 1) = vbCr Then
            OKToExit = False
        End If
        'Ensure we don't leave the buffer updated to a point part-way through a multi-character delimiter
        If Len(Delimiter) > 1 Then
            For i = 1 To Len(Delimiter) - 1
                If Mid$(Buffer, BufferUpdatedTo - i + 1, i) = Left(Delimiter, i) Then
                    OKToExit = False
                    Exit For
                End If
            Next i
            If Mid(Buffer, BufferUpdatedTo - Len(Delimiter) + 1, Len(Delimiter)) = Delimiter Then
                OKToExit = True
            End If
        End If
        If OKToExit Then Exit Do
    Loop

    'Line below arranges that when calling Instr(Buffer,....) we don't pointlessly scan the space characters _
     we can be sure that there is space in the buffer to write the extra characters thanks to
    Mid(Buffer, BufferUpdatedTo + 1, Len(Delimiter) + 3) = vbCrLf & QuoteChar & Delimiter

    Exit Sub
ErrHandler:
    Throw "#GetMoreFromStream: " & Err.Description & "!"
End Sub

Private Function MaxLngs(x As Long, y As Long) As Long
    If x > y Then
        MaxLngs = x
    Else
        MaxLngs = y
    End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CastToVariant
' Purpose    : Convert a string to the value that it represents, or return the string unchanged if conversion not
'              possible. Note that this function is passed only those fields in the file that are not quoted.
' Parameters :
'  strIn                    : The input string.
'Numbers
'  ShowNumbersAsNumbers     : If inStr is a string representation of a number should the function return that number?
'  SepsStandard             : Is the decimal separator the same as the system defaults? If True then the next two
'                             arguments are ignored.
'  DecimalSeparator         : The decimal separator used in the input string.
'  SysDecimalSeparator      : The default decimal separator on the system.
'Dates
'  ShowDatesAsDates         : If inStr represents a date should the function return taht date?
'  DateOrder                : If inStr is a string representaiton of a date it will use this date order.
'                             0 = M-D-Y, 1= D-M-Y, 2 = Y-M-D.
'  DateSeparator            : The date separator used by inStr, typically "-" or "/".
'  SysDateOrder             : The Windows system date order. 0 = M-D-Y, 1= D-M-Y, 2 = Y-M-D.
'  SysDateSeparator         : The Windows system date separator.
'Booleans
'  ShowBooleansAsBooleans   : If inStr matches "TRUE" and "FALSE" (case insensitive) should the Boolean value be returned?
'Errors
'  ShowErrorsAsErrors       : Should strings that match how errors are represented in Excel worksheets be converted to
'                             those errors values?
' -----------------------------------------------------------------------------------------------------------------------
Private Function CastToVariant(strIn As String, ShowNumbersAsNumbers As Boolean, SepsStandard As Boolean, _
    DecimalSeparator As String, SysDecimalSeparator As String, _
    ShowDatesAsDates As Boolean, DateOrder As Long, DateSeparator As String, SysDateOrder As Long, _
    SysDateSeparator As String, ShowBooleansAsBooleans As Boolean, _
    ShowErrorsAsErrors As Boolean)

    Dim bResult As Boolean
    Dim Converted As Boolean
    Dim dblResult As Double
    Dim dtResult As Date
    Dim eResult As Variant

    If ShowNumbersAsNumbers Then
        CastToDouble strIn, dblResult, SepsStandard, DecimalSeparator, SysDecimalSeparator, Converted
        If Converted Then
            CastToVariant = dblResult
            Exit Function
        End If
    End If

    If ShowDatesAsDates Then
        CastToDate strIn, dtResult, DateOrder, DateSeparator, SysDateOrder, SysDateSeparator, Converted
        If Converted Then
            CastToVariant = dtResult
            Exit Function
        End If
    End If

    If ShowBooleansAsBooleans Then
        CastToBool strIn, bResult, Converted
        If Converted Then
            CastToVariant = bResult
            Exit Function
        End If
    End If

    If ShowErrorsAsErrors Then
        CastToError strIn, eResult, Converted
        If Converted Then
            CastToVariant = eResult
            Exit Function
        End If
    End If

    CastToVariant = strIn
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CastToDouble
' Purpose    : Casts strIn to double where strIn has specified decimals separator.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub CastToDouble(strIn As String, ByRef dblOut As Double, SepsStandard As Boolean, DecimalSeparator As String, _
    SysDecimalSeparator As String, ByRef Converted As Boolean)
    
    On Error GoTo ErrHandler
    If SepsStandard Then
        dblOut = CDbl(strIn)
    Else
        dblOut = CDbl(Replace(strIn, DecimalSeparator, SysDecimalSeparator))
    End If
    Converted = True
ErrHandler:
    'Do nothing - strIn was not a string representing a number.
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CastToDate
' Purpose    : In-place conversion of a string that looks like a date into a Long or Date. No error if string cannot be
'              converted to date.
' Parameters :
'  strIn           : String
'  dtOut           : Result of cast
'  DateOrder       : The date order respected by the contents of strIn. 0 = M-D-Y, 1= D-M-Y, 2 = Y-M-D
'  DateSeparator   : The date separator used by the input
'  SysDateOrder    : The Windows system date order. 0 = M-D-Y, 1= D-M-Y, 2 = Y-M-D
'  SysDateSeparator: The Windows system date separator
'  Converted       : Boolean flipped to TRUE if conversion takes place
' -----------------------------------------------------------------------------------------------------------------------
Private Sub CastToDate(strIn As String, ByRef dtOut As Date, DateOrder As Long, DateSeparator As String, _
    SysDateOrder As Long, SysDateSeparator As String, ByRef Converted As Boolean)
    
    Dim d As String
    Dim m As String
    Dim pos1 As Long
    Dim pos2 As Long
    Dim y As String
    
    On Error GoTo ErrHandler
    pos1 = InStr(strIn, DateSeparator)
    If pos1 = 0 Then Exit Sub
    pos2 = InStr(pos1 + 1, strIn, DateSeparator)
    If pos2 = 0 Then Exit Sub

    If DateOrder = 0 Then
        m = Left$(strIn, pos1 - 1)
        d = Mid$(strIn, pos1 + 1, pos2 - pos1 - 1)
        y = Mid$(strIn, pos2 + 1)
    ElseIf DateOrder = 1 Then
        d = Left$(strIn, pos1 - 1)
        m = Mid$(strIn, pos1 + 1, pos2 - pos1 - 1)
        y = Mid$(strIn, pos2 + 1)
    ElseIf DateOrder = 2 Then
        y = Left$(strIn, pos1 - 1)
        m = Mid$(strIn, pos1 + 1, pos2 - pos1 - 1)
        d = Mid$(strIn, pos2 + 1)
    Else
        Throw "DateOrder must be 0, 1, or 2"
    End If
    If SysDateOrder = 0 Then
        dtOut = CDate(m + SysDateSeparator + d + SysDateSeparator + y)
        Converted = True
    ElseIf SysDateOrder = 1 Then
        dtOut = CDate(d + SysDateSeparator + m + SysDateSeparator + y)
        Converted = True
    ElseIf SysDateOrder = 2 Then
        dtOut = CDate(y + SysDateSeparator + m + SysDateSeparator + d)
        Converted = True
    End If

    Exit Sub
ErrHandler:
    'Do nothing - was not a string representing a date with the specified date order and date separator.
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CastToBool
' Purpose    : Convert string to Boolean, case insensitive.
' -----------------------------------------------------------------------------------------------------------------------
Private Function CastToBool(strIn As String, ByRef bOut As Boolean, ByRef Converted)
    Dim l As Long
    If VarType(strIn) = vbString Then
        l = Len(strIn)
        If l = 4 Then
            If StrComp(strIn, "true", vbTextCompare) = 0 Then
                bOut = True
                Converted = True
            End If
        ElseIf l = 5 Then
            If StrComp(strIn, "false", vbTextCompare) = 0 Then
                bOut = False
                Converted = True
            End If
        End If
    End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CastToError
' Purpose    : Convert the string representation of Excel errors back to Excel errors.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub CastToError(strIn As String, ByRef eOut As Variant, ByRef Converted As Boolean)
    On Error GoTo ErrHandler
    If Left(strIn, 1) = "#" Then
        Converted = True
        Select Case strIn 'Editing this function? Then its inverse function Encode!!!!
            Case "#DIV/0!"
                eOut = CVErr(xlErrDiv0)
            Case "#NAME?"
                eOut = CVErr(xlErrName)
            Case "#REF!"
                eOut = CVErr(xlErrRef)
            Case "#NUM!"
                eOut = CVErr(xlErrNum)
            Case "#NULL!"
                eOut = CVErr(xlErrNull)
            Case "#N/A"
                eOut = CVErr(xlErrNA)
            Case "#VALUE!"
                eOut = CVErr(xlErrValue)
            Case "#SPILL!"
                eOut = CVErr(2045)    'CVErr(xlErrNoSpill)'These constants introduced in Excel 2016
            Case "#BLOCKED!"
                eOut = CVErr(2047)    'CVErr(xlErrBlocked)
            Case "#CONNECT!"
                eOut = CVErr(2046)    'CVErr(xlErrConnect)
            Case "#UNKNOWN!"
                eOut = CVErr(2048)    'CVErr(xlErrUnknown)
            Case "#GETTING_DATA!"
                eOut = CVErr(2043)    'CVErr(xlErrGettingData)
            Case "#FIELD!"
                eOut = CVErr(2049)    'CVErr(xlErrField)
            Case "#CALC!"
                eOut = CVErr(2050)    'CVErr(xlErrField)
            Case Else
                Converted = False
        End Select
    End If

    Exit Sub
ErrHandler:
    Throw "#CastToError: " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : OStoEOL
' Purpose    : Convert text describing an operating system to the end-of-line marker employed. Note that "Mac" converts
'              to vbCr but Apple operating systems since OSX use vbLf, matching Unix.
' -----------------------------------------------------------------------------------------------------------------------
Private Function OStoEOL(OS As String, ArgName As String) As String

    Const Err_Invalid = " must be one of ""Windows"", ""Unix"" or ""Mac"", or the associented end of line characters."

    On Error GoTo ErrHandler
    Select Case LCase(OS)
        Case "windows", vbCrLf
            OStoEOL = vbCrLf
        Case "unix", "linux", vbLf
            OStoEOL = vbLf
        Case "mac", vbCr
            OStoEOL = vbCr
        Case Else
            Throw ArgName + Err_Invalid
    End Select

    Exit Function
ErrHandler:
    Throw "#OStoEOL: " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------------------------
' Procedure : CSVWrite
' Purpose   : Creates a comma-separated file on disk containing Data. Any existing file of the same name
'             is overwritten. If successful, the function returns FileName, otherwise an
'             "error string" (starts with #, ends with !) describing what went wrong.
' Arguments
' FileName  : The full name of the file, including the path.
' Data      : An array of data. Elements may be strings, numbers, dates, Booleans, empty, Excel errors
'             or null values.
' QuoteAllStrings: If TRUE (the default) then all strings in Data are quoted before being written to file. If
'             FALSE only strings containing Delimiter, line feed, carriage return or double
'             quote are quoted. Double quotes are always escaped by another double quote.
' DateFormat: A format string that determine how dates, including cells formatted as dates, appear in
'             the file. If omitted, defaults to "yyyy-mm-dd".
' DateTimeFormat: A format string that determines how dates with non-zero time part appear in the file. If
'             omitted defaults to "yyyy-mm-dd hh:mm:ss".The companion function CVSRead is
'             not currently capable of interpreting fields written in DateTime format.
' Delimiter : The delimiter string, if omitted defaults to a comma. Delimiter may have more than one
'             character.
' Unicode   : If FALSE (the default) the file written will be ascii. If TRUE the file written will be
'             unicode. An error will result if Unicode is FALSE but Data contains strings
'             with unicode characters.
' EOL       : Enter the required line ending character as "Windows" (or ascii 13 plus ascii 10), or
'             "Unix" (or ascii 10) or "Mac" (or ascii 13). If omitted defaults to
'             "Windows".
'
' Notes     : See also companion function CSVRead.
'
'             For definition of the CSV format see
'             https://tools.ietf.org/html/rfc4180#section-2
'---------------------------------------------------------------------------------------------------------
Public Function CSVWrite(FileName As String, ByVal Data As Variant, Optional QuoteAllStrings As Boolean = True, _
    Optional DateFormat As String = "yyyy-mm-dd", Optional DateTimeFormat As String = "yyyy-mm-dd hh:mm:ss", _
    Optional Delimiter As String = ",", Optional Unicode As Boolean, Optional ByVal EOL As String = vbCrLf)
Attribute CSVWrite.VB_Description = "Creates a comma-separated file on disk containing Data. Any existing file of the same name is overwritten. If successful, the function returns FileName, otherwise an ""error string"" (starts with #, ends with !) describing what went wrong."
Attribute CSVWrite.VB_ProcData.VB_Invoke_Func = " \n14"

    Const DQ = """"
    Dim EOLIsWindows As Boolean
    Dim FSO As Scripting.FileSystemObject
    Dim i As Long
    Dim j As Long
    Dim OneLine() As String
    Dim OneLineJoined As String
    Dim T As Scripting.TextStream
    
    Const Err_Delimiter = "Delimiter must not contain double quote or line feed characters"
    Const Err_Dimensions = "Data must be a range or a 2-dimensional array"

    On Error GoTo ErrHandler

    EOL = OStoEOL(EOL, "EOL")
    EOLIsWindows = EOL = vbCrLf

    If InStr(Delimiter, DQ) > 0 Or InStr(Delimiter, vbLf) > 0 Or InStr(Delimiter, vbCr) > 0 Then
        Throw Err_Delimiter
    End If

    If TypeName(Data) = "Range" Then
        'Preserve elements of type Date by using .Value, not .Value2
        Data = Data.value
    End If
    
    If NumDimensions(Data) <> 2 Then Throw Err_Dimensions
    
    Set FSO = New FileSystemObject
    Set T = FSO.CreateTextFile(FileName, True, Unicode)

    ReDim OneLine(LBound(Data, 2) To UBound(Data, 2))

    For i = LBound(Data) To UBound(Data)
        For j = LBound(Data, 2) To UBound(Data, 2)
            OneLine(j) = Encode(Data(i, j), QuoteAllStrings, DateFormat, DateTimeFormat)
        Next j
        OneLineJoined = VBA.Join(OneLine, Delimiter)
        WriteLineWrap T, OneLineJoined, EOLIsWindows, EOL, Unicode
    Next i

    T.Close: Set T = Nothing: Set FSO = Nothing
    CSVWrite = FileName
    Exit Function
ErrHandler:
    CSVWrite = "#CSVWrite: " & Err.Description & "!"
    If Not T Is Nothing Then Set T = Nothing: Set FSO = Nothing

End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : WriteLineWrap
' Purpose    : Wrapper to TextStream.Write[Line] to give more informative error message than "invalid procedure call or
'              argument" if the error is caused by attempting to write Unicode characters to an ascii file.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub WriteLineWrap(T As TextStream, text As String, EOLIsWindows As Boolean, EOL As String, Unicode As Boolean)

    Dim ErrDesc As String
    Dim ErrNum As Long
    Dim i As Long

    On Error GoTo ErrHandler
    If EOLIsWindows Then
        T.WriteLine text
    Else
        T.Write text
        T.Write EOL
    End If

    Exit Sub

ErrHandler:
    ErrNum = Err.Number
    ErrDesc = Err.Description
    If Not Unicode Then
        If ErrNum = 5 Then
            For i = 1 To Len(text)
                If AscW(Mid(text, i, 1)) > 255 Then
                    ErrDesc = "Data contains unicode characters (first found has code " & CStr(AscW(Mid(text, i, 1))) & _
                        ") which cannot be written to an ascii file. Try with argument Unicode as True"
                    Exit For
                End If
            Next i
        End If
    End If
    Throw "#WriteLineWrap: " & ErrDesc & "!"
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

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Throw
' Purpose    : Simple error handling.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub Throw(ByVal ErrorString As String)
    Err.Raise vbObjectError + 1, , ErrorString
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ThrowIfError
' Purpose   : In the event of an error, methods intended to be callable from spreadsheets
'             return an error string (starts with "#", ends with "!"). ThrowIfError allows such
'             methods to be used from VBA code while keeping error handling robust
'             MyVariable = ThrowIfError(MyFunctionThatReturnsAStringIfAnErrorHappens(...))
'---------------------------------------------------------------------------------------
Private Function ThrowIfError(Data As Variant)
    ThrowIfError = Data
    If VarType(Data) = vbString Then
        If Left$(Data, 1) = "#" Then
            If Right$(Data, 1) = "!" Then
                Throw CStr(Data)
            End If
        End If
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : NumDimensions
' Purpose   : Returns the number of dimensions in an array variable, or 0 if the variable
'             is not an array.
'---------------------------------------------------------------------------------------
Private Function NumDimensions(x As Variant) As Long
    Dim i As Long
    Dim y As Long
    If Not IsArray(x) Then
        NumDimensions = 0
        Exit Function
    Else
        On Error GoTo ExitPoint
        i = 1
        Do While True
            y = LBound(x, i)
            i = i + 1
        Loop
    End If
ExitPoint:
    NumDimensions = i - 1
End Function

