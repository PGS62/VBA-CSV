Attribute VB_Name = "modCodeFromStackOverflow"
Option Explicit

'Nigel Hefferman's code from https://stackoverflow.com/questions/12259595/load-csv-file-into-a-vba-array-rather-than-excel-sheet

'Edits: PGS 5 Oct 2021
'Public --> Private for subroutines
'Argument strName should now be passed including path
'RowDelimiter defaults to vbCrLf (was vbCr)
'Commented out calls to Application.statusbar
'Corrected Join2d to read Split2d in the "If Not RemoveQuotes" half of the if else clause of ArrayFromCSVfile
'

Public Function ArrayFromCSVfile( _
    strName As String, _
    Optional RowDelimiter As String = vbCrLf, _
    Optional FieldDelimiter = ",", _
    Optional RemoveQuotes As Boolean = True _
    ) As Variant

    ' Load a file created by FileToArray into a 2-dimensional array
    ' The file name is specified by strName, and it is exected to exist
    ' in the user's temporary folder. This is a deliberate restriction:
    ' it's always faster to copy remote files to a local drive than to
    ' edit them across the network

    ' RemoveQuotes=TRUE strips out the double-quote marks (Char 34) that
    ' encapsulate strings in most csv files.

    On Error Resume Next

    Dim objFSO As Scripting.FileSystemObject
    Dim arrData As Variant
    Dim strFile As String
    Dim strTemp As String

    Set objFSO = New Scripting.FileSystemObject
    '  strTemp = objFSO.GetSpecialFolder(Scripting.TemporaryFolder).ShortPath
    '  strFile = objFSO.BuildPath(strTemp, strName)
    strFile = strName
    If Not objFSO.FileExists(strFile) Then  ' raise an error?
        Exit Function
    End If

    'Application.StatusBar = "Reading the file... (" & strName & ")"

    If Not RemoveQuotes Then
        'As of 5 Oct 2021, code at https://stackoverflow.com/questions/12259595/load-csv-file-into-a-vba-array-rather-than-excel-sheet
        'incorrectly calls Join2d in the line below
        arrData = Split2d(objFSO.OpenTextFile(strFile, ForReading).ReadAll, RowDelimiter, FieldDelimiter)
        'Application.StatusBar = "Reading the file... Done"
    Else
        ' we have to do some allocation here...

        strTemp = objFSO.OpenTextFile(strFile, ForReading).ReadAll
        'Application.StatusBar = "Reading the file... Done"

        'Application.StatusBar = "Parsing the file..."

        strTemp = Replace$(strTemp, Chr(34) & RowDelimiter, RowDelimiter)
        strTemp = Replace$(strTemp, RowDelimiter & Chr(34), RowDelimiter)
        strTemp = Replace$(strTemp, Chr(34) & FieldDelimiter, FieldDelimiter)
        strTemp = Replace$(strTemp, FieldDelimiter & Chr(34), FieldDelimiter)

        If Right$(strTemp, Len(strTemp)) = Chr(34) Then
            strTemp = Left$(strTemp, Len(strTemp) - 1)
        End If

        If Left$(strTemp, 1) = Chr(34) Then
            strTemp = Right$(strTemp, Len(strTemp) - 1)
        End If

        'Application.StatusBar = "Parsing the file... Done"
        arrData = Split2d(strTemp, RowDelimiter, FieldDelimiter)
        strTemp = ""
    End If

    'Application.StatusBar = False

    Set objFSO = Nothing
    ArrayFromCSVfile = arrData
    Erase arrData
End Function


Private Function Split2d(ByRef strInput As String, _
    Optional RowDelimiter As String = vbCr, _
    Optional FieldDelimiter = vbTab, _
    Optional CoerceLowerBound As Long = 0 _
    ) As Variant

    ' Split up a string into a 2-dimensional array.

    ' Works like VBA.Strings.Split, for a 2-dimensional array.
    ' Check your lower bounds on return: never assume that any array in
    ' VBA is zero-based, even if you've set Option Base 0
    ' If in doubt, coerce the lower bounds to 0 or 1 by setting
    ' CoerceLowerBound
    ' Note that the default delimiters are those inserted into the
    '  string returned by ADODB.Recordset.GetString

    On Error Resume Next

    ' Coding note: we're not doing any string-handling in VBA.Strings -
    ' allocating, deallocating and (especially!) concatenating are SLOW.
    ' We're using the VBA Join & Split functions ONLY. The VBA Join,
    ' Split, & Replace functions are linked directly to fast (by VBA
    ' standards) functions in the native Windows code. Feel free to
    ' optimise further by declaring and using the Kernel string functions
    ' if you want to.

    ' ** THIS CODE IS IN THE PUBLIC DOMAIN **
    '    Nigel Heffernan   Excellerando.Blogspot.com

    Dim i   As Long
    Dim j   As Long

    Dim i_n As Long
    Dim j_n As Long

    Dim i_lBound As Long
    Dim i_uBound As Long
    Dim j_lBound As Long
    Dim j_uBound As Long

    Dim arrTemp1 As Variant
    Dim arrTemp2 As Variant

    arrTemp1 = Split(strInput, RowDelimiter)

    i_lBound = LBound(arrTemp1)
    i_uBound = UBound(arrTemp1)

    If VBA.LenB(arrTemp1(i_uBound)) <= 0 Then
        ' clip out empty last row: a common artifact in data
         'loaded from files with a terminating row delimiter
        i_uBound = i_uBound - 1
    End If

    i = i_lBound
    arrTemp2 = Split(arrTemp1(i), FieldDelimiter)

    j_lBound = LBound(arrTemp2)
    j_uBound = UBound(arrTemp2)

    If VBA.LenB(arrTemp2(j_uBound)) <= 0 Then
     ' ! potential error: first row with an empty last field...
        j_uBound = j_uBound - 1
    End If

    i_n = CoerceLowerBound - i_lBound
    j_n = CoerceLowerBound - j_lBound

    ReDim arrData(i_lBound + i_n To i_uBound + i_n, j_lBound + j_n To j_uBound + j_n)

    ' As we've got the first row already... populate it
    ' here, and start the main loop from lbound+1

    For j = j_lBound To j_uBound
        arrData(i_lBound + i_n, j + j_n) = arrTemp2(j)
    Next j

    For i = i_lBound + 1 To i_uBound Step 1

        arrTemp2 = Split(arrTemp1(i), FieldDelimiter)

        For j = j_lBound To j_uBound Step 1
            arrData(i + i_n, j + j_n) = arrTemp2(j)
        Next j

        Erase arrTemp2

    Next i

    Erase arrTemp1

    'Application.StatusBar = False

    Split2d = arrData

End Function


Private Function Join2d(ByRef InputArray As Variant, _
    Optional RowDelimiter As String = vbCr, _
    Optional FieldDelimiter = vbTab, _
    Optional SkipBlankRows As Boolean = False _
    ) As String

    ' Join up a 2-dimensional array into a string. Works like the standard
    '  VBA.Strings.Join, for a 2-dimensional array.
    ' Note that the default delimiters are those inserted into the string
    '  returned by ADODB.Recordset.GetString

    On Error Resume Next

    ' Coding note: we're not doing any string-handling in VBA.Strings -
    ' allocating, deallocating and (especially!) concatenating are SLOW.
    ' We're using the VBA Join & Split functions ONLY. The VBA Join,
    ' Split, & Replace functions are linked directly to fast (by VBA
    ' standards) functions in the native Windows code. Feel free to
    ' optimise further by declaring and using the Kernel string functions
    ' if you want to.

    ' ** THIS CODE IS IN THE PUBLIC DOMAIN **
    '   Nigel Heffernan   Excellerando.Blogspot.com

    Dim i As Long
    Dim j As Long

    Dim i_lBound As Long
    Dim i_uBound As Long
    Dim j_lBound As Long
    Dim j_uBound As Long

    Dim arrTemp1() As String
    Dim arrTemp2() As String

    Dim strBlankRow As String

    i_lBound = LBound(InputArray, 1)
    i_uBound = UBound(InputArray, 1)

    j_lBound = LBound(InputArray, 2)
    j_uBound = UBound(InputArray, 2)

    ReDim arrTemp1(i_lBound To i_uBound)
    ReDim arrTemp2(j_lBound To j_uBound)

    For i = i_lBound To i_uBound

        For j = j_lBound To j_uBound
            arrTemp2(j) = InputArray(i, j)
        Next j

        arrTemp1(i) = Join(arrTemp2, FieldDelimiter)

    Next i

    If SkipBlankRows Then

        If Len(FieldDelimiter) = 1 Then
            strBlankRow = String(j_uBound - j_lBound, FieldDelimiter)
        Else
            For j = j_lBound To j_uBound
                strBlankRow = strBlankRow & FieldDelimiter
            Next j
        End If

        Join2d = Replace(Join(arrTemp1, RowDelimiter), strBlankRow, RowDelimiter, "")
        i = Len(strBlankRow & RowDelimiter)

        If Left(Join2d, i) = strBlankRow & RowDelimiter Then
            Mid$(Join2d, 1, i) = ""
        End If

    Else

        Join2d = Join(arrTemp1, RowDelimiter)

    End If

    Erase arrTemp1

End Function
