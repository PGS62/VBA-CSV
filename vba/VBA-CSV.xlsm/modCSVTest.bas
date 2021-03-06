Attribute VB_Name = "modCSVTest"
' VBA-CSV

' Copyright (C) 2021 - Philip Swannell (https://github.com/PGS62/VBA-CSV )
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/VBA-CSV#readme

Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RunTests
' Purpose    : Multiple calls to TestCSVRead, via functions Test1, Test2 etc. Tests against many different data files,
'              and combinations of arguments. Many of the files taken from
'              https://github.com/JuliaData/CSV.jl/tree/main/test/testfiles
' -----------------------------------------------------------------------------------------------------------------------
Sub RunTests(IncludeLargeFiles As Boolean, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef NumSkipped As Long, ByRef Failures() As String)

    Dim Folder As String

    On Error GoTo ErrHandler
    Folder = ThisWorkbook.path
    Folder = Left$(Folder, InStrRev(Folder, "\")) + "testfiles\"

    If Not FolderExists(Folder) Then Throw "Cannot find folder: '" + Folder + "'"
    Test1 Folder, NumPassed, NumFailed, Failures
    Test2 Folder, NumPassed, NumFailed, Failures
    Test3 Folder, NumPassed, NumFailed, Failures
    Test4 Folder, NumPassed, NumFailed, Failures
    Test5 Folder, NumPassed, NumFailed, Failures
    Test6 Folder, NumPassed, NumFailed, Failures
    Test7 Folder, NumPassed, NumFailed, Failures
    Test8 Folder, NumPassed, NumFailed, Failures
    Test9 Folder, NumPassed, NumFailed, Failures
    Test10 Folder, NumPassed, NumFailed, Failures
    Test11 Folder, NumPassed, NumFailed, Failures
    Test12 Folder, NumPassed, NumFailed, Failures
    Test13 Folder, NumPassed, NumFailed, Failures
    Test14 Folder, NumPassed, NumFailed, Failures
    Test15 Folder, NumPassed, NumFailed, Failures
    Test16 Folder, NumPassed, NumFailed, Failures
    Test17 Folder, NumPassed, NumFailed, Failures
    Test18 Folder, NumPassed, NumFailed, Failures
    Test19 Folder, NumPassed, NumFailed, Failures
    Test20 Folder, NumPassed, NumFailed, Failures
    Test21 Folder, NumPassed, NumFailed, Failures
    Test22 Folder, NumPassed, NumFailed, Failures
    Test23 Folder, NumPassed, NumFailed, Failures
    Test24 Folder, NumPassed, NumFailed, Failures
    Test25 Folder, NumPassed, NumFailed, Failures
    Test26 Folder, NumPassed, NumFailed, Failures
    Test27 Folder, NumPassed, NumFailed, Failures
    Test28 Folder, NumPassed, NumFailed, Failures
    Test29 Folder, NumPassed, NumFailed, Failures
    Test30 Folder, NumPassed, NumFailed, Failures
    Test31 Folder, NumPassed, NumFailed, Failures
    Test32 Folder, NumPassed, NumFailed, Failures
    Test33 Folder, NumPassed, NumFailed, Failures
    Test34 Folder, NumPassed, NumFailed, Failures
    Test35 Folder, NumPassed, NumFailed, Failures
    Test36 Folder, NumPassed, NumFailed, Failures
    Test37 Folder, NumPassed, NumFailed, Failures
    Test38 Folder, NumPassed, NumFailed, Failures
    Test39 Folder, NumPassed, NumFailed, Failures
    Test40 Folder, NumPassed, NumFailed, Failures
    Test41 Folder, NumPassed, NumFailed, Failures
    Test42 Folder, NumPassed, NumFailed, Failures
    Test43 Folder, NumPassed, NumFailed, Failures
    Test44 Folder, NumPassed, NumFailed, Failures
    Test45 Folder, NumPassed, NumFailed, Failures
    Test46 Folder, NumPassed, NumFailed, Failures
    Test47 Folder, NumPassed, NumFailed, Failures
    Test48 Folder, NumPassed, NumFailed, Failures
    Test49 Folder, NumPassed, NumFailed, Failures
    Test50 Folder, NumPassed, NumFailed, Failures
    Test51 Folder, NumPassed, NumFailed, Failures
    Test52 Folder, NumPassed, NumFailed, Failures
    Test53 Folder, NumPassed, NumFailed, Failures
    Test54 Folder, NumPassed, NumFailed, Failures
    Test55 Folder, NumPassed, NumFailed, Failures
    Test56 Folder, NumPassed, NumFailed, Failures
    Test57 Folder, NumPassed, NumFailed, Failures
    Test58 Folder, NumPassed, NumFailed, Failures
    Test59 Folder, NumPassed, NumFailed, Failures
    Test60 Folder, NumPassed, NumFailed, Failures
    Test61 Folder, NumPassed, NumFailed, Failures
    Test62 Folder, NumPassed, NumFailed, Failures
    Test63 Folder, NumPassed, NumFailed, Failures
    Test64 Folder, NumPassed, NumFailed, Failures
    Test65 Folder, NumPassed, NumFailed, Failures
    Test66 Folder, NumPassed, NumFailed, Failures
    Test67 Folder, NumPassed, NumFailed, Failures
    Test68 Folder, NumPassed, NumFailed, Failures
    Test69 Folder, NumPassed, NumFailed, Failures
    Test70 Folder, NumPassed, NumFailed, Failures
    Test71 Folder, NumPassed, NumFailed, Failures
    Test72 Folder, NumPassed, NumFailed, Failures
    Test73 Folder, NumPassed, NumFailed, Failures
    Test74 Folder, NumPassed, NumFailed, Failures
    Test75 Folder, NumPassed, NumFailed, Failures
    Test76 Folder, NumPassed, NumFailed, Failures
    Test77 Folder, NumPassed, NumFailed, Failures
    Test78 Folder, NumPassed, NumFailed, Failures
    If IncludeLargeFiles Then
        Test79 Folder, NumPassed, NumFailed, Failures
    Else
        NumSkipped = NumSkipped + 1
    End If
    Test80 Folder, NumPassed, NumFailed, Failures
    Test81 Folder, NumPassed, NumFailed, Failures
    Test82 Folder, NumPassed, NumFailed, Failures
    Test83 Folder, NumPassed, NumFailed, Failures
    Test84 Folder, NumPassed, NumFailed, Failures
    Test85 Folder, NumPassed, NumFailed, Failures
    Test86 Folder, NumPassed, NumFailed, Failures
    Test87 Folder, NumPassed, NumFailed, Failures
    Test88 Folder, NumPassed, NumFailed, Failures
    Test89 Folder, NumPassed, NumFailed, Failures
    Test90 Folder, NumPassed, NumFailed, Failures
    Test91 Folder, NumPassed, NumFailed, Failures
    If IncludeLargeFiles Then
        Test92 Folder, NumPassed, NumFailed, Failures
    Else
        NumSkipped = NumSkipped + 1
    End If
    Test93 Folder, NumPassed, NumFailed, Failures
    Test94 Folder, NumPassed, NumFailed, Failures
    Test95 Folder, NumPassed, NumFailed, Failures
    Test96 Folder, NumPassed, NumFailed, Failures
    Test97 Folder, NumPassed, NumFailed, Failures
    Test98 Folder, NumPassed, NumFailed, Failures
    Test99 Folder, NumPassed, NumFailed, Failures
    Test100 Folder, NumPassed, NumFailed, Failures
    Test101 Folder, NumPassed, NumFailed, Failures
    Test102 Folder, NumPassed, NumFailed, Failures
    Test103 Folder, NumPassed, NumFailed, Failures
    Test104 Folder, NumPassed, NumFailed, Failures
    Test105 Folder, NumPassed, NumFailed, Failures
    Test106 Folder, NumPassed, NumFailed, Failures
    Test107 Folder, NumPassed, NumFailed, Failures
    Test108 Folder, NumPassed, NumFailed, Failures
    Test109 Folder, NumPassed, NumFailed, Failures
    Test110 Folder, NumPassed, NumFailed, Failures
    Test111 Folder, NumPassed, NumFailed, Failures
    Test112 Folder, NumPassed, NumFailed, Failures
    Test113 Folder, NumPassed, NumFailed, Failures
    Test114 Folder, NumPassed, NumFailed, Failures
    Test115 Folder, NumPassed, NumFailed, Failures
    Test116 Folder, NumPassed, NumFailed, Failures
    Test117 Folder, NumPassed, NumFailed, Failures
    Test118 Folder, NumPassed, NumFailed, Failures
    Test119 Folder, NumPassed, NumFailed, Failures
    Test120 Folder, NumPassed, NumFailed, Failures
    Test121 Folder, NumPassed, NumFailed, Failures
    Test122 Folder, NumPassed, NumFailed, Failures
    Test123 Folder, NumPassed, NumFailed, Failures
    Test124 Folder, NumPassed, NumFailed, Failures
    Test125 Folder, NumPassed, NumFailed, Failures
    Test126 Folder, NumPassed, NumFailed, Failures
    Test127 Folder, NumPassed, NumFailed, Failures
    If IncludeLargeFiles Then
        Test128 Folder, NumPassed, NumFailed, Failures
    Else
        NumSkipped = NumSkipped + 1
    End If
    Test128 Folder, NumPassed, NumFailed, Failures
    Test129 Folder, NumPassed, NumFailed, Failures
    Test130 Folder, NumPassed, NumFailed, Failures
    Test131 Folder, NumPassed, NumFailed, Failures
    Test132 Folder, NumPassed, NumFailed, Failures
    Test133 Folder, NumPassed, NumFailed, Failures
    Test134 Folder, NumPassed, NumFailed, Failures
    Test135 Folder, NumPassed, NumFailed, Failures
    Test136 Folder, NumPassed, NumFailed, Failures
    Test137 Folder, NumPassed, NumFailed, Failures
    Test138 Folder, NumPassed, NumFailed, Failures
    Test139 Folder, NumPassed, NumFailed, Failures
    Test140 Folder, NumPassed, NumFailed, Failures
    Test141 Folder, NumPassed, NumFailed, Failures
    Test142 Folder, NumPassed, NumFailed, Failures
    Test143 Folder, NumPassed, NumFailed, Failures
    Test144 Folder, NumPassed, NumFailed, Failures
    Test145 Folder, NumPassed, NumFailed, Failures
    Test146 Folder, NumPassed, NumFailed, Failures
    Test147 Folder, NumPassed, NumFailed, Failures
    Test148 Folder, NumPassed, NumFailed, Failures
    Test149 Folder, NumPassed, NumFailed, Failures
    Test150 Folder, NumPassed, NumFailed, Failures
    Test151 Folder, NumPassed, NumFailed, Failures
    Test152 Folder, NumPassed, NumFailed, Failures
    Test153 Folder, NumPassed, NumFailed, Failures
    Test154 Folder, NumPassed, NumFailed, Failures
    Test155 Folder, NumPassed, NumFailed, Failures
    Test156 Folder, NumPassed, NumFailed, Failures
    Test157 Folder, NumPassed, NumFailed, Failures
    Test158 Folder, NumPassed, NumFailed, Failures
    Test159 Folder, NumPassed, NumFailed, Failures
    Test160 Folder, NumPassed, NumFailed, Failures
    Test161 Folder, NumPassed, NumFailed, Failures
    Test162 Folder, NumPassed, NumFailed, Failures
    Test163 Folder, NumPassed, NumFailed, Failures
    Test164 Folder, NumPassed, NumFailed, Failures
    Test165 Folder, NumPassed, NumFailed, Failures
    Test166 Folder, NumPassed, NumFailed, Failures
    Test167 Folder, NumPassed, NumFailed, Failures
    Test168 Folder, NumPassed, NumFailed, Failures
    Test169 Folder, NumPassed, NumFailed, Failures
    Test170 Folder, NumPassed, NumFailed, Failures
    Test171 Folder, NumPassed, NumFailed, Failures
    Test172 Folder, NumPassed, NumFailed, Failures
    Test173 Folder, NumPassed, NumFailed, Failures
    Test174 Folder, NumPassed, NumFailed, Failures
    Test175 Folder, NumPassed, NumFailed, Failures
    Test176 Folder, NumPassed, NumFailed, Failures
    Test177 Folder, NumPassed, NumFailed, Failures
    Test178 Folder, NumPassed, NumFailed, Failures
    Test179 Folder, NumPassed, NumFailed, Failures
    Test180 Folder, NumPassed, NumFailed, Failures
    Test181 Folder, NumPassed, NumFailed, Failures
    Test182 Folder, NumPassed, NumFailed, Failures
    Test183 Folder, NumPassed, NumFailed, Failures
    Test184 Folder, NumPassed, NumFailed, Failures
    Test185 Folder, NumPassed, NumFailed, Failures
    Test186 Folder, NumPassed, NumFailed, Failures
    Test187 Folder, NumPassed, NumFailed, Failures
    Test188 Folder, NumPassed, NumFailed, Failures
    Test189 Folder, NumPassed, NumFailed, Failures
    Test190 Folder, NumPassed, NumFailed, Failures
    Test191 Folder, NumPassed, NumFailed, Failures
    Test192 Folder, NumPassed, NumFailed, Failures
    Test193 Folder, NumPassed, NumFailed, Failures
    Test194 Folder, NumPassed, NumFailed, Failures
    Test195 Folder, NumPassed, NumFailed, Failures
    Test196 Folder, NumPassed, NumFailed, Failures
    Test197 Folder, NumPassed, NumFailed, Failures
    Test198 Folder, NumPassed, NumFailed, Failures
    Test199 Folder, NumPassed, NumFailed, Failures
    Test200 Folder, NumPassed, NumFailed, Failures
    Test201 Folder, NumPassed, NumFailed, Failures
    Test202 Folder, NumPassed, NumFailed, Failures
    Exit Sub
ErrHandler:
    Throw "#RunTests (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RunTestsFromButton
' Purpose    : Code behind the "Run Tests" button on the Tests worksheet
' -----------------------------------------------------------------------------------------------------------------------
Public Sub RunTestsFromButton()
    Dim Failures() As String
    Dim IncludeLargeFiles As Boolean
    Dim NumFailed As Long
    Dim NumPassed As Long
    Dim NumSkipped As Long
    Dim ProtectContents As Boolean

    On Error GoTo ErrHandler

    IncludeLargeFiles = shTest.Range("IncludeLargeFiles").value

    Failures = VBA.Split(vbNullString) 'Creates array of length zero!

    RunTests IncludeLargeFiles, NumPassed, NumFailed, NumSkipped, Failures

    With shTest
        ProtectContents = .ProtectContents
        .Unprotect
        .Range("NumPassed").value = NumPassed
        .Range("NumFailed").value = NumFailed
        .Range("NumSkipped").value = NumSkipped
        PasteFailures NumFailed, Failures
        .Protect Contents:=ProtectContents
    End With

    Exit Sub
ErrHandler:
    MsgBox "#RunTestsFromButton (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical
End Sub

Private Sub PasteFailures(NumFailures As Long, Optional Failures As Variant)
    On Error GoTo ErrHandler
    With shTestResults
        .Unprotect
        .UsedRange.EntireColumn.Delete
        With .Cells(1, 1)
            .value = "Test Results"
            .Font.Size = 22
        End With
        If NumFailures > 0 Then
            With .Cells(3, 1).Resize(NumFailures)
                .value = Transpose(Failures)
                .WrapText = False
            End With
        Else
            .Cells(3, 1).value = "All tests passed."
        End If
        .Protect , , True
    End With

    Exit Sub
ErrHandler:
    Throw "#PasteFailures (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FolderExists
' Purpose   : Returns True or False. Does not matter whether FolderPath has a terminating
'             backslash.
' -----------------------------------------------------------------------------------------------------------------------
Private Function FolderExists(ByVal FolderPath As String) As Boolean
    Dim F As Scripting.Folder
    Dim FSO As Scripting.FileSystemObject
    On Error GoTo ErrHandler
    Set FSO = New FileSystemObject
    Set F = FSO.GetFolder(FolderPath)
    FolderExists = True
    Exit Function
ErrHandler:
    FolderExists = False
End Function

Function FileExists(ByVal FilePath As String) As Boolean
    Dim F As Scripting.File
    Dim FSO As Scripting.FileSystemObject
    On Error GoTo ErrHandler
    Set FSO = New FileSystemObject
    Set F = FSO.GetFile(FilePath)
    FileExists = True
    Exit Function
ErrHandler:
    FileExists = False
End Function

Private Sub AccumulateResults(TestRes As Boolean, ByRef NumPassed As Long, ByRef NumFailed As Long, _
    WhatDiffers As String, ByRef Failures() As String)
    If TestRes Then
        NumPassed = NumPassed + 1
    Else
        NumFailed = NumFailed + 1
        ReDim Preserve Failures(LBound(Failures) To UBound(Failures) + 1)
        Failures(UBound(Failures)) = WhatDiffers
    End If
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CastDoublesToDates
' Purpose    : Cast all elements of x that are of type Double to type Date
' -----------------------------------------------------------------------------------------------------------------------
Private Sub CastDoublesToDates(ByRef x As Variant)
    Dim i As Long
    Dim j As Long
    If NumDimensions(x) = 2 Then
        For i = LBound(x, 1) To UBound(x, 1)
            For j = LBound(x, 2) To UBound(x, 2)
                If VarType(x(i, j)) = vbDouble Then
                    x(i, j) = CDate(x(i, j))
                End If
            Next
        Next
    End If
End Sub

Private Sub Test1(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test one row of data.cscv"
    Expected = HStack(1#, 2#, 3#)
    FileName = "test_one_row_of_data.cscv"
    TestRes = TestCSVRead(1, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test1 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test2(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test empty file newlines"
    FileName = "test_empty_file_newlines.csv"
    Expected = HStack(Array(Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty))
    TestRes = TestCSVRead(2, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:="N", ShowMissingsAs:=Empty, IgnoreEmptyLines:=False)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test2 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test3(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test single column"
    Expected = HStack(Array("col1", 1#, 2#, 3#))
    FileName = "test_single_column.csv"
    TestRes = TestCSVRead(3, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test3 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test4(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "comma decimal"
    FileName = "comma_decimal.csv"
    Expected = HStack(Array("x", 3.14, 1#), Array("y", 1#, 1#))
    TestRes = TestCSVRead(4, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:="N", ShowMissingsAs:=Empty, DecimalSeparator:=",")
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test4 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test5(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test missing last column"
    FileName = "test_missing_last_column.csv"
    Expected = HStack( _
        Array("A", 1#, 4#), _
        Array("B", 2#, 5#), _
        Array("C", 3#, 6#), _
        Array("D", Empty, Empty))
    TestRes = TestCSVRead(5, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:="N", ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test5 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test6(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "initial spaces when ignore repeated"
    FileName = "test_issue_326.wsv"
    Expected = HStack(Array("A", 1#, 11#), Array("B", 2#, 22#))
    TestRes = TestCSVRead(6, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:=True, Delimiter:=" ", IgnoreRepeated:=True, ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test6 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test7(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test not enough columns"
    FileName = "test_not_enough_columns.csv"
    Expected = HStack( _
        Array("A", 1#, 4#), _
        Array("B", 2#, 5#), _
        Array("C", 3#, 6#), _
        Array("D", Empty, Empty), _
        Array("E", Empty, Empty))

    TestRes = TestCSVRead(7, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:=True, ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test7 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test8(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test comments1"
    FileName = "test_comments1.csv"
    Expected = HStack(Array("a", 1#, 7#), Array("b", 2#, 8#), Array("c", 3#, 9#))
    TestRes = TestCSVRead(8, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:=True, Comment:="#", ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test8 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test9(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test comments multichar"
    FileName = "test_comments_multichar.csv"
    Expected = HStack(Array("a", 1#, 7#), Array("b", 2#, 8#), Array("c", 3#, 9#))
    TestRes = TestCSVRead(9, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:=True, Comment:="//")
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test9 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test10(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test correct trailing missings"
    FileName = "test_correct_trailing_missings.csv"
    Expected = HStack( _
        Array("A", 1#, 4#), _
        Array("B", 2#, 5#), _
        Array("C", 3#, 6#), _
        Array("D", Empty, Empty), _
        Array("E", Empty, Empty))
    TestRes = TestCSVRead(10, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:=True, ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test10 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test11(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test not enough columns2"
    FileName = "test_not_enough_columns2.csv"
    Expected = HStack( _
        Array("A", 1#, 6#), _
        Array("B", 2#, 7#), _
        Array("C", 3#, 8#), _
        Array("D", 4#, Empty), _
        Array("E", 5#, Empty))

    TestRes = TestCSVRead(11, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:=True, ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test11 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test12(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test tab null empty.txt"
    FileName = "test_tab_null_empty.txt"
    Expected = HStack( _
        Array("A", 1#, 2#), _
        Array("B", 2000#, Empty), _
        Array("C", "x", "y"), _
        Array("D", 100#, 200#))

    TestRes = TestCSVRead(12, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:=True, ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test12 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test13(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test basic"
    FileName = "test_basic.csv"
    Expected = HStack(Array("col1", 1#, 4#, 7#), Array("col2", 2#, 5#, 8#), Array("col3", 3#, 6#, 9#))
    TestRes = TestCSVRead(13, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:=True, ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test13 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test14(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test basic pipe"
    FileName = "test_basic_pipe.csv"
    Expected = HStack(Array("col1", 1#, 4#, 7#), Array("col2", 2#, 5#, 8#), Array("col3", 3#, 6#, 9#))
    TestRes = TestCSVRead(14, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:=True, ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test14 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test15(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test mac line endings"
    FileName = "test_mac_line_endings.csv"
    Expected = HStack(Array("col1", 1#, 4#, 7#), Array("col2", 2#, 5#, 8#), Array("col3", 3#, 6#, 9#))
    TestRes = TestCSVRead(15, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:=True, ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test15 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test16(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test newline line endings"
    FileName = "test_newline_line_endings.csv"
    Expected = HStack(Array("col1", 1#, 4#, 7#), Array("col2", 2#, 5#, 8#), Array("col3", 3#, 6#, 9#))
    TestRes = TestCSVRead(16, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:=True, ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test16 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test17(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test delim.tsv"
    FileName = "test_delim.tsv"
    Expected = HStack(Array("col1", 1#, 4#, 7#), Array("col2", 2#, 5#, 8#), Array("col3", 3#, 6#, 9#))
    TestRes = TestCSVRead(17, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:=True, ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test17 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test18(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test delim.wsv"
    FileName = "test_delim.wsv"
    Expected = HStack(Array("col1", 1#, 4#, 7#), Array("col2", 2#, 5#, 8#), Array("col3", 3#, 6#, 9#))
    TestRes = TestCSVRead(18, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:=True, Delimiter:=" ", ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test18 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test19(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test tab null string.txt"
    FileName = "test_tab_null_string.txt"
    Expected = HStack( _
        Array("A", 1#, 2#), _
        Array("B", 2000#, Empty), _
        Array("C", "x", "y"), _
        Array("D", 100#, 200#))

    TestRes = TestCSVRead(19, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        MissingStrings:="NULL", _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test19 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test20(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test crlf line endings"
    FileName = "test_crlf_line_endings.csv"
    Expected = HStack(Array("col1", 1#, 4#, 7#), Array("col2", 2#, 5#, 8#), Array("col3", 3#, 6#, 9#))
    TestRes = TestCSVRead(20, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:="N", ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test20 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test21(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test header on row 4"
    FileName = "test_header_on_row_4.csv"
    Expected = HStack(Array("col1", 1#, 4#, 7#), Array("col2", 2#, 5#, 8#), Array("col3", 3#, 6#, 9#))
    TestRes = TestCSVRead(21, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        SkipToRow:=4, _
        ShowMissingsAs:=Empty, _
        IgnoreEmptyLines:=False)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test21 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test22(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test missing last field"
    FileName = "test_missing_last_field.csv"
    Expected = HStack(Array("col1", 1#, 4#), Array("col2", 2#, 5#), Array("col3", 3#, Empty))
    TestRes = TestCSVRead(22, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:=True, ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test22 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test23(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test no header"
    FileName = "test_no_header.csv"
    Expected = HStack(Array(1#, 4#, 7#), Array(2#, 5#, 8#), Array(3#, 6#, 9#))
    TestRes = TestCSVRead(23, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:=True, ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test23 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test24(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test dates"
    FileName = "test_dates.csv"
    Expected = HStack(Array("col1", CDate("2015-Jan-01"), CDate("2015-Jan-02"), CDate("2015-Jan-03")))
    TestRes = TestCSVRead(24, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:=True, DateFormat:="Y-M-D", _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test24 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test25(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test excel date formats"
    FileName = "test_excel_date_formats.csv"
    Expected = HStack(Array("col1", CDate("2015-Jan-01"), CDate("2015-Feb-01"), CDate("2015-Mar-01")))
    TestRes = TestCSVRead(25, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:=True, DateFormat:="D/M/Y", ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test25 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test26(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test repeated delimiters"
    FileName = "test_repeated_delimiters.csv"
    Expected = HStack( _
        Array("a", CDate("1899-Dec-31"), CDate("1899-Dec-31"), CDate("1899-Dec-31")), _
        Array("b", 2#, 2#, 2#), _
        Array("c", 3#, 3#, 3#), _
        Array("d", 4#, 4#, 4#), _
        Array("e", 5#, 5#, 5#))

    TestRes = TestCSVRead(26, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:=True, Delimiter:=" ", IgnoreRepeated:=True, ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test26 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test27(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test simple quoted"
    FileName = "test_simple_quoted.csv"
    Expected = HStack(Array("col1", "quoted field 1"), Array("col2", "quoted field 2"))
    TestRes = TestCSVRead(27, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test27 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test28(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test footer missing"
    FileName = "test_footer_missing.csv"
    Expected = HStack( _
        Array("col1", "1", "4", "7", "10", Empty), _
        Array("col2", "2", "5", "8", "11", Empty), _
        Array("col3", "3", "6", "9", "12", Empty))

    TestRes = TestCSVRead(28, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test28 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test29(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String
    'The is test is fragile in that line endings in text files can get flipped from vbLf to vbCrLf as files are pushed and pulled to git.

    On Error GoTo ErrHandler
    TestDescription = "test quoted delim and newline"
    FileName = "test_quoted_delim_and_newline.csv"
    Expected = HStack(Array("col1", "quoted ,field 1"), Array("col2", "quoted" + vbCrLf + " field 2"))
    TestRes = TestCSVRead(29, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test29 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test30(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test missing value"
    FileName = "test_missing_value.csv"
    Expected = HStack(Array("col1", 1#, 4#, 7#), Array("col2", 2#, Empty, 8#), Array("col3", 3#, 6#, 9#))
    TestRes = TestCSVRead(30, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:=True, ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test30 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test31(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test truestrings"
    FileName = "test_truestrings.csv"
    Expected = HStack( _
        Array("int", 1#, 2#, 3#, 4#, 5#, 6#), _
        Array("bools", True, True, True, False, False, False))

    TestRes = TestCSVRead(31, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        TrueStrings:=HStack("T", "TRUE", "true"), _
        FalseStrings:=HStack("F", "FALSE", "false"), _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test31 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test32(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test floats"
    FileName = "test_floats.csv"
    Expected = HStack(Array("col1", 1#, 4#, 7#), Array("col2", 2#, 5#, 8#), Array("col3", 3#, 6#, 9#))
    TestRes = TestCSVRead(32, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:=True, ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test32 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test33(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test utf8"
    FileName = "test_utf8.csv"
    Expected = HStack(Array("col1", 1#, 4#, 7#), Array("col2", 2#, 5#, 8#), Array("col3", 3#, 6#, 9#))
    TestRes = TestCSVRead(33, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:=True, ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test33 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test34(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test windows"
    FileName = "test_windows.csv"
    Expected = HStack(Array("col1", 1#, 4#, 7#), Array("col2", 2#, 5#, 8#), Array("col3", 3#, 6#, 9#))
    TestRes = TestCSVRead(34, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:=True, ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test34 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test35(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test missing value NULL"
    FileName = "test_missing_value_NULL.csv"
    Expected = HStack(Array("col1", 1#, 4#, 7#), Array("col2", 2#, Empty, 8#), Array("col3", 3#, 6#, 9#))
    TestRes = TestCSVRead(35, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        MissingStrings:="NULL", _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test35 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test36(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    'Note we must pass "Q" option to treat quoted numbers as numbers
    TestDescription = "test quoted numbers"
    FileName = "test_quoted_numbers.csv"
    Expected = HStack( _
        Array("col1", 123#, "abc", "123abc"), _
        Array("col2", 1#, 42#, 12#), _
        Array("col3", 1#, 42#, 12#))

    TestRes = TestCSVRead(36, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:="NQ", ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test36 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test37(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    'We don't support SkipFooter
    TestDescription = "test 2 footer rows"
    FileName = "test_2_footer_rows.csv"
    Expected = HStack( _
        Array("col1", 1#, 4#, 7#, 10#, 13#), _
        Array("col2", 2#, 5#, 8#, 11#, 14#), _
        Array("col3", 3#, 6#, 9#, 12#, 15#))

    TestRes = TestCSVRead(37, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        ShowMissingsAs:=Empty, _
        IgnoreEmptyLines:=True)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test37 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test38(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test utf8 with BOM"
    FileName = "test_utf8_with_BOM.csv"
    Expected = HStack(Array("col1", 1#, 4#, 7#), Array("col2", 2#, 5#, 8#), Array("col3", 3#, 6#, 9#))
    TestRes = TestCSVRead(38, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:=True, ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test38 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test39(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    'We don't distinguish between different types of number, so this test a bit moot
    TestDescription = "types override"
    FileName = "types_override.csv"
    Expected = HStack( _
        Array("col1", "A", "B", "C"), _
        Array("col2", 1#, 5#, 9#), _
        Array("col3", 2#, 6#, 10#), _
        Array("col4", 3#, 7#, 11#), _
        Array("col5", 4#, 8#, 12#))

    TestRes = TestCSVRead(39, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:=True, ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test39 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test40(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "issue 198 part2"
    FileName = "issue_198_part2.csv"
    Expected = HStack( _
        Array("A", "a", "b", "c", "d"), _
        Array("B", -0.367, Empty, Empty, -0.364), _
        Array("C", -0.371, Empty, Empty, -0.371), _
        Array(Empty, Empty, Empty, Empty, Empty))

    TestRes = TestCSVRead(40, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        MissingStrings:="++", _
        ShowMissingsAs:=Empty, _
        DecimalSeparator:=",")
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test40 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test41(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    'Not sure how julia handles this, could not find in https://github.com/JuliaData/CSV.jl/blob/main/test/testfiles.jl
    TestDescription = "test mixed date formats"
    FileName = "test_mixed_date_formats.csv"
    Expected = HStack(Array("col1", "01/01/2015", "01/02/2015", "01/03/2015", CDate("2015-Jan-02"), CDate("2015-Jan-03")))
    TestRes = TestCSVRead(41, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:=True, DateFormat:="Y-M-D", _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test41 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test42(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test multiple missing"
    FileName = "test_multiple_missing.csv"
    Expected = HStack( _
        Array("col1", 1#, 4#, 7#, 7#), _
        Array("col2", 2#, Empty, Empty, Empty), _
        Array("col3", 3#, 6#, 9#, 9#))

    TestRes = TestCSVRead(42, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        MissingStrings:=HStack("NULL", "NA", "\N"), _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test42 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test43(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test string delimiters"
    FileName = "test_string_delimiters.csv"
    Expected = HStack( _
        Array("num1", 1#, 1#), _
        Array("num2", 1193#, 661#), _
        Array("num3", 5#, 3#), _
        Array("num4", 978300760#, 978302109#))

    TestRes = TestCSVRead(43, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:=True, Delimiter:="::", ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test43 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test44(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "bools"
    FileName = "bools.csv"
    Expected = HStack( _
        Array("col1", True, False, True, False), _
        Array("col2", False, True, True, False), _
        Array("col3", 1#, 2#, 3#, 4#))

    TestRes = TestCSVRead(44, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:=True, ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test44 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test45(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "boolext"
    FileName = "boolext.csv"
    Expected = HStack( _
        Array("col1", True, False, True, False), _
        Array("col2", False, True, True, False), _
        Array("col3", 1#, 2#, 3#, 4#))

    TestRes = TestCSVRead(45, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:=True, ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test45 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test46(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test comment first row"
    FileName = "test_comment_first_row.csv"
    Expected = HStack(Array("a", 1#, 7#), Array("b", 2#, 8#), Array("c", 3#, 9#))
    TestRes = TestCSVRead(46, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:=True, Comment:="#", ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test46 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test47(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    'NB this parses differently from how parsed by CSV.jl, we put col5, row one as number, they as string thanks to the presence of not-parsable-to-number in the cell below (the culprit is the comma in "2,773.9000")
    TestDescription = "issue 207"
    FileName = "issue_207.csv"
    Expected = HStack( _
        Array("a", 1863001#, 1863209#), _
        Array("b", 134#, 137#), _
        Array("c", 10000#, 0#), _
        Array("d", 1.0009, 1#), _
        Array("e", 1#, "2,773.9000"), _
        Array("f", -0.002033899, Empty))

    TestRes = TestCSVRead(47, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:=True, ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test47 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test48(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test comments multiple"
    FileName = "test_comments_multiple.csv"
    Expected = HStack( _
        Array("a", 1#, 7#, 10#, 13#), _
        Array("b", 2#, 8#, 11#, 14#), _
        Array("c", 3#, 9#, 12#, 15#))

    TestRes = TestCSVRead(48, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:=True, Comment:="#", ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test48 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test49(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    'NotePad++ identifies the encoding of this file as UTF-16 Little Endian. There is no BOM, so we have to explicitly pass Encoding as "UTF-16"
    TestDescription = "test utf16"
    FileName = "test_utf16.csv"
    Expected = HStack(Array("col1", 1#, 4#, 7#), Array("col2", 2#, 5#, 8#), Array("col3", 3#, 6#, 9#))
    TestRes = TestCSVRead(49, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:=True, ShowMissingsAs:=Empty, Encoding:="UTF-16")
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test49 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test50(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    'NotePad++ identifies the encoding of this file as UTF-16 Little Endian. There is no BOM, so we have to explicitly explicitly pass Encoding as "UTF-16"
    TestDescription = "test utf16 le"
    FileName = "test_utf16_le.csv"
    Expected = HStack(Array("col1", 1#, 4#, 7#), Array("col2", 2#, 5#, 8#), Array("col3", 3#, 6#, 9#))
    TestRes = TestCSVRead(50, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:=True, ShowMissingsAs:=Empty, Encoding:="UTF-16")
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test50 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test51(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test types"
    FileName = "test_types.csv"
    Expected = HStack( _
        Array("int", 1#), _
        Array("float", 1#), _
        Array("date", CDate("2018-Jan-01")), _
        Array("datetime", CDate("2018-Jan-01")), _
        Array("bool", True), _
        Array("string", "hey"), _
        Array("weakrefstring", "there"), _
        Array("missing", Empty))

    TestRes = TestCSVRead(51, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        DateFormat:="ISO", _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test51 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test52(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test 508"
    FileName = "test_508.csv"
    Expected = HStack( _
        Array("Yes", "Yes", "Yes", "Yes", "No", "Yes"), _
        Array("Medium rare", "Medium", "Medium", "Medium rare", Empty, "Rare"))

    TestRes = TestCSVRead(52, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:=True, Comment:="#", ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test52 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test53(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "issue 198"
    Expected = HStack( _
        Array(Empty, 43208#, 43207#, 43206#, 43205#, 43204#, 43203#), _
        Array("Taux de l'Eonia (moyenne mensuelle)", -0.368, -0.368, -0.367, Empty, Empty, -0.364), _
        Array("EURIBOR ? 1 mois", -0.371, -0.371, -0.371, Empty, Empty, -0.371), _
        Array("EURIBOR ? 12 mois", -0.189, -0.189, -0.189, Empty, Empty, -0.19), _
        Array("EURIBOR ? 3 mois", -0.328, -0.328, -0.329, Empty, Empty, -0.329), _
        Array("EURIBOR ? 6 mois", -0.271, -0.27, -0.27, Empty, Empty, -0.271), _
        Array("EURIBOR ? 9 mois", -0.219, -0.219, -0.219, Empty, Empty, -0.219))
    FileName = "issue_198.csv"
    TestRes = TestCSVRead(53, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        DateFormat:="D/M/Y", _
        MissingStrings:="-", _
        ShowMissingsAs:=Empty, _
        DecimalSeparator:=",", _
        Encoding:="UTF-8")
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test53 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test54(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "error comment.txt"
    FileName = "error_comment.txt"
    Expected = HStack( _
        Array("fluid", "Ar", "C2H4", "CO2", "CO", "CH4", "H2", "Kr", "Xe"), _
        Array("col2", 150.86, 282.34, 304.12, 132.85, 190.56, 32.98, 209.4, 289.74), _
        Array("col3", 48.98, 50.41, 73.74, 34.94, 45.99, 12.93, 55#, 58.4), _
        Array("acentric_factor", -0.002, 0.087, 0.225, 0.045, 0.011, -0.217, 0.005, 0.008))

    TestRes = TestCSVRead(54, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:="N", Comment:="#", ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test54 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test55(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "bug555.txt"
    FileName = "bug555.txt"
    Expected = HStack( _
        Array("RESULTAT", "A0", "B0", "C0"), _
        Array("NOM_CHAM", "A1", "B1", "C1"), _
        Array("INST", 0#, 0#, 0#), _
        Array("NUME_ORDRE", 0#, 0#, 0#), _
        Array("NOEUD", "N1", "N2", "N3"), _
        Array("COOR_X", 0#, 2.3, 2.5), _
        Array("COOR_Y", 2.27374E-15, 0#, 0#), _
        Array("COOR_Z", 0#, 0#, 0#), _
        Array("TEMP", 0.0931399, 0.311013, 0.424537))

    TestRes = TestCSVRead(55, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:=True, Delimiter:=" ", IgnoreRepeated:=True, ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test55 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test56(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "precompile small"
    FileName = "precompile_small.csv"
    Expected = HStack( _
        Array("int", 1#, 1#, 1#, 1#, 1#, 1#, 1#, 1#, 1#, Empty), _
        Array("float", 2#, 2#, 2#, 2#, 2#, 2#, 2#, 2#, 2#, Empty), _
        Array("pool", "a", "a", "a", "a", "a", "a", "a", "a", "a", Empty), _
        Array("string", "RTrBP", "aqbcM", "jN9r4", "aWGyX", "yyBbB", "sJLTp", "7N1Ky", "O8MBD", "EIidc", Empty), _
        Array("bool", True, True, True, True, True, True, True, True, True, Empty), _
        Array("date", CDate("2020-Jun-20"), CDate("2020-Jun-20"), CDate("2020-Jun-20"), CDate("2020-Jun-20"), CDate("2020-Jun-20"), CDate("2020-Jun-20"), CDate("2020-Jun-20"), CDate("2020-Jun-20"), CDate("2020-Jun-20"), Empty), _
        Array("datetime", CDate("2020-Jun-20"), CDate("2020-Jun-20"), CDate("2020-Jun-20"), CDate("2020-Jun-20"), CDate("2020-Jun-20"), CDate("2020-Jun-20"), CDate("2020-Jun-20"), CDate("2020-Jun-20"), CDate("2020-Jun-20"), Empty), _
        Array("time", CDate("12:00:00"), CDate("12:00:00"), CDate("12:00:00"), CDate("12:00:00"), CDate("12:00:00"), CDate("12:00:00"), CDate("12:00:00"), CDate("12:00:00"), CDate("12:00:00"), Empty))

    TestRes = TestCSVRead(56, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        DateFormat:="ISO", _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test56 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test57(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "stocks"
    FileName = "stocks.csv"
    Expected = HStack( _
        Array("Stock Name", "AXP", "BA", "CAT", "CSC", "CVX", "DD", "DIS", "GE", "GS", "HD", "IBM", "INTC", "JNJ", "JPM", "KO", "MCD", "MMM", "MRK", "MSFT", "NKE", "PFE", "PG", "T", "TRV", "UNH", "UTX", "V", "VZ", "WMT", "XOM"), _
        Array("Company Name", "American Express Co", "Boeing Co", "Caterpillar Inc", "Cisco Systems Inc", "Chevron Corp", "Dupont E I De Nemours & Co", "Walt Disney Co", "General Electric Co", "Goldman Sachs Group Inc", _
        "Home Depot Inc", "International Business Machines Co...", "Intel Corp", "Johnson & Johnson", "JPMorgan Chase and Co", "The Coca-Cola Co", "McDonald's Corp", "3M Co", "Merck & Co Inc", "Microsoft Corp", "Nike Inc", "Pfizer Inc", _
        "Procter & Gamble Co", "AT&T Inc", "Travelers Companies Inc", "UnitedHealth Group Inc", "United Technologies Corp", "Visa Inc", "Verizon Communications Inc", "Wal-Mart Stores Inc", "Exxon Mobil Corp"))

    TestRes = TestCSVRead(57, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:="T", _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test57 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test58(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    'Tests handling of lines that start with a delimiter when IgnoreRepeated = true
    TestDescription = "test repeated delim 371"
    FileName = "test_repeated_delim_371.csv"
    Expected = HStack( _
        Array("FAMILY", "A", "A", "A", "A", "A", "A", "EPGP013951", "EPGP014065", "EPGP014065", "EPGP014065", "EP07", "83346_EPGP014244", "83346_EPGP014244", "83506", "87001"), _
        Array("PERSON", "EP01223", "EP01227", "EP01228", "EP01228", "EP01227", "EP01228", "EPGP013952", "EPGP014066", "EPGP014065", "EPGP014068", "706", "T3011", "T3231", "T17255", "301"), _
        Array("MARKER", "rs710865", "rs11249215", "rs11249215", "rs10903129", "rs621559", "rs1514175", "rs773564", "rs2794520", "rs296547", "rs296547", "rs10927875", "rs2251760", "rs2251760", "rs2475335", "rs2413583"), _
        Array("RATIO", "0.0214", "0.0107", "0.00253", "0.0116", "0.00842", "0.0202", "0.00955", "0.0193", "0.0135", "0.0239", "0.0157", "0.0154", "0.0154", "0.00784", "0.0112"))

    TestRes = TestCSVRead(58, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, Delimiter:=" ", IgnoreRepeated:=True, ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test58 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test59(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "TechCrunchcontinentalUSA"
    Expected = HStack( _
        Array("permalink", "lifelock", "lifelock"), _
        Array("company", "LifeLock", "LifeLock"), _
        Array("numEmps", Empty, Empty), _
        Array("category", "web", "web"), _
        Array("city", "Tempe", "Tempe"), _
        Array("state", "AZ", "AZ"), _
        Array("fundedDate", 39203#, 38991#), _
        Array("raisedAmt", 6850000#, 6000000#), _
        Array("raisedCurrency", "USD", "USD"), _
        Array("round", "b", "a"))
    FileName = "TechCrunchcontinentalUSA.csv"
    TestRes = TestCSVRead(59, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:="ND", _
        DateFormat:="D-M-Y", _
        NumRows:=3, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test59 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test60(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "issue 120"
    FileName = "issue_120.csv"
    Expected = HStack( _
        Array(3528489623.48857, 3528489624.48866, 3528489625.48857, 3528489626.48866, 3528489627.48875), _
        Array(312.73, 312.49, 312.74, 312.49, 312.62), _
        Array(0#, 0#, 0#, 0#, 0#), _
        Array(41.87425, 41.87623, 41.87155, 41.86422, 41.87615), _
        Array(297.6302, 297.6342, 297.6327, 297.632, 297.6324), _
        Array(0#, 0#, 0#, 0#, 0#), _
        Array(286.3423, 286.3563, 286.3723, 286.3837, 286.397), _
        Array(-99.99, -99.99, -99.99, -99.99, -99.99), _
        Array(-99.99, -99.99, -99.99, -99.99, -99.99), _
        Array(12716#, 12716#, 12716#, 12716#, 12716#), _
        Array(0#, 0#, 0#, 0#, 0#), _
        Array(0#, 0#, 0#, 0#, 0#), _
        Array(0#, 0#, 0#, 0#, 0#), _
        Array(Empty, Empty, Empty, Empty, Empty), _
        Array(-24.81942, -24.8206, -24.82111, -24.82091, -24.82035), _
        Array(853.8073, 852.1921, 853.4257, 854.1342, 851.171), _
        Array(0#, 0#, 0#, 0#, 0#), _
        Array(0#, 0#, 0#, 0#, 0#), _
        Array(60.07, 38.27, 61.38, 49.23, 42.49), _
        Array(132.356, 132.356, 132.356, 132.356, 132.356))

    TestRes = TestCSVRead(60, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:=True, ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test60 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test61(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    'Tests trimming fields
    TestDescription = "census.txt"
    FileName = "census.txt"
    Expected = HStack( _
        Array("GEOID", 601#, 602#, 603#), _
        Array("POP10", 18570#, 41520#, 54689#), _
        Array("HU10", 7744#, 18073#, 25653#), _
        Array("ALAND", 166659789#, 79288158#, 81880442#), _
        Array("AWATER", 799296#, 4446273#, 183425#), _
        Array("ALAND_SQMI", 64.348, 30.613, 31.614), _
        Array("AWATER_SQMI", 0.309, 1.717, 0.071), _
        Array("INTPTLAT", 18.180555, 18.362268, 18.455183), _
        Array("INTPTLONG", -66.749961, -67.17613, -67.119887))

    TestRes = TestCSVRead(61, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:="NT", Delimiter:=vbTab, ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test61 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test62(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "double quote quotechar and escapechar"
    FileName = "double_quote_quotechar_and_escapechar.csv"
    Expected = HStack( _
        Array("APINo", 33101000000000#, 33001000000000#, 33009000000000#, 33043000000000#, 33031000000000#, 33023000000000#, 33055000000000#, 33043000000000#, 33075000000000#, 33101000000000#, 33047000000000#, 33105000000000#, 33105000000000#, 33059000000000#, 33065000000000#, 33029000000000#, 33077000000000#, 33101000000000#, 33015000000000#, 33071000000000#, 33057000000000#, 33055000000000#, 33029000000000#, 33043000000000#), _
        Array("FileNo", 1#, 2#, 3#, 4#, 5#, 6#, 7#, 8#, 9#, 10#, 11#, 12#, 13#, 14#, 15#, 16#, 17#, 18#, 19#, 20#, 21#, 22#, 23#, 24#), _
        Array("CurrentWellName", "BLUM     1", "DAVIS WELL     1", "GREAT NORTH. O AND G PIPELINE CO.     1", "ROBINSON PATD LAND     1", "GLENFIELD OIL COMPANY     1", "NORTHWEST OIL CO.     1", "OIL SYNDICATE     1", "ARMSTRONG     1", "GEHRINGER     1", "PETROLEUM CO.     1", "BURNSTAD     1", "OIL COMPANY     1", "NELS KAMP     1", "EXPLORATION-NORTH DAKOTA     1", "WACHTER     16-18", "FRANKLIN INVESTMENT CO.     1", "RUDDY BROS     1", "J. H. KLINE     1", "STRATIGRAPHIC TEST     1", "AANSTAD STRATIGRAPHIC TEST     1", "FRITZ LEUTZ     1", "VAUGHN HANSON     1", "J. J. WEBER     1", "NORTH DAKOTA STATE A     1"), _
        Array("LeaseName", "BLUM", "DAVIS WELL", "GREAT NORTH. O AND G PIPELINE CO.", "ROBINSON PATD LAND", "GLENFIELD OIL COMPANY", "NORTHWEST OIL CO.", "OIL SYNDICATE", "ARMSTRONG", "GEHRINGER", "PETROLEUM CO.", "BURNSTAD", "OIL COMPANY", "NELS KAMP", "EXPLORATION-NORTH DAKOTA", "WACHTER", "FRANKLIN INVESTMENT CO.", "RUDDY BROS", "J. H. KLINE", "STRATIGRAPHIC TEST", "AANSTAD STRATIGRAPHIC TEST", "FRITZ LEUTZ", "VAUGHN HANSON", "J. J. WEBER", "NORTH DAKOTA STATE A"), _
        Array("OriginalWellName", "PIONEER OIL & GAS #1", "DAVIS WELL #1", "GREAT NORTHERN OIL & GAS PIPELINE #1", "ROBINSON PAT'D LAND #1", "GLENFIELD OIL COMPANY #1", "#1", "H. HANSON OIL SYNDICATE #1", "ARMSTRONG #1", "GEHRINGER #1", "VELVA PETROLEUM CO. #1", "BURNSTAD #1", "BIG VIKING #1", "NELS KAMP #1", "EXPLORATION-NORTH DAKOTA #1", "E. L. SEMLING #1", "FRANKLIN INVESTMENT CO. #1", "RUDDY BROS #1", "J. H. KLINE #1", "STRATIGRAPHIC TEST #1", "AANSTAD STRATIGRAPHIC TEST #1", "FRITZ LEUTZ #1", "VAUGHN HANSON #1", "J. J. WEBER #1", "NORTH DAKOTA STATE ""A"" #1"))

    TestRes = TestCSVRead(62, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:=True, ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test62 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test63(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "baseball"
    FileName = "baseball.csv"
    Expected = HStack( _
        Array("Rk", 1#, 2#, 3#, 4#, 5#, Empty, 6#, 7#, 8#, 9#, Empty, 10#, 11#, 12#, 13#, 14#, 15#, 16#, 17#, 18#, 19#, 20#, 21#, 22#, 23#, 24#, 25#, 26#, 27#, 28#, 29#, 30#, Empty, Empty, Empty), _
        Array("Year", 1978#, 1979#, 1980#, 1981#, 1981#, Empty, 1982#, 1983#, 1984#, 1985#, Empty, 1990#, 1991#, 1992#, 1993#, 1994#, 1995#, 1996#, 1997#, 1998#, 1999#, 2000#, 2001#, 2002#, 2003#, 2004#, 2005#, 2006#, 2007#, 2008#, 2009#, 2010#, Empty, Empty, Empty), _
        Array("Age", 37#, 38#, 39#, 40#, 40#, Empty, 41#, 42#, 43#, 44#, Empty, 49#, 50#, 51#, 52#, 53#, 54#, 55#, 56#, 57#, 58#, 59#, 60#, 61#, 62#, 63#, 64#, 65#, 66#, 67#, 68#, 69#, Empty, Empty, Empty), _
        Array("Tm", "Atlanta Braves", "Atlanta Braves", "Atlanta Braves", "Atlanta Braves", "Atlanta Braves", Empty, "Toronto Blue Jays", "Toronto Blue Jays", "Toronto Blue Jays", "Toronto Blue Jays", Empty, "Atlanta Braves", "Atlanta Braves", "Atlanta Braves", "Atlanta Braves", "Atlanta Braves", "Atlanta Braves", "Atlanta Braves", "Atlanta Braves", "Atlanta Braves", "Atlanta Braves", "Atlanta Braves", "Atlanta Braves", "Atlanta Braves", "Atlanta Braves", "Atlanta Braves", "Atlanta Braves", "Atlanta Braves", "Atlanta Braves", "Atlanta Braves", "Atlanta Braves", "Atlanta Braves", "Toronto Blue Jays", "Atlanta Braves", Empty), _
        Array("Lg", "NL", "NL", "NL", "NL", "NL", Empty, "AL", "AL", "AL", "AL", Empty, "NL", "NL", "NL", "NL", "NL", "NL", "NL", "NL", "NL", "NL", "NL", "NL", "NL", "NL", "NL", "NL", "NL", "NL", "NL", "NL", "NL", Empty, Empty, Empty), _
        Array(Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, "2nd of 2", Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, "4 years", "25 years", "29 years"), _
        Array("W", 69#, 66#, 81#, 25#, 25#, Empty, 78#, 89#, 89#, 99#, Empty, 40#, 94#, 98#, 104#, 68#, 90#, 96#, 101#, 106#, 103#, 95#, 88#, 101#, 101#, 96#, 90#, 79#, 84#, 72#, 86#, 91#, 355#, 2149#, 2504#), _
        Array("L", 93#, 94#, 80#, 29#, 27#, Empty, 84#, 73#, 73#, 62#, Empty, 57#, 68#, 64#, 58#, 46#, 54#, 66#, 61#, 56#, 59#, 67#, 74#, 59#, 61#, 66#, 72#, 83#, 78#, 90#, 76#, 71#, 292#, 1709#, 2001#), _
        Array("W-L%", 0.426, 0.413, 0.503, 0.463, 0.481, Empty, 0.481, 0.549, 0.549, 0.615, Empty, 0.412, 0.58, 0.605, 0.642, 0.596, 0.625, 0.593, 0.623, 0.654, 0.636, 0.586, 0.543, 0.631, 0.623, 0.593, 0.556, 0.488, 0.519, 0.444, 0.531, 0.562, 0.549, 0.557, 0.556), _
        Array("G", 162#, 160#, 161#, 55#, 52#, Empty, 162#, 162#, 163#, 161#, Empty, 97#, 162#, 162#, 162#, 114#, 144#, 162#, 162#, 162#, 162#, 162#, 162#, 161#, 162#, 162#, 162#, 162#, 162#, 162#, 162#, 162#, 648#, 3860#, 4508#), _
        Array("Finish", 6#, 6#, 4#, 4#, 5#, Empty, 6#, 4#, 2#, 1#, Empty, 6#, 1#, 1#, 1#, 2#, 1#, 1#, 1#, 1#, 1#, 1#, 1#, 1#, 1#, 1#, 1#, 3#, 3#, 4#, 3#, 2#, 3.3, 2.4, 2.5), _
        Array("Wpost", 0#, 0#, 0#, 0#, 0#, Empty, 0#, 0#, 0#, 3#, Empty, 0#, 7#, 6#, 2#, 0#, 11#, 9#, 5#, 5#, 7#, 0#, 4#, 2#, 2#, 2#, 1#, 0#, 0#, 0#, 0#, 1#, 3#, 64#, 67#), _
        Array("Lpost", 0#, 0#, 0#, 0#, 0#, Empty, 0#, 0#, 0#, 4#, Empty, 0#, 7#, 7#, 4#, 0#, 3#, 7#, 4#, 4#, 7#, 3#, 4#, 3#, 3#, 3#, 3#, 0#, 0#, 0#, 0#, 3#, 4#, 65#, 69#), _
        Array("W-L%post", Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, 0.429, Empty, Empty, 0.5, 0.462, 0.333, Empty, 0.786, 0.562, 0.556, 0.556, 0.5, 0#, 0.5, 0.4, 0.4, 0.4, 0.25, Empty, Empty, Empty, Empty, 0.25, 0.429, 0.496, 0.493), _
        Array(Empty, Empty, Empty, Empty, "First half of season", "Second half of season", Empty, Empty, Empty, Empty, Empty, Empty, Empty, "NL Pennant", "NL Pennant", Empty, Empty, "WS Champs", "NL Pennant", Empty, Empty, "NL Pennant", Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, "5 Pennants and 1 World Series Title", "5 Pennants and 1 World Series Title"))

    TestRes = TestCSVRead(63, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:="N", ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test63 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test64(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test converttypes arg"
    FileName = "test_converttypes_arg.csv"
    Expected = HStack( _
        Array("Number", "Date", "Boolean", "Error", "String", "String", "String", "String", "String", "String"), _
        Array(44424#, CDate("2021-Aug-18"), True, "#DIV/0!", "1", "16-Aug-2021", "TRUE", "#DIV/0!", "abc", "abc""def"))

    TestRes = TestCSVRead(64, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        DateFormat:="Y-M-D", _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test64 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test65(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test converttypes arg"
    FileName = "test_converttypes_arg.csv"
    Expected = HStack( _
        Array("Number", "Date", "Boolean", "Error", "String", "String", "String", "String", "String", "String"), _
        Array("44424", "2021-08-18", "True", "#DIV/0!", "1", "16-Aug-2021", "TRUE", "#DIV/0!", "abc", "abc""def"))

    TestRes = TestCSVRead(65, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test65 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test66(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test converttypes arg"
    FileName = "test_converttypes_arg.csv"
    Expected = HStack( _
        Array("Number", "Date", "Boolean", "Error", "String", "String", "String", "String", "String", "String"), _
        Array(44424#, "2021-08-18", "True", "#DIV/0!", "1", "16-Aug-2021", "TRUE", "#DIV/0!", "abc", "abc""def"))

    TestRes = TestCSVRead(66, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:="N", ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test66 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test67(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test converttypes arg"
    FileName = "test_converttypes_arg.csv"
    Expected = HStack( _
        Array("Number", "Date", "Boolean", "Error", "String", "String", "String", "String", "String", "String"), _
        Array("44424", CDate("2021-Aug-18"), "True", "#DIV/0!", "1", "16-Aug-2021", "TRUE", "#DIV/0!", "abc", "abc""def"))

    TestRes = TestCSVRead(67, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:="D", DateFormat:="Y-M-D", ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test67 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test68(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test converttypes arg"
    FileName = "test_converttypes_arg.csv"
    Expected = HStack( _
        Array("Number", "Date", "Boolean", "Error", "String", "String", "String", "String", "String", "String"), _
        Array("44424", "2021-08-18", True, "#DIV/0!", "1", "16-Aug-2021", "TRUE", "#DIV/0!", "abc", "abc""def"))

    TestRes = TestCSVRead(68, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:="B", ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test68 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test69(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test converttypes arg"
    FileName = "test_converttypes_arg.csv"
    Expected = HStack( _
        Array("Number", "Date", "Boolean", "Error", "String", "String", "String", "String", "String", "String"), _
        Array("44424", "2021-08-18", "True", CVErr(2007), "1", "16-Aug-2021", "TRUE", "#DIV/0!", "abc", "abc""def"))

    TestRes = TestCSVRead(69, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:="E", ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test69 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test70(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test converttypes arg"
    FileName = "test_converttypes_arg.csv"
    Expected = HStack( _
        Array("Number", "Date", "Boolean", "Error", "String", "String", "String", "String", "String", "String"), _
        Array(44424#, "2021-08-18", "True", "#DIV/0!", 1#, "16-Aug-2021", "TRUE", "#DIV/0!", "abc", "abc""def"))

    TestRes = TestCSVRead(70, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:="NQ", ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test70 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test71(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test converttypes arg"
    FileName = "test_converttypes_arg.csv"
    Expected = HStack( _
        Array("Number", "Date", "Boolean", "Error", "String", "String", "String", "String", "String", "String"), _
        Array("44424", CDate("2021-Aug-18"), "True", "#DIV/0!", "1", CDate("2021-Aug-16"), "TRUE", "#DIV/0!", "abc", "abc""def"))

    TestRes = TestCSVRead(71, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:="DQ", _
        DateFormat:="Y-M-D", _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test71 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test72(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test converttypes arg"
    FileName = "test_converttypes_arg.csv"
    Expected = HStack( _
        Array("Number", "Date", "Boolean", "Error", "String", "String", "String", "String", "String", "String"), _
        Array("44424", "2021-08-18", True, "#DIV/0!", "1", "16-Aug-2021", True, "#DIV/0!", "abc", "abc""def"))

    TestRes = TestCSVRead(72, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:="BQ", ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test72 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test73(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test converttypes arg"
    FileName = "test_converttypes_arg.csv"
    Expected = HStack( _
        Array("Number", "Date", "Boolean", "Error", "String", "String", "String", "String", "String", "String"), _
        Array("44424", "2021-08-18", "True", CVErr(2007), "1", "16-Aug-2021", "TRUE", CVErr(2007), "abc", "abc""def"))

    TestRes = TestCSVRead(73, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:="EQ", ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test73 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test74(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test converttypes arg"
    FileName = "test_converttypes_arg.csv"
    Expected = HStack( _
        Array("Number", "Date", "Boolean", "Error", "String", "String", "String", "String", "String", "String"), _
        Array(44424#, CDate("2021-Aug-18"), True, CVErr(2007), "1", "16-Aug-2021", "TRUE", "#DIV/0!", "abc", "abc""def"))

    TestRes = TestCSVRead(74, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, ConvertTypes:="NDBE", DateFormat:="Y-M-D", ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test74 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test75(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test converttypes arg"
    FileName = "test_converttypes_arg.csv"
    Expected = HStack( _
        Array("Number", "Date", "Boolean", "Error", "String", "String", "String", "String", "String", "String"), _
        Array(44424#, CDate("2021-Aug-18"), True, CVErr(2007), 1#, CDate("2021-Aug-16"), True, CVErr(2007), "abc", "abc""def"))

    TestRes = TestCSVRead(75, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:="NDBEQ", _
        DateFormat:="Y-M-D", _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test75 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test76(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "latest (1)"
    FileName = "latest (1).csv"
    Expected = Empty
    TestRes = TestCSVRead(76, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:="ND", _
        DateFormat:="ISO", _
        MissingStrings:="\N", _
        ShowMissingsAs:=Empty, _
        NumRowsExpected:=1000, _
        NumColsExpected:=25)
    If TestRes Then
        'Same test as here:
        'https://github.com/JuliaData/CSV.jl/blob/953636a363525e3027d690b8a30448d115249bf9/test/testfiles.jl#L317
        TestRes = IsEmpty(Observed(UBound(Observed, 1) - 2, LBound(Observed, 2) + 16))
        If Not TestRes Then WhatDiffers = "Test 76 latest (1) FAILED, Test was that element in 17th col, last but 2 row should be empty"
    End If
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test76 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test77(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "int64 overflow"
    FileName = "int64_overflow.csv"
    Expected = HStack(Array("col1", 1#, 2#, 3#, 9.22337203685478E+18))
    TestRes = TestCSVRead(77, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        ShowMissingsAs:=Empty, _
        RelTol:=0.000000000000001)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test77 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test78(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "FL insurance sample"
    FileName = "FL_insurance_sample.csv"
    Expected = HStack( _
        Array("policyID", 119736#), _
        Array("statecode", "FL"), _
        Array("county", "CLAY COUNTY"), _
        Array("eq_site_limit", 498960#), _
        Array("hu_site_limit", 498960#), _
        Array("fl_site_limit", 498960#), _
        Array("fr_site_limit", 498960#), _
        Array("tiv_2011", 498960#), _
        Array("tiv_2012", 792148.9), _
        Array("eq_site_deductible", 0#), _
        Array("hu_site_deductible", 9979.2), _
        Array("fl_site_deductible", 0#), _
        Array("fr_site_deductible", 0#), _
        Array("point_latitude", 30.102261), _
        Array("point_longitude", -81.711777), _
        Array("line", "Residential"), _
        Array("construction", "Masonry"), _
        Array("point_granularity", 1#))

    TestRes = TestCSVRead(78, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:="N", _
        NumRows:=2, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test78 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test79(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "FL insurance sample"
    FileName = "FL_insurance_sample.csv"
    Expected = Empty
    TestRes = TestCSVRead(79, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:="N", _
        ShowMissingsAs:=Empty, _
        NumRowsExpected:=36635, _
        NumColsExpected:=18)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test79 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test80(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test float in int column"
    FileName = "test_float_in_int_column.csv"
    Expected = HStack(Array("col1", 1#, 4#, 7#), Array("col2", 2#, 5.4, 8#), Array("col3", 3#, 6#, 9#))
    TestRes = TestCSVRead(80, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test80 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test81(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test skip args"
    FileName = "test_skip_args.csv"
    Expected = HStack(Array("3,3", "4,3", "5,3", "6,3", "7,3", "8,3", "9,3", "10,3", Empty, Empty))
    TestRes = TestCSVRead(81, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, SkipToRow:=3, SkipToCol:=3, NumRows:=10, NumCols:=1, ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test81 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test82(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test skip args"
    FileName = "test_skip_args.csv"
    Expected = HStack("6,5", "6,6", "6,7", "6,8", "6,9", "6,10", Empty, Empty)
    TestRes = TestCSVRead(82, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, SkipToRow:=6, SkipToCol:=5, NumRows:=1, NumCols:=8, ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test82 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test83(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test skip args"
    FileName = "test_skip_args.csv"
    Expected = HStack( _
        Array("8,8", "9,8", "10,8", Empty), _
        Array("8,9", "9,9", "10,9", Empty), _
        Array("8,10", "9,10", "10,10", Empty), _
        Array(Empty, Empty, Empty, Empty))

    TestRes = TestCSVRead(83, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, SkipToRow:=8, SkipToCol:=8, NumRows:=4, NumCols:=4, ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test83 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test84(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test skip args with comments"
    FileName = "test_skip_args_with_comments.csv"
    Expected = HStack(Array("3,3", "4,3", "5,3", "6,3", "7,3", "8,3", "9,3", "10,3", Empty, Empty))
    TestRes = TestCSVRead(84, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, Comment:="#", SkipToRow:=3, SkipToCol:=3, NumRows:=10, NumCols:=1, ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test84 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test85(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test skip args with comments"
    FileName = "test_skip_args_with_comments.csv"
    Expected = HStack("6,5", "6,6", "6,7", "6,8", "6,9", "6,10", Empty, Empty)
    TestRes = TestCSVRead(85, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, Comment:="#", SkipToRow:=6, SkipToCol:=5, NumRows:=1, NumCols:=8, ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test85 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test86(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test skip args with comments"
    FileName = "test_skip_args_with_comments.csv"
    Expected = HStack( _
        Array("8,8", "9,8", "10,8", Empty), _
        Array("8,9", "9,9", "10,9", Empty), _
        Array("8,10", "9,10", "10,10", Empty), _
        Array(Empty, Empty, Empty, Empty))

    TestRes = TestCSVRead(86, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, Comment:="#", SkipToRow:=8, SkipToCol:=8, NumRows:=4, NumCols:=4, ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test86 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test87(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test triangular"
    FileName = "test_triangular.csv"
    Expected = HStack( _
        Array(1#, 1#, 1#, 1#, 1#, 1#, 1#, 1#, 1#, 1#), _
        Array(Empty, 1#, 1#, 1#, 1#, 1#, 1#, 1#, 1#, 1#), _
        Array(Empty, Empty, 1#, 1#, 1#, 1#, 1#, 1#, 1#, 1#), _
        Array(Empty, Empty, Empty, 1#, 1#, 1#, 1#, 1#, 1#, 1#), _
        Array(Empty, Empty, Empty, Empty, 1#, 1#, 1#, 1#, 1#, 1#), _
        Array(Empty, Empty, Empty, Empty, Empty, 1#, 1#, 1#, 1#, 1#), _
        Array(Empty, Empty, Empty, Empty, Empty, Empty, 1#, 1#, 1#, 1#), _
        Array(Empty, Empty, Empty, Empty, Empty, Empty, Empty, 1#, 1#, 1#), _
        Array(Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, 1#, 1#), _
        Array(Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, 1#))

    TestRes = TestCSVRead(87, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test87 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test88(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test strange delimiter"
    FileName = "test_strange_delimiter.csv"
    Expected = HStack( _
        Array(1#, 6#, 11#, 16#, 21#, 26#, 31#, 36#, 41#, 46#), _
        Array(2#, 7#, 12#, 17#, 22#, 27#, 32#, 37#, 42#, 47#), _
        Array(3#, 8#, 13#, 18#, 23#, 28#, 33#, 38#, 43#, 48#), _
        Array(4#, 9#, 14#, 19#, 24#, 29#, 34#, 39#, 44#, 49#), _
        Array(5#, 10#, 15#, 20#, 25#, 30#, 35#, 40#, 45#, 50#))

    TestRes = TestCSVRead(88, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        Delimiter:="{""}", _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test88 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test89(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test ignoring repeated multicharacter delimiter"
    FileName = "test_ignoring_repeated_multicharacter_delimiter.csv"
    Expected = HStack( _
        Array(1#, 6#, 11#, 16#, 21#, 26#, 31#, 36#, 41#, 46#), _
        Array(2#, 7#, 12#, 17#, 22#, 27#, 32#, 37#, 42#, 47#), _
        Array(3#, 8#, 13#, 18#, 23#, 28#, 33#, 38#, 43#, 48#), _
        Array(4#, 9#, 14#, 19#, 24#, 29#, 34#, 39#, 44#, 49#), _
        Array(5#, 10#, 15#, 20#, 25#, 30#, 35#, 40#, 45#, 50#))

    TestRes = TestCSVRead(89, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        Delimiter:="Delim", _
        IgnoreRepeated:=True, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test89 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test90(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test empty file"
    FileName = "test_empty_file.csv"
    Expected = "#CSVRead: #InferDelimiter: File is empty!!"
    TestRes = TestCSVRead(90, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test90 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test91(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "table test.txt"
    FileName = "table_test.txt"
    Expected = VStack(Array("ind_50km", "nse_gsurf_cfg1", "r_gsurf_cfg1", "bias_gsurf_cfg1", "ngrids", "nse_hatmo_cfg1", "r_hatmo_cfg1", "bias_hatmo_cfg1", "nse_latmo_cfg1", "r_latmo_cfg1", "bias_latmo_cfg1", "nse_melt_cfg1", "r_melt_cfg1", "bias_melt_cfg1", "nse_rnet_cfg1", "r_rnet_cfg1", "bias_rnet_cfg1", "nse_rof_cfg1", "r_rof_cfg1", "bias_rof_cfg1", "nse_snowdepth_cfg1", "r_snowdepth_cfg1", "bias_snowdepth_cfg1", "nse_swe_cfg1", "r_swe_cfg1", "bias_swe_cfg1", "nse_gsurf_cfg2", "r_gsurf_cfg2", "bias_gsurf_cfg2", "nse_hatmo_cfg2", "r_hatmo_cfg2", "bias_hatmo_cfg2", "nse_latmo_cfg2", "r_latmo_cfg2", "bias_latmo_cfg2", _
        "nse_melt_cfg2", "r_melt_cfg2", "bias_melt_cfg2", "nse_rnet_cfg2", "r_rnet_cfg2", "bias_rnet_cfg2", "nse_rof_cfg2", "r_rof_cfg2", "bias_rof_cfg2", "nse_snowdepth_cfg2", "r_snowdepth_cfg2", "bias_snowdepth_cfg2", "nse_swe_cfg2", "r_swe_cfg2", "bias_swe_cfg2", "nse_gsurf_cfg3", "r_gsurf_cfg3", "bias_gsurf_cfg3", "nse_hatmo_cfg3", "r_hatmo_cfg3", "bias_hatmo_cfg3", "nse_latmo_cfg3", "r_latmo_cfg3", "bias_latmo_cfg3", "nse_melt_cfg3", "r_melt_cfg3", "bias_melt_cfg3", "nse_rnet_cfg3", "r_rnet_cfg3", "bias_rnet_cfg3", "nse_rof_cfg3", "r_rof_cfg3", "bias_rof_cfg3", "nse_snowdepth_cfg3", "r_snowdepth_cfg3", _
        "bias_snowdepth_cfg3", "nse_swe_cfg3", "r_swe_cfg3", "bias_swe_cfg3", "nse_gsurf_cfg4", "r_gsurf_cfg4", "bias_gsurf_cfg4", "nse_hatmo_cfg4", "r_hatmo_cfg4", "bias_hatmo_cfg4", "nse_latmo_cfg4", "r_latmo_cfg4", "bias_latmo_cfg4", "nse_melt_cfg4", "r_melt_cfg4", "bias_melt_cfg4", "nse_rnet_cfg4", "r_rnet_cfg4", "bias_rnet_cfg4", "nse_rof_cfg4", "r_rof_cfg4", "bias_rof_cfg4", "nse_snowdepth_cfg4", "r_snowdepth_cfg4", "bias_snowdepth_cfg4", "nse_swe_cfg4", "r_swe_cfg4", "bias_swe_cfg4", "nse_gsurf_cfg5", "r_gsurf_cfg5", "bias_gsurf_cfg5", "nse_hatmo_cfg5", "r_hatmo_cfg5", "bias_hatmo_cfg5", "nse_latmo_cfg5", _
        "r_latmo_cfg5", "bias_latmo_cfg5", "nse_melt_cfg5", "r_melt_cfg5", "bias_melt_cfg5", "nse_rnet_cfg5", "r_rnet_cfg5", "bias_rnet_cfg5", "nse_rof_cfg5", "r_rof_cfg5", "bias_rof_cfg5", "nse_snowdepth_cfg5", "r_snowdepth_cfg5", "bias_snowdepth_cfg5", "nse_swe_cfg5", "r_swe_cfg5", "bias_swe_cfg5", "nse_gsurf_cfg6", "r_gsurf_cfg6", "bias_gsurf_cfg6", "nse_hatmo_cfg6", "r_hatmo_cfg6", "bias_hatmo_cfg6", "nse_latmo_cfg6", "r_latmo_cfg6", "bias_latmo_cfg6", "nse_melt_cfg6", "r_melt_cfg6", "bias_melt_cfg6", "nse_rnet_cfg6", "r_rnet_cfg6", "bias_rnet_cfg6", "nse_rof_cfg6", "r_rof_cfg6", "bias_rof_cfg6", _
        "nse_snowdepth_cfg6", "r_snowdepth_cfg6", "bias_snowdepth_cfg6", "nse_swe_cfg6", "r_swe_cfg6", "bias_swe_cfg6", "nse_gsurf_cfg7", "r_gsurf_cfg7", "bias_gsurf_cfg7", "nse_hatmo_cfg7", "r_hatmo_cfg7", "bias_hatmo_cfg7", "nse_latmo_cfg7", "r_latmo_cfg7", "bias_latmo_cfg7", "nse_melt_cfg7", "r_melt_cfg7", "bias_melt_cfg7", "nse_rnet_cfg7", "r_rnet_cfg7", "bias_rnet_cfg7", "nse_rof_cfg7", "r_rof_cfg7", "bias_rof_cfg7", "nse_snowdepth_cfg7", "r_snowdepth_cfg7", "bias_snowdepth_cfg7", "nse_swe_cfg7", "r_swe_cfg7", "bias_swe_cfg7", "nse_gsurf_cfg8", "r_gsurf_cfg8", "bias_gsurf_cfg8", "nse_hatmo_cfg8", "r_hatmo_cfg8", _
        "bias_hatmo_cfg8", "nse_latmo_cfg8", "r_latmo_cfg8", "bias_latmo_cfg8", "nse_melt_cfg8", "r_melt_cfg8", "bias_melt_cfg8", "nse_rnet_cfg8", "r_rnet_cfg8", "bias_rnet_cfg8", "nse_rof_cfg8", "r_rof_cfg8", "bias_rof_cfg8", "nse_snowdepth_cfg8", "r_snowdepth_cfg8", "bias_snowdepth_cfg8", "nse_swe_cfg8", "r_swe_cfg8", "bias_swe_cfg8", "nse_gsurf_cfg9", "r_gsurf_cfg9", "bias_gsurf_cfg9", "nse_hatmo_cfg9", "r_hatmo_cfg9", "bias_hatmo_cfg9", "nse_latmo_cfg9", "r_latmo_cfg9", "bias_latmo_cfg9", "nse_melt_cfg9", "r_melt_cfg9", "bias_melt_cfg9", "nse_rnet_cfg9", "r_rnet_cfg9", "bias_rnet_cfg9", "nse_rof_cfg9", _
        "r_rof_cfg9", "bias_rof_cfg9", "nse_snowdepth_cfg9", "r_snowdepth_cfg9", "bias_snowdepth_cfg9", "nse_swe_cfg9", "r_swe_cfg9", "bias_swe_cfg9", "nse_gsurf_cfg10", "r_gsurf_cfg10", "bias_gsurf_cfg10", "nse_hatmo_cfg10", "r_hatmo_cfg10", "bias_hatmo_cfg10", "nse_latmo_cfg10", "r_latmo_cfg10", "bias_latmo_cfg10", "nse_melt_cfg10", "r_melt_cfg10", "bias_melt_cfg10", "nse_rnet_cfg10", "r_rnet_cfg10", "bias_rnet_cfg10", "nse_rof_cfg10", "r_rof_cfg10", "bias_rof_cfg10", "nse_snowdepth_cfg10", "r_snowdepth_cfg10", "bias_snowdepth_cfg10", "nse_swe_cfg10", "r_swe_cfg10", "bias_swe_cfg10", "nse_gsurf_cfg11", "r_gsurf_cfg11", "bias_gsurf_cfg11", _
        "nse_hatmo_cfg11", "r_hatmo_cfg11", "bias_hatmo_cfg11", "nse_latmo_cfg11", "r_latmo_cfg11", "bias_latmo_cfg11", "nse_melt_cfg11", "r_melt_cfg11", "bias_melt_cfg11", "nse_rnet_cfg11", "r_rnet_cfg11", "bias_rnet_cfg11", "nse_rof_cfg11", "r_rof_cfg11", "bias_rof_cfg11", "nse_snowdepth_cfg11", "r_snowdepth_cfg11", "bias_snowdepth_cfg11", "nse_swe_cfg11", "r_swe_cfg11", "bias_swe_cfg11", "nse_gsurf_cfg12", "r_gsurf_cfg12", "bias_gsurf_cfg12", "nse_hatmo_cfg12", "r_hatmo_cfg12", "bias_hatmo_cfg12", "nse_latmo_cfg12", "r_latmo_cfg12", "bias_latmo_cfg12", "nse_melt_cfg12", "r_melt_cfg12", "bias_melt_cfg12", "nse_rnet_cfg12", "r_rnet_cfg12", _
        "bias_rnet_cfg12", "nse_rof_cfg12", "r_rof_cfg12", "bias_rof_cfg12", "nse_snowdepth_cfg12", "r_snowdepth_cfg12", "bias_snowdepth_cfg12", "nse_swe_cfg12", "r_swe_cfg12", "bias_swe_cfg12", "nse_gsurf_cfg13", "r_gsurf_cfg13", "bias_gsurf_cfg13", "nse_hatmo_cfg13", "r_hatmo_cfg13", "bias_hatmo_cfg13", "nse_latmo_cfg13", "r_latmo_cfg13", "bias_latmo_cfg13", "nse_melt_cfg13", "r_melt_cfg13", "bias_melt_cfg13", "nse_rnet_cfg13", "r_rnet_cfg13", "bias_rnet_cfg13", "nse_rof_cfg13", "r_rof_cfg13", "bias_rof_cfg13", "nse_snowdepth_cfg13", "r_snowdepth_cfg13", "bias_snowdepth_cfg13", "nse_swe_cfg13", "r_swe_cfg13", "bias_swe_cfg13", "nse_gsurf_cfg14", _
        "r_gsurf_cfg14", "bias_gsurf_cfg14", "nse_hatmo_cfg14", "r_hatmo_cfg14", "bias_hatmo_cfg14", "nse_latmo_cfg14", "r_latmo_cfg14", "bias_latmo_cfg14", "nse_melt_cfg14", "r_melt_cfg14", "bias_melt_cfg14", "nse_rnet_cfg14", "r_rnet_cfg14", "bias_rnet_cfg14", "nse_rof_cfg14", "r_rof_cfg14", "bias_rof_cfg14", "nse_snowdepth_cfg14", "r_snowdepth_cfg14", "bias_snowdepth_cfg14", "nse_swe_cfg14", "r_swe_cfg14", "bias_swe_cfg14", "nse_gsurf_cfg15", "r_gsurf_cfg15", "bias_gsurf_cfg15", "nse_hatmo_cfg15", "r_hatmo_cfg15", "bias_hatmo_cfg15", "nse_latmo_cfg15", "r_latmo_cfg15", "bias_latmo_cfg15", "nse_melt_cfg15", "r_melt_cfg15", "bias_melt_cfg15", _
        "nse_rnet_cfg15", "r_rnet_cfg15", "bias_rnet_cfg15", "nse_rof_cfg15", "r_rof_cfg15", "bias_rof_cfg15", "nse_snowdepth_cfg15", "r_snowdepth_cfg15", "bias_snowdepth_cfg15", "nse_swe_cfg15", "r_swe_cfg15", "bias_swe_cfg15", "nse_gsurf_cfg16", "r_gsurf_cfg16", "bias_gsurf_cfg16", "nse_hatmo_cfg16", "r_hatmo_cfg16", "bias_hatmo_cfg16", "nse_latmo_cfg16", "r_latmo_cfg16", "bias_latmo_cfg16", "nse_melt_cfg16", "r_melt_cfg16", "bias_melt_cfg16", "nse_rnet_cfg16", "r_rnet_cfg16", "bias_rnet_cfg16", "nse_rof_cfg16", "r_rof_cfg16", "bias_rof_cfg16", "nse_snowdepth_cfg16", "r_snowdepth_cfg16", "bias_snowdepth_cfg16", "nse_swe_cfg16", "r_swe_cfg16", _
        "bias_swe_cfg16", "nse_gsurf_cfg17", "r_gsurf_cfg17", "bias_gsurf_cfg17", "nse_hatmo_cfg17", "r_hatmo_cfg17", "bias_hatmo_cfg17", "nse_latmo_cfg17", "r_latmo_cfg17", "bias_latmo_cfg17", "nse_melt_cfg17", "r_melt_cfg17", "bias_melt_cfg17", "nse_rnet_cfg17", "r_rnet_cfg17", "bias_rnet_cfg17", "nse_rof_cfg17", "r_rof_cfg17", "bias_rof_cfg17", "nse_snowdepth_cfg17", "r_snowdepth_cfg17", "bias_snowdepth_cfg17", "nse_swe_cfg17", "r_swe_cfg17", "bias_swe_cfg17", "nse_gsurf_cfg18", "r_gsurf_cfg18", "bias_gsurf_cfg18", "nse_hatmo_cfg18", "r_hatmo_cfg18", "bias_hatmo_cfg18", "nse_latmo_cfg18", "r_latmo_cfg18", "bias_latmo_cfg18", "nse_melt_cfg18", _
        "r_melt_cfg18", "bias_melt_cfg18", "nse_rnet_cfg18", "r_rnet_cfg18", "bias_rnet_cfg18", "nse_rof_cfg18", "r_rof_cfg18", "bias_rof_cfg18", "nse_snowdepth_cfg18", "r_snowdepth_cfg18", "bias_snowdepth_cfg18", "nse_swe_cfg18", "r_swe_cfg18", "bias_swe_cfg18", "nse_gsurf_cfg19", "r_gsurf_cfg19", "bias_gsurf_cfg19", "nse_hatmo_cfg19", "r_hatmo_cfg19", "bias_hatmo_cfg19", "nse_latmo_cfg19", "r_latmo_cfg19", "bias_latmo_cfg19", "nse_melt_cfg19", "r_melt_cfg19", "bias_melt_cfg19", "nse_rnet_cfg19", "r_rnet_cfg19", "bias_rnet_cfg19", "nse_rof_cfg19", "r_rof_cfg19", "bias_rof_cfg19", "nse_snowdepth_cfg19", "r_snowdepth_cfg19", "bias_snowdepth_cfg19", _
        "nse_swe_cfg19", "r_swe_cfg19", "bias_swe_cfg19", "nse_gsurf_cfg20", "r_gsurf_cfg20", "bias_gsurf_cfg20", "nse_hatmo_cfg20", "r_hatmo_cfg20", "bias_hatmo_cfg20", "nse_latmo_cfg20", "r_latmo_cfg20", "bias_latmo_cfg20", "nse_melt_cfg20", "r_melt_cfg20", "bias_melt_cfg20", "nse_rnet_cfg20", "r_rnet_cfg20", "bias_rnet_cfg20", "nse_rof_cfg20", "r_rof_cfg20", "bias_rof_cfg20", "nse_snowdepth_cfg20", "r_snowdepth_cfg20", "bias_snowdepth_cfg20", "nse_swe_cfg20", "r_swe_cfg20", "bias_swe_cfg20", "nse_gsurf_cfg21", "r_gsurf_cfg21", "bias_gsurf_cfg21", "nse_hatmo_cfg21", "r_hatmo_cfg21", "bias_hatmo_cfg21", "nse_latmo_cfg21", "r_latmo_cfg21", _
        "bias_latmo_cfg21", "nse_melt_cfg21", "r_melt_cfg21", "bias_melt_cfg21", "nse_rnet_cfg21", "r_rnet_cfg21", "bias_rnet_cfg21", "nse_rof_cfg21", "r_rof_cfg21", "bias_rof_cfg21", "nse_snowdepth_cfg21", "r_snowdepth_cfg21", "bias_snowdepth_cfg21", "nse_swe_cfg21", "r_swe_cfg21", "bias_swe_cfg21", "nse_gsurf_cfg22", "r_gsurf_cfg22", "bias_gsurf_cfg22", "nse_hatmo_cfg22", "r_hatmo_cfg22", "bias_hatmo_cfg22", "nse_latmo_cfg22", "r_latmo_cfg22", "bias_latmo_cfg22", "nse_melt_cfg22", "r_melt_cfg22", "bias_melt_cfg22", "nse_rnet_cfg22", "r_rnet_cfg22", "bias_rnet_cfg22", "nse_rof_cfg22", "r_rof_cfg22", "bias_rof_cfg22", "nse_snowdepth_cfg22", _
        "r_snowdepth_cfg22", "bias_snowdepth_cfg22", "nse_swe_cfg22", "r_swe_cfg22", "bias_swe_cfg22", "nse_gsurf_cfg23", "r_gsurf_cfg23", "bias_gsurf_cfg23", "nse_hatmo_cfg23", "r_hatmo_cfg23", "bias_hatmo_cfg23", "nse_latmo_cfg23", "r_latmo_cfg23", "bias_latmo_cfg23", "nse_melt_cfg23", "r_melt_cfg23", "bias_melt_cfg23", "nse_rnet_cfg23", "r_rnet_cfg23", "bias_rnet_cfg23", "nse_rof_cfg23", "r_rof_cfg23", "bias_rof_cfg23", "nse_snowdepth_cfg23", "r_snowdepth_cfg23", "bias_snowdepth_cfg23", "nse_swe_cfg23", "r_swe_cfg23", "bias_swe_cfg23", "nse_gsurf_cfg24", "r_gsurf_cfg24", "bias_gsurf_cfg24", "nse_hatmo_cfg24", "r_hatmo_cfg24", "bias_hatmo_cfg24", _
        "nse_latmo_cfg24", "r_latmo_cfg24", "bias_latmo_cfg24", "nse_melt_cfg24", "r_melt_cfg24", "bias_melt_cfg24", "nse_rnet_cfg24", "r_rnet_cfg24", "bias_rnet_cfg24", "nse_rof_cfg24", "r_rof_cfg24", "bias_rof_cfg24", "nse_snowdepth_cfg24", "r_snowdepth_cfg24", "bias_snowdepth_cfg24", "nse_swe_cfg24", "r_swe_cfg24", "bias_swe_cfg24", "nse_gsurf_cfg25", "r_gsurf_cfg25", "bias_gsurf_cfg25", "nse_hatmo_cfg25", "r_hatmo_cfg25", "bias_hatmo_cfg25", "nse_latmo_cfg25", "r_latmo_cfg25", "bias_latmo_cfg25", "nse_melt_cfg25", "r_melt_cfg25", "bias_melt_cfg25", "nse_rnet_cfg25", "r_rnet_cfg25", "bias_rnet_cfg25", "nse_rof_cfg25", "r_rof_cfg25", _
        "bias_rof_cfg25", "nse_snowdepth_cfg25", "r_snowdepth_cfg25", "bias_snowdepth_cfg25", "nse_swe_cfg25", "r_swe_cfg25", "bias_swe_cfg25", "nse_gsurf_cfg26", "r_gsurf_cfg26", "bias_gsurf_cfg26", "nse_hatmo_cfg26", "r_hatmo_cfg26", "bias_hatmo_cfg26", "nse_latmo_cfg26", "r_latmo_cfg26", "bias_latmo_cfg26", "nse_melt_cfg26", "r_melt_cfg26", "bias_melt_cfg26", "nse_rnet_cfg26", "r_rnet_cfg26", "bias_rnet_cfg26", "nse_rof_cfg26", "r_rof_cfg26", "bias_rof_cfg26", "nse_snowdepth_cfg26", "r_snowdepth_cfg26", "bias_snowdepth_cfg26", "nse_swe_cfg26", "r_swe_cfg26", "bias_swe_cfg26", "nse_gsurf_cfg27", "r_gsurf_cfg27", "bias_gsurf_cfg27", "nse_hatmo_cfg27", _
        "r_hatmo_cfg27", "bias_hatmo_cfg27", "nse_latmo_cfg27", "r_latmo_cfg27", "bias_latmo_cfg27", "nse_melt_cfg27", "r_melt_cfg27", "bias_melt_cfg27", "nse_rnet_cfg27", "r_rnet_cfg27", "bias_rnet_cfg27", "nse_rof_cfg27", "r_rof_cfg27", "bias_rof_cfg27", "nse_snowdepth_cfg27", "r_snowdepth_cfg27", "bias_snowdepth_cfg27", "nse_swe_cfg27", "r_swe_cfg27", "bias_swe_cfg27", "nse_gsurf_cfg28", "r_gsurf_cfg28", "bias_gsurf_cfg28", "nse_hatmo_cfg28", "r_hatmo_cfg28", "bias_hatmo_cfg28", "nse_latmo_cfg28", "r_latmo_cfg28", "bias_latmo_cfg28", "nse_melt_cfg28", "r_melt_cfg28", "bias_melt_cfg28", "nse_rnet_cfg28", "r_rnet_cfg28", "bias_rnet_cfg28", _
        "nse_rof_cfg28", "r_rof_cfg28", "bias_rof_cfg28", "nse_snowdepth_cfg28", "r_snowdepth_cfg28", "bias_snowdepth_cfg28", "nse_swe_cfg28", "r_swe_cfg28", "bias_swe_cfg28", "nse_gsurf_cfg29", "r_gsurf_cfg29", "bias_gsurf_cfg29", "nse_hatmo_cfg29", "r_hatmo_cfg29", "bias_hatmo_cfg29", "nse_latmo_cfg29", "r_latmo_cfg29", "bias_latmo_cfg29", "nse_melt_cfg29", "r_melt_cfg29", "bias_melt_cfg29", "nse_rnet_cfg29", "r_rnet_cfg29", "bias_rnet_cfg29", "nse_rof_cfg29", "r_rof_cfg29", "bias_rof_cfg29", "nse_snowdepth_cfg29", "r_snowdepth_cfg29", "bias_snowdepth_cfg29", "nse_swe_cfg29", "r_swe_cfg29", "bias_swe_cfg29", "nse_gsurf_cfg30", "r_gsurf_cfg30", _
        "bias_gsurf_cfg30", "nse_hatmo_cfg30", "r_hatmo_cfg30", "bias_hatmo_cfg30", "nse_latmo_cfg30", "r_latmo_cfg30", "bias_latmo_cfg30", "nse_melt_cfg30", "r_melt_cfg30", "bias_melt_cfg30", "nse_rnet_cfg30", "r_rnet_cfg30", "bias_rnet_cfg30", "nse_rof_cfg30", "r_rof_cfg30", "bias_rof_cfg30", "nse_snowdepth_cfg30", "r_snowdepth_cfg30", "bias_snowdepth_cfg30", "nse_swe_cfg30", "r_swe_cfg30", "bias_swe_cfg30", "nse_gsurf_cfg31", "r_gsurf_cfg31", "bias_gsurf_cfg31", "nse_hatmo_cfg31", "r_hatmo_cfg31", "bias_hatmo_cfg31", "nse_latmo_cfg31", "r_latmo_cfg31", "bias_latmo_cfg31", "nse_melt_cfg31", "r_melt_cfg31", "bias_melt_cfg31", "nse_rnet_cfg31", _
        "r_rnet_cfg31", "bias_rnet_cfg31", "nse_rof_cfg31", "r_rof_cfg31", "bias_rof_cfg31", "nse_snowdepth_cfg31", "r_snowdepth_cfg31", "bias_snowdepth_cfg31", "nse_swe_cfg31", "r_swe_cfg31", "bias_swe_cfg31", "nse_gsurf_cfg32", "r_gsurf_cfg32", "bias_gsurf_cfg32", "nse_hatmo_cfg32", "r_hatmo_cfg32", "bias_hatmo_cfg32", "nse_latmo_cfg32", "r_latmo_cfg32", "bias_latmo_cfg32", "nse_melt_cfg32", "r_melt_cfg32", "bias_melt_cfg32", "nse_rnet_cfg32", "r_rnet_cfg32", "bias_rnet_cfg32", "nse_rof_cfg32", "r_rof_cfg32", "bias_rof_cfg32", "nse_snowdepth_cfg32", "r_snowdepth_cfg32", "bias_snowdepth_cfg32", "nse_swe_cfg32", "r_swe_cfg32", "bias_swe_cfg32"))

    TestRes = TestCSVRead(91, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        NumRows:=1, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test91 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test92(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim k As Long
    Dim m As Long
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler

    TestDescription = "pandas zeros"
    FileName = "pandas_zeros.csv"
    Expected = Empty
    Observed = Empty
    TestRes = TestCSVRead(92, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:="N", _
        ShowMissingsAs:=Empty, _
        NumRowsExpected:=100001, _
        NumColsExpected:=50)
    If TestRes Then
        Dim Total As Double
        For k = 1 To 50
            For m = 1 To 100001
                Total = Total + Observed(m, k)
            Next
        Next
        If Total <> 2499772 Then
            TestRes = False
            WhatDiffers = "Test 92 pandas zeros FAILED, Test was that sum of elements be 2,499,772, but instead its " + Format$(Total, "###,###")
        End If
    End If
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test92 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
Private Sub Test93(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "heat flux.dat"
    FileName = "heat_flux.dat"
    Expected = HStack( _
        Array("#t", 0#, 0.05), _
        Array("heat_flux", 1.14914917397E-07, 1.14914917397E-07))

    TestRes = TestCSVRead(93, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        Delimiter:=" ", _
        IgnoreRepeated:=True, _
        NumRows:=3, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test93 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test94(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    'UTF-8 BOM, and streamed
    TestDescription = "fecal samples"
    FileName = "fecal_samples.csv"
    Expected = HStack( _
        Array("SampleID", "C0052_5F_1A"), Array("Mother_Child", "C"), _
        Array("SubjectID", 52#), Array("MaternalID", "0052_m"), _
        Array("TimePoint", 5#), Array("Fecal_EtOH", "F"), _
        Array("CollectionRep", 1#), Array("DOC", CDate("2017-Jul-25")), _
        Array("RAInitials_DOC", Empty), Array("DOF", CDate("2017-Jul-25")), _
        Array("RAInitials_DOF", Empty), Array("Date_Brought_In", Empty), _
        Array("RAInitials_Brought", Empty), Array("Date_Shipped", CDate("2017-Aug-17")), _
        Array("RAInitials_Shipped", "SR"), Array("Date_Aliquoted", CDate("2017-Sep-08")), _
        Array("Number_Replicates", "A,B,C,D"), Array("RAInitials_Aliquot", "SR"), _
        Array("StorageBox", "Box 1"), Array("DOE", CDate("2017-Nov-21")), _
        Array("Extract_number", "5 of 1"), Array("AliquotRep", "A"), _
        Array("DNABox", "Box 1"), Array("KitUsed", "RNeasy PowerMicrobiome"), _
        Array("RAInitials_Extract", "SR"), Array("DNAConc", 15#), _
        Array("DOM", CDate("2018-Feb-20")), Array("Mgx_processed", "Sequenced"), _
        Array("Mgx_batch", "Batch 1"), Array("DO16S", CDate("2018-Jun-13")), _
        Array("16S_processed", "Sequenced"), Array("16S_batch", "Batch 1"), _
        Array("16S_plate", "Plate 4"), Array("Notes", "DOC recorded as DOF"), _
        Array("Discrepancies", "Written as 52E but needs to be changed to 52F"), _
        Array("Batch 1 Mapping", "052E.7-25-17.F"), Array("Mgx_batch Mapping", "Mgx_batch001"), _
        Array("16S_batch Mapping", "16S_batch001"), Array("Mother/Child Dyads", Empty))

    TestRes = TestCSVRead(94, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        DateFormat:="Y-M-D", _
        NumRows:=2, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test94 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test95(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test d-m-y with time"
    FileName = "test_d-m-y_with_time.csv"
    Expected = HStack( _
        Array(CDate("2021-Sep-01 16:23:13"), CDate("2022-Oct-09 04:16:13"), CDate("2022-Dec-27 13:56:15"), CDate("2022-May-07 08:56:31"), CDate("2024-Jan-14 05:29:48"), _
        CDate("2023-Jan-16 08:12:25"), CDate("2023-Dec-10 13:35:13"), CDate("2023-Jan-11 20:59:27"), CDate("2021-Oct-28 07:31:59"), CDate("2023-Jul-21 00:02:45"), CDate("2021-Dec-16 19:15:38")))

    TestRes = TestCSVRead(95, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:="D", _
        Delimiter:=",", _
        DateFormat:="D-M-Y", _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test95 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test96(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test m-d-y with time"
    FileName = "test_m-d-y_with_time.csv"
    Expected = HStack( _
        Array(CDate("2021-Sep-01 16:23:13"), CDate("2022-Oct-09 04:16:13"), CDate("2022-Dec-27 13:56:15"), CDate("2022-May-07 08:56:31"), CDate("2024-Jan-14 05:29:48"), _
        CDate("2023-Jan-16 08:12:25"), CDate("2023-Dec-10 13:35:13"), CDate("2023-Jan-11 20:59:27"), CDate("2021-Oct-28 07:31:59"), CDate("2023-Jul-21 00:02:45"), CDate("2021-Dec-16 19:15:38")))

    TestRes = TestCSVRead(96, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:="D", _
        Delimiter:=",", _
        DateFormat:="M-D-Y", _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test96 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test97(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test y-m-d with time"
    FileName = "test_y-m-d_with_time.csv"
    Expected = HStack( _
        Array(CDate("2021-Sep-01 16:23:13"), CDate("2022-Oct-09 04:16:13"), CDate("2022-Dec-27 13:56:15"), CDate("2022-May-07 08:56:31"), CDate("2024-Jan-14 05:29:48"), _
        CDate("2023-Jan-16 08:12:25"), CDate("2023-Dec-10 13:35:13"), CDate("2023-Jan-11 20:59:27"), CDate("2021-Oct-28 07:31:59"), CDate("2023-Jul-21 00:02:45"), CDate("2021-Dec-16 19:15:38")))

    TestRes = TestCSVRead(97, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:="D", _
        Delimiter:=",", _
        DateFormat:="Y-M-D", _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test97 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test98(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "attenu"
    FileName = "attenu.csv"
    Expected = HStack( _
        Array("Event", 1#, 2#, 2#, 2#, 2#, 2#, 2#, 2#, 2#), _
        Array("Mag", 7#, 7.4, 7.4, 7.4, 7.4, 7.4, 7.4, 7.4, 7.4), _
        Array("Station", "117", "1083", "1095", "283", "135", "475", "113", "1008", "1028"), _
        Array("Dist", 12#, 148#, 42#, 85#, 107#, 109#, 156#, 224#, 293#), _
        Array("Accel", 0.359, 0.014, 0.196, 0.135, 0.062, 0.054, 0.014, 0.018, 0.01))

    TestRes = TestCSVRead(98, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:="N", _
        NumRows:=10, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test98 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test99(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    'We test that the first column converts (via CSVRead) to the same date as the third column (via CDate) _
     to within a very small (10 microsecond) tolerance to cope with floating point inaccuracies
    TestDescription = "test good ISO8601 with DateFormat = ISO"
    FileName = "test_good_ISO8601.csv"
    Expected = ThrowIfError(CSVRead(Folder + FileName, ConvertTypes:="N", SkipToRow:=2, NumCols:=1, SkipToCol:=3))
    CastDoublesToDates Expected
    TestRes = TestCSVRead(99, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        Delimiter:=",", _
        DateFormat:="ISO", _
        SkipToRow:=2, _
        NumCols:=1, _
        ShowMissingsAs:=Empty, _
        AbsTol:=0.01 / 24 / 60 / 60 / 1000) '10 microsecond tolerance
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test99 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test100(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    'We test that the first column converts (via CSVRead) to the same date as the fourth column (via CDate) _
     to within a very small (10 microsecond) tolerance to cope with floating point inaccuracies
    TestDescription = "test good ISO8601 with DateFormat = ISOZ"
    FileName = "test_good_ISO8601.csv"
    Expected = ThrowIfError(CSVRead(Folder + FileName, ConvertTypes:="N", SkipToRow:=2, NumCols:=1, SkipToCol:=4))
    CastDoublesToDates Expected
    TestRes = TestCSVRead(100, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        Delimiter:=",", _
        DateFormat:="ISOZ", _
        SkipToRow:=2, _
        NumCols:=1, _
        ShowMissingsAs:=Empty, _
        AbsTol:=0.01 / 24 / 60 / 60 / 1000) '10 microsecond tolerance
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test100 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test101(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    'Test that parsing strings that almost but not correct ISO8601 does not convert to dates
    TestDescription = "test bad ISO8601"
    FileName = "test_bad_ISO8601.csv"
    Expected = ThrowIfError(CSVRead(Folder + FileName, False, ",", SkipToRow:=2, SkipToCol:=2))
    TestRes = TestCSVRead(101, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:="D", _
        Delimiter:=",", _
        DateFormat:="ISO", _
        SkipToRow:=2, _
        SkipToCol:=2, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test101 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test102(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    'We test that the first column converts (via CSVRead) to the same date as the second column (via CDate) _
     to within a very small (10 microsecond) tolerance to cope with floating point inaccuracies
    TestDescription = "test good Y-M-D"
    FileName = "test_good_Y-M-D.csv"
    Expected = ThrowIfError(CSVRead(Folder + FileName, ConvertTypes:="N", SkipToRow:=2, NumCols:=1, SkipToCol:=2))
    CastDoublesToDates Expected
    TestRes = TestCSVRead(102, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        Delimiter:=",", _
        DateFormat:="Y-M-D", _
        SkipToRow:=2, _
        NumCols:=1, _
        ShowMissingsAs:=Empty, _
        AbsTol:=0.01 / 24 / 60 / 60 / 1000) '10 microsecond tolerance
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test102 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test103(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    'Test that parsing strings that almost but not correct Y-M-D does not convert to dates
    TestDescription = "test bad Y-M-D"
    FileName = "test_bad_Y-M-D.csv"
    Expected = ThrowIfError(CSVRead(Folder + FileName, False, ",", SkipToRow:=2, SkipToCol:=2))
    TestRes = TestCSVRead(103, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:="D", _
        Delimiter:=",", _
        DateFormat:="Y-M-D", _
        SkipToRow:=2, _
        SkipToCol:=2, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test103 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test104(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    'We test that the first column converts (via CSVRead) to the same date as the second column (via CDate) _
     to within a very small (10 microsecond) tolerance to cope with floating point inaccuracies
    TestDescription = "test good D-M-Y"
    FileName = "test_good_D-M-Y.csv"
    Expected = ThrowIfError(CSVRead(Folder + FileName, ConvertTypes:="N", SkipToRow:=2, NumCols:=1, SkipToCol:=2))
    CastDoublesToDates Expected
    TestRes = TestCSVRead(104, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        Delimiter:=",", _
        DateFormat:="D-M-Y", _
        SkipToRow:=2, _
        NumCols:=1, _
        ShowMissingsAs:=Empty, _
        AbsTol:=0.01 / 24 / 60 / 60 / 1000) '10 microsecond tolerance
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test104 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test105(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    'Test that parsing strings that almost but not correct D-M-Y does not convert to dates
    TestDescription = "test bad D-M-Y"
    FileName = "test_bad_D-M-Y.csv"
    Expected = ThrowIfError(CSVRead(Folder + FileName, False, ",", SkipToRow:=2, SkipToCol:=2))
    TestRes = TestCSVRead(105, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:="D", _
        Delimiter:=",", _
        DateFormat:="D-M-Y", _
        SkipToRow:=2, _
        SkipToCol:=2, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test105 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test106(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    'We test that the first column converts (via CSVRead) to the same date as the second column (via CDate) _
     to within a very small (10 microsecond) tolerance to cope with floating point inaccuracies
    TestDescription = "test good M-D-Y"
    FileName = "test_good_M-D-Y.csv"
    Expected = ThrowIfError(CSVRead(Folder + FileName, ConvertTypes:="N", SkipToRow:=2, NumCols:=1, SkipToCol:=2))
    CastDoublesToDates Expected
    TestRes = TestCSVRead(106, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        Delimiter:=",", _
        DateFormat:="M-D-Y", _
        SkipToRow:=2, _
        NumCols:=1, _
        ShowMissingsAs:=Empty, _
        AbsTol:=0.01 / 24 / 60 / 60 / 1000) '10 microsecond tolerance
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test106 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test107(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    'Test that parsing strings that almost but not correct M-D-Y does not convert to dates
    TestDescription = "test bad M-D-Y"
    FileName = "test_bad_M-D-Y.csv"
    Expected = ThrowIfError(CSVRead(Folder + FileName, False, ",", SkipToRow:=2, SkipToCol:=2))
    TestRes = TestCSVRead(107, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:="D", _
        Delimiter:=",", _
        DateFormat:="M-D-Y", _
        SkipToRow:=2, _
        SkipToCol:=2, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test107 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test108(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "ampm"
    FileName = "ampm.csv"
    Expected = HStack( _
        Array("ID", 4473#, 3513#), _
        Array("INTERLOCK_NUMBER", Empty, "REDACTED"), _
        Array("INTERLOCK_DESCRIPTION", "This is a test interlock", "<p>REDACTED:<br><br>REDACTED</p><p><br></p>"), _
        Array("TYPE", Empty, Empty), _
        Array("CREATE_DATE", CDate("2012-Feb-09"), CDate("1998-Jul-22 16:37:01")), _
        Array("MODIFY_DATE", CDate("2018-Nov-02"), CDate("2019-Jun-20")), _
        Array("USERNAME", "REDACTED   ", "REDACTED       "), _
        Array("UNIT", "U         ", "A         "), _
        Array("AREA", "U         ", "RM        "), _
        Array("PURPOSE", "This is a test interlock", " REDACTED"), _
        Array("PID", Empty, Empty), _
        Array("LOCATION", Empty, Empty), _
        Array("FUNC_DATE", CDate("1900-Jan-01"), CDate("1900-Jan-01")), _
        Array("FUNC_BY", Empty, Empty), _
        Array("TECHNICAL_DESCRIPTION", "This is a test interlock", "<p>REDACTED<br><br>REDACTED<br>REDACTED</p>"), _
        Array("types", "O", "SC"))

    TestRes = TestCSVRead(108, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        DateFormat:="M/D/Y", _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test108 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test109(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "time"
    FileName = "time.csv"
    Expected = HStack(Array("time", CDate("00:00:00"), CDate("00:10:00")), Array("value", 1#, 2#))
    TestRes = TestCSVRead(109, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test109 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test110(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test datetimes"
    'CDate can't cope with fractions of a second, so adjust via 0.001/86400 term
    Expected = HStack(Array("col1", CDate("2015-Jan-01"), CDate("2015-Jan-02 00:00:01"), CDate("2015-Jan-03 00:12:00") + 0.001 / 86400))
    FileName = "test_datetimes.csv"
    TestRes = TestCSVRead(110, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        Delimiter:=",", _
        DateFormat:="Y-M-D", _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test110 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test111(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "dash as null"
    FileName = "dash_as_null.csv"
    Expected = HStack(Array("x", 1#, Empty), Array("y", 2#, 4#))
    TestRes = TestCSVRead(111, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        MissingStrings:="-", _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test111 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test112(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    'Different from Julia equivalent in that elements of first column have different type whereas Julia parses col 1 to be all strings
    TestDescription = "test null only column"
    FileName = "test_null_only_column.csv"
    Expected = HStack(Array("col1", 123#, "abc", "123abc"), Array("col2", Empty, Empty, Empty))
    TestRes = TestCSVRead(112, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        MissingStrings:="NA", _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test112 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test113(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test one row of data"
    FileName = "test_one_row_of_data.csv"
    Expected = HStack(1#, 2#, 3#)
    TestRes = TestCSVRead(113, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test113 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test114(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "plus as null"
    FileName = "plus_as_null.csv"
    Expected = HStack(Array("x", 1#, Empty), Array("y", CDate("1900-Jan-01"), CDate("1900-Jan-03")))
    TestRes = TestCSVRead(114, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        MissingStrings:="+", _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test114 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test115(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "categorical"
    FileName = "categorical.csv"
    Expected = HStack(Array("cat", "a", "a", "a", "b", "b", "b", "b", "b", "b", "b", "c", "c", "c", "c", "a"))
    TestRes = TestCSVRead(115, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test115 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test116(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test file issue 154"
    FileName = "test_file_issue_154.csv"
    Expected = HStack( _
        Array("a", 0#, 12#), _
        Array(" b", 1#, 5#), _
        Array(" ", " ", " "), _
        Array(Empty, " comment ", Empty))

    TestRes = TestCSVRead(116, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test116 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test117(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test int sentinel"
    FileName = "test_int_sentinel.csv"
    Expected = HStack( _
        Array("id", 1#, 2#, 3#, 4#, 5#, 6#, 7#, 8#, 9#, 10#, 11#, 12#, 13#, 14#, 15#, 16#, 17#, 18#, 19#), _
        Array("firstname", "Lawrence", "Benjamin", "Wayne", "Sean", "Charles", "Linda", "Steve", "Jacqueline", "Tammy", "Nicholas", "Irene", "Gary", "David", "Jennifer", "Gary", "Theresa", "Carl", "Judy", "Jane"), _
        Array("lastname", "Powell", "Chavez", "Burke", "Richards", "Long", "Rose", "Gardner", "Roberts", "Reynolds", "Ramos", "King", "Banks", "Knight", "Collins", "Vasquez", "Mason", "Williams", "Howard", "Harris"), _
        Array("salary", 87216.81, 57043.38, 46134.09, 45046.21, 30555.6, 88894.06, 32414.46, 54839.54, 62300.64, 57661.69, 55565.61, 57620.06, 49729.65, 86834#, 47974.45, 67476.24, 71048.06, 53110.54, 52664.59), _
        Array("hourlyrate", 26.47, 39.44, 33.8, 15.64, 17.67, 34.6, 36.39, 26.27, 37.67, 21.37, 13.88, 15.68, 10.39, 10.18, 24.52, 41.47, 29.67, 42.1, 16.48), _
        Array("hiredate", 37355#, 40731#, 42419#, 36854#, 37261#, 39583#, 38797#, Empty, 36686#, 37519#, 38821#, Empty, 37489#, 39239#, 40336#, 36794#, 39764#, 42123#, 38319#), _
        Array("lastclockin", CDate("2002-Jan-17 21:32:00"), CDate("2000-Sep-25 06:36:00"), CDate("2002-Sep-13 08:28:00"), CDate("2011-Jul-10 11:24:00"), CDate("2003-Feb-11 11:43:00"), CDate("2016-Jan-21 06:32:00"), CDate("2004-Jan-12 12:36:00"), Empty, CDate("2006-Dec-30 09:48:00"), CDate("2016-Apr-07 14:07:00"), CDate("2015-Mar-19 15:01:00"), Empty, CDate("2005-Jun-29 11:14:00"), CDate("2001-Sep-17 11:47:00"), CDate("2014-Aug-30 02:41:00"), CDate("2015-Nov-07 01:23:00"), CDate("2009-Sep-06 20:21:00"), CDate("2011-May-14 14:38:00"), CDate("2000-Oct-17 14:18:00")))

    TestRes = TestCSVRead(117, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        DateFormat:="ISO", _
        NumRows:=20, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test117 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test118(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String
    'The is test is fragile in that line endings in text files can get flipped from vbLf to vbCrLf as files are pushed and pulled to git.
    On Error GoTo ErrHandler
    TestDescription = "escape row starts"
    Expected = HStack( _
        Array("5111", "escaped row with " + vbCrLf + " newlines " + vbCrLf + "  " + vbCrLf + "  " + vbCrLf + "  in it", "5113"), _
        Array("5112", "5113", "5114"))
    FileName = "escape_row_starts.csv"
    TestRes = TestCSVRead(118, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        SkipToRow:=5112, _
        NumRows:=3, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test118 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test119(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "Sacramentorealestatetransactions"
    FileName = "Sacramentorealestatetransactions.csv"
    Expected = Empty
    TestRes = TestCSVRead(119, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        ShowMissingsAs:=Empty, _
        NumRowsExpected:=986, _
        NumColsExpected:=12)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test119 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test120(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "log001 vehicle status flags 0.txt"
    FileName = "log001_vehicle_status_flags_0.txt"
    TestRes = TestCSVRead(120, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        ShowMissingsAs:=Empty, _
        NumRowsExpected:=282, _
        NumColsExpected:=31)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test120 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test121(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "SalesJan2009"
    FileName = "SalesJan2009.csv"
    Expected = HStack( _
        Array("Transaction_date", CDate("2009-Jan-02 06:17:00"), CDate("2009-Jan-02 04:53:00"), CDate("2009-Jan-02 13:08:00"), CDate("2009-Jan-03 14:44:00"), CDate("2009-Jan-04 12:56:00"), CDate("2009-Jan-04 13:19:00"), CDate("2009-Jan-04 20:11:00"), CDate("2009-Jan-02 20:09:00"), CDate("2009-Jan-04 13:17:00"), CDate("2009-Jan-04 14:11:00"), CDate("2009-Jan-05 02:42:00"), CDate("2009-Jan-05 05:39:00"), CDate("2009-Jan-02 09:16:00"), CDate("2009-Jan-05 10:08:00"), CDate("2009-Jan-02 14:18:00"), CDate("2009-Jan-04 01:05:00"), CDate("2009-Jan-05 11:37:00"), CDate("2009-Jan-06 05:02:00"), CDate("2009-Jan-06 07:45:00")), _
        Array("Product", "Product1", "Product1", "Product1", "Product1", "Product2", "Product1", "Product1", "Product1", "Product1", "Product1", "Product1", "Product1", "Product1", "Product1", "Product1", "Product1", "Product1", "Product1", "Product2"), _
        Array("Price", 1200#, 1200#, 1200#, 1200#, 3600#, 1200#, 1200#, 1200#, 1200#, 1200#, 1200#, 1200#, 1200#, 1200#, 1200#, 1200#, 1200#, 1200#, 3600#), _
        Array("Payment_Type", "Mastercard", "Visa", "Mastercard", "Visa", "Visa", "Visa", "Mastercard", "Mastercard", "Mastercard", "Visa", "Diners", "Amex", "Mastercard", "Visa", "Visa", "Diners", "Visa", "Diners", "Visa"), _
        Array("Name", "carolina", "Betina", "Federica e Andrea", "Gouya", "Gerd W ", "LAURENCE", "Fleur", "adam", "Renee Elisabeth", "Aidan", "Stacy", "Heidi", "Sean ", "Georgia", "Richard", "Leanne", "Janet", "barbara", "Sabine"), _
        Array("City", "Basildon", "Parkville                   ", "Astoria                     ", "Echuca", "Cahaba Heights              ", "Mickleton                   ", "Peoria                      ", "Martin                      ", "Tel Aviv", "Chatou", "New York                    ", "Eindhoven", "Shavano Park                ", "Eagle                       ", "Riverside                   ", "Julianstown", "Ottawa", "Hyderabad", "London"), _
        Array("State", "England", "MO", "OR", "Victoria", "AL", "NJ", "IL", "TN", "Tel Aviv", "Ile-de-France", "NY", "Noord-Brabant", "TX", "ID", "NJ", "Meath", "Ontario", "Andhra Pradesh", "England"), _
        Array("Country", "United Kingdom", "United States", "United States", "Australia", "United States", "United States", "United States", "United States", "Israel", "France", "United States", "Netherlands", "United States", "United States", "United States", "Ireland", "Canada", "India", "United Kingdom"), _
        Array("Account_Created", CDate("2009-Jan-02 06:00:00"), CDate("2009-Jan-02 04:42:00"), CDate("2009-Jan-01 16:21:00"), CDate("2005-Sep-25 21:13:00"), CDate("2008-Nov-15 15:47:00"), CDate("2008-Sep-24 15:19:00"), CDate("2009-Jan-03 09:38:00"), CDate("2009-Jan-02 17:43:00"), CDate("2009-Jan-04 13:03:00"), CDate("2008-Jun-03 04:22:00"), CDate("2009-Jan-05 02:23:00"), CDate("2009-Jan-05 04:55:00"), CDate("2009-Jan-02 08:32:00"), CDate("2008-Nov-11 15:53:00"), CDate("2008-Dec-09 12:07:00"), CDate("2009-Jan-04"), CDate("2009-Jan-05 09:35:00"), CDate("2009-Jan-06 02:41:00"), CDate("2009-Jan-06 07:00:00")), _
        Array("Last_Login", CDate("2009-Jan-02 06:08:00"), CDate("2009-Jan-02 07:49:00"), CDate("2009-Jan-03 12:32:00"), CDate("2009-Jan-03 14:22:00"), CDate("2009-Jan-04 12:45:00"), CDate("2009-Jan-04 13:04:00"), CDate("2009-Jan-04 19:45:00"), CDate("2009-Jan-04 20:01:00"), CDate("2009-Jan-04 22:10:00"), CDate("2009-Jan-05 01:17:00"), CDate("2009-Jan-05 04:59:00"), CDate("2009-Jan-05 08:15:00"), CDate("2009-Jan-05 09:05:00"), CDate("2009-Jan-05 10:05:00"), CDate("2009-Jan-05 11:01:00"), CDate("2009-Jan-05 13:36:00"), CDate("2009-Jan-05 19:24:00"), CDate("2009-Jan-06 07:52:00"), CDate("2009-Jan-06 09:17:00")), _
        Array("Latitude", 51.5, 39.195, 46.18806, -36.1333333, 33.52056, 39.79, 40.69361, 36.34333, 32.0666667, 48.8833333, 40.71417, 51.45, 29.42389, 43.69556, 40.03222, 53.6772222, 45.4166667, 17.3833333, 51.52721), _
        Array("Longitude", -1.1166667, -94.68194, -123.83, 144.75, -86.8025, -75.23806, -89.58889, -88.85028, 34.7666667, 2.15, -74.00639, 5.4666667, -98.49333, -116.35306, -74.95778, -6.3191667, -75.7, 78.4666667, 0.14559))

    TestRes = TestCSVRead(121, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        DateFormat:="M/D/Y", _
        NumRows:=20, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test121 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test122(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "GSM2230757 human1 umifm counts"
    FileName = "GSM2230757_human1_umifm_counts.csv"
    Expected = Empty
    TestRes = TestCSVRead(122, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        ShowMissingsAs:=Empty, _
        NumRowsExpected:=4, _
        NumColsExpected:=20128)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test122 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test123(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "SacramentocrimeJanuary2006"
    FileName = "SacramentocrimeJanuary2006.csv"
    Expected = HStack( _
        Array(CDate("2006-Jan-31 23:31:00"), CDate("2006-Jan-31 23:36:00"), CDate("2006-Jan-31 23:40:00"), CDate("2006-Jan-31 23:41:00"), CDate("2006-Jan-31 23:45:00"), CDate("2006-Jan-31 23:50:00")), _
        Array("39TH ST / STOCKTON BLVD", "26TH ST / G ST", "4011 FREEPORT BLVD", "30TH ST / K ST", "5303 FRANKLIN BLVD", "COBBLE COVE LN / COBBLE SHORES DR"), _
        Array(6#, 3#, 4#, 3#, 4#, 4#), _
        Array("6B        ", "3B        ", "4A        ", "3C        ", "4B        ", "4C        "), _
        Array(1005#, 728#, 957#, 841#, 969#, 1294#), _
        Array("CASUALTY REPORT", "594(B)(2)(A) VANDALISM/ -$400", "459 PC  BURGLARY BUSINESS", "TRAFFIC-ACCIDENT INJURY", "3056 PAROLE VIO - I RPT", "TRAFFIC-ACCIDENT-NON INJURY"), _
        Array(7000#, 2999#, 2203#, 5400#, 7000#, 5400#), _
        Array(38.5566387, 38.57783198, 38.53759051, 38.57203045, 38.52718667, 38.47962803), _
        Array(-121.4597445, -121.4704595, -121.4925914, -121.4670118, -121.4712477, -121.5286345))

    TestRes = TestCSVRead(123, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        DateFormat:="M/D/Y", _
        SkipToRow:=7580, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test123 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test124(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test empty lines"
    Expected = HStack(Array("a", "1", "4", "7"), Array("b", "2", "5", "8"), Array("c", "3", "6", "9"))
    FileName = "test_empty_lines.csv"
    TestRes = TestCSVRead(124, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test124 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test125(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test padding"
    FileName = "test_padding.csv"
    Expected = HStack( _
        Array("col1", 1#, 4#, 7#, Empty), _
        Array("col2", 2#, 5#, 8#, Empty), _
        Array("col3", 3#, 6#, 9#, Empty), _
        Array(Empty, Empty, Empty, Empty, Empty))

    TestRes = TestCSVRead(125, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        NumRows:=5, _
        NumCols:=4, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test125 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test126(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test not delimited"
    FileName = "test_not_delimited.csv"
    Expected = HStack(Array("col1,col2,col3", "1,2,3", "4,5,6", "7,8,9"))
    TestRes = TestCSVRead(126, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        Delimiter:="False", _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test126 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test127(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test string first argument"
    FileName = "col1,col2,col3" & vbLf & "1,2,3" & vbLf & "4,5,6" & vbLf & "7,8,9"
    Expected = HStack( _
        Array("col1", "1", "4", "7"), _
        Array("col2", "2", "5", "8"), _
        Array("col3", "3", "6", "9"))

    TestRes = TestCSVRead(127, TestDescription, Expected, FileName, Observed, WhatDiffers, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test127 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test128(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "Fielding"
    FileName = "Fielding.csv"
    Expected = Empty
    TestRes = TestCSVRead(128, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:="N", _
        ShowMissingsAs:=Empty, _
        NumRowsExpected:=167939, _
        NumColsExpected:=18)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test128 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test129(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "precompile"
    FileName = "precompile.csv"
    Expected = HStack( _
        Array("int", 1#), _
        Array("float", 2#), _
        Array("pool", "a"), _
        Array("string", "RTrBP"), _
        Array("bool", True), _
        Array("date", CDate("2020-Jun-20")), _
        Array("datetime", CDate("2020-Jun-20")), _
        Array("time", CDate("12:00:00")))

    TestRes = TestCSVRead(129, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        DateFormat:="ISO", _
        NumRows:=2, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test129 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test130(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "precompile"
    FileName = "precompile.csv"
    Expected = Empty
    TestRes = TestCSVRead(130, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        DateFormat:="ISO", _
        ShowMissingsAs:=Empty, _
        NumRowsExpected:=5002, _
        NumColsExpected:=8)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test130 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test131(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "big types"
    FileName = "big_types.csv"
    Expected = HStack( _
        Array("time", CDate("12:00:00"), CDate("12:00:00")), _
        Array("bool", True, True), _
        Array("lazy", "hey", "hey"), _
        Array("lazy_missing", Empty, "ho"))

    TestRes = TestCSVRead(131, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        DateFormat:="ISO", _
        NumRows:=3, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test131 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test132(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test headers"
    FileName = "test_headers.csv"
    Expected = HStack(Array(2, 12), Array(3, 13))
    TestRes = TestCSVRead(132, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        DateFormat:="ISO", _
        SkipToRow:=2, _
        SkipToCol:=2, _
        NumRows:=2, _
        NumCols:=2, _
        ShowMissingsAs:=Empty, _
        HeaderRowNum:=1#, _
        ExpectedHeaderRow:=HStack("Col2", "Col3"))
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test132 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test133(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test ragged headers"
    FileName = "test_ragged_headers.csv"
    Expected = HStack( _
        Array(1#, 2#, 4#), _
        Array(Empty, 3#, 5#), _
        Array(Empty, Empty, 6#), _
        Array(Empty, Empty, Empty), _
        Array(Empty, Empty, Empty), _
        Array(Empty, Empty, Empty), _
        Array(Empty, Empty, Empty), _
        Array(Empty, Empty, Empty), _
        Array(Empty, Empty, Empty), _
        Array(Empty, Empty, Empty))

    TestRes = TestCSVRead(133, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        IgnoreEmptyLines:=False, _
        SkipToRow:=2, _
        NumRows:=3, _
        ShowMissingsAs:=Empty, _
        HeaderRowNum:=1#, _
        ExpectedHeaderRow:=HStack("Col1", "Col2", "Col3", "Col4", "Col5", "Col6", "Col7", "Col8", "Col9", "Col10"))
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test133 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test134(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test header on row 4"
    FileName = "test_header_on_row_4.csv"
    Expected = HStack(Array("1", "4", "7"), Array("2", "5", "8"), Array("3", "6", "9"))
    TestRes = TestCSVRead(134, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        IgnoreEmptyLines:=False, _
        SkipToRow:=5, _
        ShowMissingsAs:=Empty, _
        HeaderRowNum:=4#, _
        ExpectedHeaderRow:=HStack("col1", "col2", "col3"))
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test134 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test135(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test bad inputs"
    FileName = "test_bad_inputs.csv"
    Expected = "#CSVRead: Delimiter character must be passed as a string, FALSE for no delimiter. Omit to guess from file contents!"
    TestRes = TestCSVRead(135, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        Delimiter:=1, _
        IgnoreEmptyLines:=False, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test135 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test136(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test bad inputs"
    FileName = "test_bad_inputs.csv"
    Expected = "#CSVRead: Delimiter must have at least one character and cannot start with a double quote, line feed or carriage return!"
    TestRes = TestCSVRead(136, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        Delimiter:="""bad", _
        IgnoreEmptyLines:=False, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test136 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test137(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test bad inputs"
    FileName = "test_bad_inputs.csv"
    Expected = "#CSVRead: NumCols must be positive to read a given number of columns, or zero or omitted to read all columns from SkipToCol to the maximum column encountered.!"
    TestRes = TestCSVRead(137, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        IgnoreEmptyLines:=False, _
        NumCols:=-1, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test137 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test138(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test bad inputs"
    FileName = "test_bad_inputs.csv"
    Expected = "#CSVRead: NumRows must be positive to read a given number of rows, or zero or omitted to read all rows from SkipToRow to the end of the file.!"
    TestRes = TestCSVRead(138, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        IgnoreEmptyLines:=False, _
        NumRows:=-1, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test138 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test139(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test bad inputs"
    FileName = "test_bad_inputs.csv"
    Expected = "#CSVRead: DecimalSeparator must be a single character!"
    TestRes = TestCSVRead(139, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        IgnoreEmptyLines:=False, _
        ShowMissingsAs:=Empty, _
        DecimalSeparator:="bad")
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test139 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test140(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test bad inputs"
    FileName = "test_bad_inputs.csv"
    Expected = "#CSVRead: DecimalSeparator must not be equal to the first character of Delimiter or to a line-feed or carriage-return!"
    TestRes = TestCSVRead(140, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:="N", _
        Delimiter:=",", _
        IgnoreEmptyLines:=False, _
        ShowMissingsAs:=Empty, _
        DecimalSeparator:=",")
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test140 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test141(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test bad inputs"
    FileName = "test_bad_inputs.csv"
    Expected = "#CSVRead: SkipToCol must be at least 1.!"
    TestRes = TestCSVRead(141, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        IgnoreEmptyLines:=False, _
        SkipToCol:=-1, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test141 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test142(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test bad inputs"
    FileName = "test_bad_inputs.csv"
    Expected = "#CSVRead: HeaderRowNum must be greater than or equal to zero and less than or equal to SkipToRow!"
    TestRes = TestCSVRead(142, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        IgnoreEmptyLines:=False, _
        SkipToRow:=-1, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test142 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test143(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test bad inputs"
    FileName = "test_bad_inputs.csv"
    Expected = "#CSVRead: Comment must not contain double-quote, line feed or carriage return!"
    TestRes = TestCSVRead(143, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        Comment:="bad""", _
        IgnoreEmptyLines:=False, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test143 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test144(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test bad inputs"
    FileName = "test_bad_inputs.csv"
    Expected = "#CSVRead: HeaderRowNum must be greater than or equal to zero and less than or equal to SkipToRow!"
    TestRes = TestCSVRead(144, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        IgnoreEmptyLines:=False, _
        ShowMissingsAs:=Empty, _
        HeaderRowNum:=-1#)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test144 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test145(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test bad inputs"
    Expected = "#CSVRead: #ParseEncoding: Encoding argument can usually be omitted, but otherwise Encoding be either ""ASCII"", ""ANSI"", ""UTF-8"", ""UTF-8-BOM"", ""UTF-16"" or ""UTF-16-BOM""!!"
    FileName = "test_bad_inputs.csv"
    TestRes = TestCSVRead(145, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        IgnoreEmptyLines:=False, _
        ShowMissingsAs:=Empty, _
        Encoding:="BAD")
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test145 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test146(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test bad inputs"
    FileName = "test_bad_inputs.csv"
    Expected = "#CSVRead: #ParseConvertTypes: #OneDArrayToTwoDArray: If ConvertTypes is given as a 1-dimensional array, each element must be a 1-dimensional array with two elements!!!"
    TestRes = TestCSVRead(146, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        IgnoreEmptyLines:=False, _
        ShowMissingsAs:=Empty, _
        ConvertTypes:=Array(Array(1, "N", "BAD")))
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test146 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test147(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test bad inputs"
    FileName = "test_bad_inputs.csv"
    Expected = "#CSVRead: #ParseConvertTypes: ConvertTypes is ambiguous, it can be interpreted as two rows, or as two columns!!"
    TestRes = TestCSVRead(147, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=Array(Array(1, "D"), Array("B", "N")), _
        IgnoreEmptyLines:=False, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test147 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test148(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test bad inputs"
    FileName = "test_bad_inputs.csv"
    Expected = "#CSVRead: #ParseConvertTypes: Column identifiers in the left column (or top row) of ConvertTypes must be strings or non-negative whole numbers but ConvertTypes(1,1) is of type Boolean!!"
    TestRes = TestCSVRead(148, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=Array(Array(True, "D"), Array("B", "N"), Array(1, "D")), _
        IgnoreEmptyLines:=False, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test148 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test149(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test bad inputs"
    FileName = "test_bad_inputs.csv"
    Expected = "#CSVRead: #ParseConvertTypes: #ParseCTString: ConvertTypes is incorrect, ""Q"" indicates that conversion should apply even to quoted fields, but none of ""N"", ""D"", ""B"" or ""E"" are present to indicate which type conversion to apply!!!"
    TestRes = TestCSVRead(149, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:="Q", _
        IgnoreEmptyLines:=False, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test149 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test150(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test bad inputs"
    FileName = "test_bad_inputs.csv"
    Expected = "#CSVRead: Delimiter must have at least one character and cannot start with a double quote, line feed or carriage return!"
    TestRes = TestCSVRead(150, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        Delimiter:="""", _
        IgnoreEmptyLines:=False, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test150 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test151(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test bad inputs"
    FileName = "test_bad_inputs.csv"
    Expected = "#CSVRead: #MakeSentinels: #AddKeysToDict: #AddKeyToDict: TrueStrings must be omitted or provided as string or an array of strings that represent Boolean value True but '2' is of type Double!!!!"
    TestRes = TestCSVRead(151, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:="B", _
        IgnoreEmptyLines:=False, _
        TrueStrings:=2#, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test151 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test152(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test bad inputs"
    FileName = "test_bad_inputs.csv"
    Expected = "#CSVRead: #MakeSentinels: TrueStrings has been provided, but type conversion for Booleans is not switched on for any column!!"
    TestRes = TestCSVRead(152, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        IgnoreEmptyLines:=False, _
        TrueStrings:="Bad", _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test152 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test153(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test bad inputs"
    FileName = "test_bad_inputs.csv"
    Expected = "#CSVRead: #MakeSentinels: FalseStrings has been provided, but type conversion for Booleans is not switched on for any column!!"
    TestRes = TestCSVRead(153, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        IgnoreEmptyLines:=False, _
        FalseStrings:="Bad", _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test153 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test154(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test ragged headers, take 2"
    FileName = "test_ragged_headers.csv"
    Expected = HStack(Array(0#, "15"), Array(0#, Empty))
    TestRes = TestCSVRead(154, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        IgnoreEmptyLines:=False, _
        SkipToRow:=5, _
        SkipToCol:=5, _
        NumRows:=2, _
        NumCols:=2, _
        ShowMissingsAs:=Empty, _
        HeaderRowNum:=1#, _
        ExpectedHeaderRow:=HStack("Col5", "Col6"))
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test154 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test155(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test column-by-column"
    FileName = "test_column-by-column.csv"
    Expected = HStack( _
        Array("Type", "Number", "Date", "Boolean", "Error", "String", "String", "String", "String", "String", "String", "String"), _
        Array("Col A", 44424#, CDate("2021-Aug-24 12:49:13"), True, "#DIV/0!", "1", "2021-08-24T12:49:13", "TRUE", "#DIV/0!", "abc", "abc""def", "Line" + vbCrLf + "Feed"), _
        Array("Col B", 44424#, "2021-08-24T12:49:13", "True", "#DIV/0!", "1", "2021-08-24T12:49:13", "TRUE", "#DIV/0!", "abc", "abc""def", "Line" + vbCrLf + "Feed"), _
        Array("Col C", "44424", CDate("2021-Aug-24 12:49:13"), "True", "#DIV/0!", "1", "2021-08-24T12:49:13", "TRUE", "#DIV/0!", "abc", "abc""def", "Line" + vbCrLf + "Feed"), _
        Array("Col D", "44424", "2021-08-24T12:49:13", True, CVErr(2007), "1", "2021-08-24T12:49:13", "TRUE", "#DIV/0!", "abc", "abc""def", "Line" + vbCrLf + "Feed"), _
        Array("Col E", 44424#, "2021-08-24T12:49:13", "True", "#DIV/0!", 1#, "2021-08-24T12:49:13", "TRUE", "#DIV/0!", "abc", "abc""def", "Line" + vbCrLf + "Feed"), _
        Array("Col F", "44424", CDate("2021-Aug-24 12:49:13"), "True", "#DIV/0!", "1", CDate("2021-Aug-24 12:49:13"), "TRUE", "#DIV/0!", "abc", "abc""def", "Line" + vbCrLf + "Feed"), _
        Array("Col G", "44424", "2021-08-24T12:49:13", True, CVErr(2007), "1", "2021-08-24T12:49:13", True, CVErr(2007), "abc", "abc""def", "Line" + vbCrLf + "Feed"))

    TestRes = TestCSVRead(155, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=HStack(Array(0#, "Col B", "Col C", "Col D", "Col E", "Col F", "Col G"), Array(True, "N", "D", "BE", "NQ", "DQ", "BEQ")), _
        DateFormat:="ISO", _
        IgnoreEmptyLines:=False, _
        ShowMissingsAs:=Empty, _
        HeaderRowNum:=1#, _
        SkipToRow:=1#, _
        ExpectedHeaderRow:=HStack("Type", "Col A", "Col B", "Col C", "Col D", "Col E", "Col F", "Col G"))
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test155 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test156(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test not delimited.txt"
    FileName = "test_not_delimited.txt"
    Expected = VStack( _
        VStack("using CSV#= as of 29 Aug 2021 one needs the main version of CSV, latest released version", "0.8.5 fails to load some of the files. See https://github.com/JuliaData/CSV.jl/issues/879", "Use:", "]add CSV#main", "=#", "using DataFrames", vbNullString, "function benchmark()", "    benchmark_csvs_in_folder(""C:/Users/phili/AppData/Local/Temp/VBA-CSV/Performance"", ", "    ""C:/Users/phili/AppData/Local/Temp/VBA-CSV/Performance/JuliaResults.txt"")"), _
        VStack("end", vbNullString, """""""", "   benchmark_csvs_in_folder(foldername::String, outputfile::String)", "Benchmark all .csv files in `foldername`, writing results to `outputfile`.", """""""", "function benchmark_csvs_in_folder(foldername::String, outputfile::String)", "    files = readdir(foldername, join=true)", "    files = filter(x -> x[end - 3:end] == "".csv"", files)# only .csv files", vbNullString), _
        VStack("    n = length(files)", "    times = fill(0.0, n)", "    numcalls = fill(0, n)", "    statuses = fill(""OK"", n)", vbNullString, "    foo = benchmarkonefile(files[1], 1)# for compilation ""warmup""", "    for (f, i) in zip(files, 1:n)", "        println(i, f)", "        try", "            times[i], numcalls[i] = benchmarkonefile(f, 5)"), _
        VStack("        catch e", "            statuses[i] = ""$e""", "        end", "    end", "    times", vbNullString, "    result = DataFrame(filename=replace.(files, ""/"" => ""\\""), time=times, ", "                        status=statuses, numcalls=numcalls)", "    CSV.write(outputfile, result)", "end"), _
        VStack(vbNullString, """""""", "    benchmarkonefile(filename::String, timeout::Int)", "Average time (over sufficient trials to take `timeout` seconds) to load file `filename` to", "a DataFrame, using CSV.File.", """""""", "function benchmarkonefile(filename::String, timeout::Int)", "    i = 0", "    time2 = time() # needed to give time2 scope outside the loop.", "    time1 = time()"), _
        VStack("    while true", "        i = i + 1", "        res = CSV.File(filename, header=false, type=String) |> DataFrame", "        time2 = time()", "        time2 - time1 < timeout || break", "    end", "    (time2 - time1) / i, i", "end"))

    TestRes = TestCSVRead(156, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        Delimiter:=False, _
        IgnoreEmptyLines:=False, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test156 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test157(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test bad inputs"
    FileName = "test_bad_inputs.csv"
    Expected = "#CSVRead: Delimiter character must be passed as a string, FALSE for no delimiter. Omit to guess from file contents!"
    TestRes = TestCSVRead(157, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        Delimiter:=99#, _
        IgnoreEmptyLines:=False, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test157 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test158(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test bad inputs"
    FileName = "test_bad_inputs.csv"
    Expected = "#CSVRead: SkipToCol (4) exceeds the number of columns in the file (3)!"
    TestRes = TestCSVRead(158, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        IgnoreEmptyLines:=False, _
        SkipToCol:=4, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test158 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test159(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test too few headers"
    FileName = "test_too_few_headers.csv"
    Expected = HStack( _
        Array("Col1", 1#, 11#, 21#, 31#, 41#, 51#, 61#, 71#, 81#, 91#), _
        Array("Col2", 2#, 12#, 22#, 32#, 42#, 52#, 62#, 72#, 82#, 92#), _
        Array("Col3", 3#, 13#, 23#, 33#, 43#, 53#, 63#, 73#, 83#, 93#), _
        Array("Col4", 4#, 14#, 24#, 34#, 44#, 54#, 64#, 74#, 84#, 94#), _
        Array("Col5", 5#, 15#, 25#, 35#, 45#, 55#, 65#, 75#, 85#, 95#), _
        Array(vbNullString, 6#, 16#, 26#, 36#, 46#, 56#, 66#, 76#, 86#, 96#), _
        Array(vbNullString, 7#, 17#, 27#, 37#, 47#, 57#, 67#, 77#, 87#, 97#), _
        Array(vbNullString, 8#, 18#, 28#, 38#, 48#, 58#, 68#, 78#, 88#, 98#), _
        Array(vbNullString, 9#, 19#, 29#, 39#, 49#, 59#, 69#, 79#, 89#, 99#), _
        Array(vbNullString, 10#, 20#, 30#, 40#, 50#, 60#, 70#, 80#, 90#, 100#))

    TestRes = TestCSVRead(159, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        IgnoreEmptyLines:=False, _
        ShowMissingsAs:=Empty, _
        HeaderRowNum:=1#, _
        SkipToRow:=1, _
        ExpectedHeaderRow:=HStack("Col1", "Col2", "Col3", "Col4", "Col5", vbNullString, vbNullString, vbNullString, vbNullString, vbNullString))
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test159 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test160(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test bad inputs"
    FileName = "test_bad_inputs.csv"
    Expected = "#CSVRead: #ParseConvertTypes: #ParseCTString: ConvertTypes must be Boolean or string with allowed letters NDBETQ. ""N"" show numbers as numbers, ""D"" show dates as dates, ""B"" show Booleans as Booleans, ""E"" show Excel errors as errors, ""T"" to trim leading and trailing spaces from fields, ""Q"" rules NDBE apply even to quoted fields, TRUE = ""NDB"" (convert unquoted numbers, dates and Booleans), FALSE = no conversion Found unrecognised character 'X'!!!"
    TestRes = TestCSVRead(160, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:="XYZ", _
        IgnoreEmptyLines:=False, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test160 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test161(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test quoted headers"
    FileName = "test_quoted_headers.csv"
    Expected = HStack( _
        Array("Col1", 1#, 11#, 21#, 31#, 41#, 51#, 61#, 71#, 81#, 91#), _
        Array("Col2", 2#, 12#, 22#, 32#, 42#, 52#, 62#, 72#, 82#, 92#), _
        Array("Col3", 3#, 13#, 23#, 33#, 43#, 53#, 63#, 73#, 83#, 93#), _
        Array("Col4", 4#, 14#, 24#, 34#, 44#, 54#, 64#, 74#, 84#, 94#), _
        Array("Col5", 5#, 15#, 25#, 35#, 45#, 55#, 65#, 75#, 85#, 95#), _
        Array("Col6", 6#, 16#, 26#, 36#, 46#, 56#, 66#, 76#, 86#, 96#), _
        Array("Col7", 7#, 17#, 27#, 37#, 47#, 57#, 67#, 77#, 87#, 97#), _
        Array("Col8", 8#, 18#, 28#, 38#, 48#, 58#, 68#, 78#, 88#, 98#), _
        Array("Col9", 9#, 19#, 29#, 39#, 49#, 59#, 69#, 79#, 89#, 99#), _
        Array("Col10", 10#, 20#, 30#, 40#, 50#, 60#, 70#, 80#, 90#, 100#))

    TestRes = TestCSVRead(161, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        IgnoreEmptyLines:=False, _
        ShowMissingsAs:=Empty, _
        HeaderRowNum:=1#, _
        SkipToRow:=1, _
        ExpectedHeaderRow:=HStack("Col1", "Col2", "Col3", "Col4", "Col5", "Col6", "Col7", "Col8", "Col9", "Col10"))
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test161 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test162(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test bad inputs"
    FileName = "test_bad_inputs.csv"
    'Error string varies according to date format on PC, so use simple regexp
    'Expected = "#CSVRead: #ParseDateFormat: DateFormat not valid should be one of 'ISO', 'ISOZ', 'M-D-Y', 'D-M-Y', 'Y-M-D', 'M/D/Y', 'D/M/Y' or 'Y/M/D'. Omit to use the default date format on this PC which is ""M/D/Y""!!"
    Expected = "DateFormat not valid"
    TestRes = TestCSVRead(162, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        DateFormat:="BAD", _
        IgnoreEmptyLines:=False, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test162 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test163(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test bad inputs"
    FileName = "test_bad_inputs.csv"
    'Error string varies according to date format on PC, so use simple regexp
    'Expected = "#CSVRead: #ParseDateFormat: DateFormat not valid should be one of 'ISO', 'ISOZ', 'M-D-Y', 'D-M-Y', 'Y-M-D', 'M/D/Y', 'D/M/Y' or 'Y/M/D'. Omit to use the default date format on this PC which is ""M/D/Y""!!"
    Expected = "DateFormat not valid"
    TestRes = TestCSVRead(163, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        DateFormat:="Y-M/D", _
        IgnoreEmptyLines:=False, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test163 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test164(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test bad inputs"
    FileName = "test_bad_inputs.csv"
    Expected = "#CSVRead: #ParseConvertTypes: #OneDArrayToTwoDArray: If ConvertTypes is given as a 1-dimensional array, each element must be a 1-dimensional array with two elements!!!"
    TestRes = TestCSVRead(164, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=Array(1.5, "N"), _
        IgnoreEmptyLines:=False, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test164 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test165(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test bad inputs"
    FileName = "test_bad_inputs.csv"
    Expected = "#CSVRead: #ParseConvertTypes: Column identifiers in the left column (or top row) of ConvertTypes must be strings or non-negative whole numbers but ConvertTypes(1,1) is 1.5!!"
    TestRes = TestCSVRead(165, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=HStack(Array(1.5, 2#), Array("N", "N")), _
        IgnoreEmptyLines:=False, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test165 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test166(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test not delimited.txt"
    FileName = "test_not_delimited.txt"
    Expected = HStack(Array("Use:", "]add CSV#main", "=#"))
    TestRes = TestCSVRead(166, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        Delimiter:=False, _
        IgnoreEmptyLines:=False, _
        SkipToRow:=3, _
        NumRows:=3, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test166 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test167(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test dateformat"
    FileName = "test_dateformat.csv"
    Expected = HStack( _
        Array("Date1", CDate("2021-Aug-30"), CDate("2021-Aug-31")), _
        Array("Date2", CDate("2021-Aug-31"), CDate("2021-Sep-01")))

    TestRes = TestCSVRead(167, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        DateFormat:="YYYY-MM-DD", _
        IgnoreEmptyLines:=False, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test167 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test168(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test bad inputs"
    FileName = "test_bad_inputs.csv"
    Expected = HStack(Array("Col1", "1", "4", "7"), Array("Col2", 2#, 5#, 8#), Array("Col3", "3", "6", "9"))
    TestRes = TestCSVRead(168, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=HStack(Array(1#, False), Array(2#, True), Array(3#, False)), _
        IgnoreEmptyLines:=False, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test168 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test169(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test bad inputs"
    FileName = "test_bad_inputs.csv"
    Expected = "#CSVRead: #ParseConvertTypes: ConvertTypes is contradictory. Column 2 is specified to be converted using two different conversion rules: B and N!!"
    TestRes = TestCSVRead(169, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=HStack(Array(1#, "NB"), Array(1#, "BN"), Array(2#, "N"), Array(2#, "B")), _
        IgnoreEmptyLines:=False, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test169 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test170(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "download airline safety"
    FileName = "https://raw.githubusercontent.com/fivethirtyeight/data/master/airline-safety/airline-safety.csv"
    Expected = HStack( _
        Array("airline", "Aer Lingus", "Aeroflot*", "Aerolineas Argentinas", "Aeromexico*", "Air Canada", "Air France", "Air India*", "Air New Zealand*", "Alaska Airlines*", "Alitalia", "All Nippon Airways", "American*", "Austrian Airlines", "Avianca", "British Airways*", "Cathay Pacific*", "China Airlines", "Condor", "COPA", "Delta / Northwest*", "Egyptair", "El Al", "Ethiopian Airlines", "Finnair", "Garuda Indonesia", "Gulf Air", "Hawaiian Airlines", "Iberia", "Japan Airlines", "Kenya Airways", "KLM*", "Korean Air", "LAN Airlines", "Lufthansa*", "Malaysia Airlines", "Pakistan International", "Philippine Airlines", "Qantas*", "Royal Air Maroc", "SAS*", "Saudi Arabian", "Singapore Airlines", "South African", "Southwest Airlines", "Sri Lankan / AirLanka", "SWISS*", "TACA", "TAM", "TAP - Air Portugal", "Thai Airways", "Turkish Airlines", "United / Continental*", "US Airways / America West*", "Vietnam Airlines", "Virgin Atlantic", "Xiamen Airlines"), _
        Array("avail_seat_km_per_week", 320906734#, 1197672318#, 385803648#, 596871813#, 1865253802#, 3004002661#, 869253552#, 710174817#, 965346773#, 698012498#, 1841234177#, 5228357340#, 358239823#, 396922563#, 3179760952#, 2582459303#, 813216487#, 417982610#, 550491507#, 6525658894#, 557699891#, 335448023#, 488560643#, 506464950#, 613356665#, 301379762#, 493877795#, 1173203126#, 1574217531#, 277414794#, 1874561773#, 1734522605#, 1001965891#, 3426529504#, 1039171244#, 348563137#, 413007158#, 1917428984#, 295705339#, 682971852#, 859673901#, 2376857805#, 651502442#, 3276525770#, 325582976#, 792601299#, 259373346#, 1509195646#, 619130754#, 1702802250#, 1946098294#, 7139291291#, 2455687887#, 625084918#, 1005248585#, 430462962#), _
        Array("incidents_85_99", 2#, 76#, 6#, 3#, 2#, 14#, 2#, 3#, 5#, 7#, 3#, 21#, 1#, 5#, 4#, 0#, 12#, 2#, 3#, 24#, 8#, 1#, 25#, 1#, 10#, 1#, 0#, 4#, 3#, 2#, 7#, 12#, 3#, 6#, 3#, 8#, 7#, 1#, 5#, 5#, 7#, 2#, 2#, 1#, 2#, 2#, 3#, 8#, 0#, 8#, 8#, 19#, 16#, 7#, 1#, 9#), _
        Array("fatal_accidents_85_99", 0#, 14#, 0#, 1#, 0#, 4#, 1#, 0#, 0#, 2#, 1#, 5#, 0#, 3#, 0#, 0#, 6#, 1#, 1#, 12#, 3#, 1#, 5#, 0#, 3#, 0#, 0#, 1#, 1#, 0#, 1#, 5#, 2#, 1#, 1#, 3#, 4#, 0#, 3#, 0#, 2#, 2#, 1#, 0#, 1#, 1#, 1#, 3#, 0#, 4#, 3#, 8#, 7#, 3#, 0#, 1#), _
        Array("fatalities_85_99", 0#, 128#, 0#, 64#, 0#, 79#, 329#, 0#, 0#, 50#, 1#, 101#, 0#, 323#, 0#, 0#, 535#, 16#, 47#, 407#, 282#, 4#, 167#, 0#, 260#, 0#, 0#, 148#, 520#, 0#, 3#, 425#, 21#, 2#, 34#, 234#, 74#, 0#, 51#, 0#, 313#, 6#, 159#, 0#, 14#, 229#, 3#, 98#, 0#, 308#, 64#, 319#, 224#, 171#, 0#, 82#), _
        Array("incidents_00_14", 0#, 6#, 1#, 5#, 2#, 6#, 4#, 5#, 5#, 4#, 7#, 17#, 1#, 0#, 6#, 2#, 2#, 0#, 0#, 24#, 4#, 1#, 5#, 0#, 4#, 3#, 1#, 5#, 0#, 2#, 1#, 1#, 0#, 3#, 3#, 10#, 2#, 5#, 3#, 6#, 11#, 2#, 1#, 8#, 4#, 3#, 1#, 7#, 0#, 2#, 8#, 14#, 11#, 1#, 0#, 2#), _
        Array("fatal_accidents_00_14", 0#, 1#, 0#, 0#, 0#, 2#, 1#, 1#, 1#, 0#, 0#, 3#, 0#, 0#, 0#, 0#, 1#, 0#, 0#, 2#, 1#, 0#, 2#, 0#, 2#, 1#, 0#, 0#, 0#, 2#, 0#, 0#, 0#, 0#, 2#, 2#, 1#, 0#, 0#, 1#, 0#, 1#, 0#, 0#, 0#, 0#, 1#, 2#, 0#, 1#, 2#, 2#, 2#, 0#, 0#, 0#), _
        Array("fatalities_00_14", 0#, 88#, 0#, 0#, 0#, 337#, 158#, 7#, 88#, 0#, 0#, 416#, 0#, 0#, 0#, 0#, 225#, 0#, 0#, 51#, 14#, 0#, 92#, 0#, 22#, 143#, 0#, 0#, 0#, 283#, 0#, 0#, 0#, 0#, 537#, 46#, 1#, 0#, 0#, 110#, 0#, 83#, 0#, 0#, 0#, 0#, 3#, 188#, 0#, 1#, 84#, 109#, 23#, 0#, 0#, 0#))

    TestRes = TestCSVRead(170, TestDescription, Expected, FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        IgnoreEmptyLines:=False, _
        ShowMissingsAs:=Empty, _
        HeaderRowNum:=1#, _
        SkipToRow:=1, _
        ExpectedHeaderRow:=HStack( _
        "airline", _
        "avail_seat_km_per_week", _
        "incidents_85_99", _
        "fatal_accidents_85_99", _
        "fatalities_85_99", _
        "incidents_00_14", _
        "fatal_accidents_00_14", _
        "fatalities_00_14"))
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test170 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test171(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test column-by-column"
    FileName = "test_column-by-column.csv"
    Expected = HStack( _
        Array("Type", "Number", "Date", "Boolean", "Error", "String", "String", "String", "String", "String", "String", "String"), _
        Array("Col A", 44424#, "2021-08-24T12:49:13", "True", "#DIV/0!", "1", "2021-08-24T12:49:13", "TRUE", "#DIV/0!", "abc", "abc""def", "Line" + vbCrLf + "Feed"), _
        Array("Col B", "44424", "2021-08-24T12:49:13", True, "#DIV/0!", "1", "2021-08-24T12:49:13", "TRUE", "#DIV/0!", "abc", "abc""def", "Line" + vbCrLf + "Feed"), _
        Array("Col C", "44424", CDate("2021-Aug-24 12:49:13"), "True", "#DIV/0!", "1", "2021-08-24T12:49:13", "TRUE", "#DIV/0!", "abc", "abc""def", "Line" + vbCrLf + "Feed"), _
        Array("Col D", "44424", "2021-08-24T12:49:13", "True", CVErr(2007), "1", "2021-08-24T12:49:13", "TRUE", "#DIV/0!", "abc", "abc""def", "Line" + vbCrLf + "Feed"), _
        Array("Col E", 44424#, "2021-08-24T12:49:13", "True", "#DIV/0!", 1#, "2021-08-24T12:49:13", "TRUE", "#DIV/0!", "abc", "abc""def", "Line" + vbCrLf + "Feed"), _
        Array("Col F", "44424", "2021-08-24T12:49:13", True, "#DIV/0!", "1", "2021-08-24T12:49:13", True, "#DIV/0!", "abc", "abc""def", "Line" + vbCrLf + "Feed"), _
        Array("Col G", "44424", "2021-08-24T12:49:13", "True", CVErr(2007), "1", "2021-08-24T12:49:13", "TRUE", CVErr(2007), "abc", "abc""def", "Line" + vbCrLf + "Feed"))

    TestRes = TestCSVRead(171, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=HStack( _
        Array(0#, False), _
        Array(2#, "N"), _
        Array(3#, "B"), _
        Array(4#, "D"), _
        Array(5#, "E"), _
        Array(6#, "NQ"), _
        Array(7#, "BQ"), _
        Array(8#, "EQ")), _
        DateFormat:="ISO", _
        IgnoreEmptyLines:=False, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test171 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test172(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test column-by-column"
    FileName = "test_column-by-column.csv"
    Expected = "#CSVRead: #ParseConvertTypes: Column identifiers in the left column (or top row) of ConvertTypes must be strings or non-negative whole numbers but ConvertTypes(1,1) is -2!!"
    TestRes = TestCSVRead(172, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=HStack( _
        Array(-2#, False), _
        Array(2#, "N"), _
        Array(3#, "B"), _
        Array(4#, "D"), _
        Array(5#, "E"), _
        Array(6#, "NQ"), _
        Array(7#, "BQ"), _
        Array(8#, "EQ")), _
        IgnoreEmptyLines:=False, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test172 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test173(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test column-by-column"
    FileName = "test_column-by-column.csv"
    Expected = "#CSVRead: #ParseConvertTypes: Type Conversion given in bottom row (or right column) of ConvertTypes must be Booleans or strings containing letters NDBETQ but ConvertTypes(2,2) is string ""XX""!!"
    TestRes = TestCSVRead(173, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=HStack(Array(0#, False), Array(2#, "XX"), Array(3#, "B"), Array(4#, "D")), _
        IgnoreEmptyLines:=False, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test173 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test174(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test column-by-column"
    FileName = "test_column-by-column.csv"
    Expected = "#CSVRead: #ParseConvertTypes: Type Conversion given in bottom row (or right column) of ConvertTypes must be Booleans or strings containing letters NDBETQ but ConvertTypes(2,1) is of type Error!!"
    TestRes = TestCSVRead(174, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=HStack(Array(3#, CVErr(2007)), Array(4#, "D"), Array(5#, "E")), _
        IgnoreEmptyLines:=False, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test174 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test175(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test column-by-column"
    FileName = "test_column-by-column.csv"
    Expected = "#CSVRead: #ParseConvertTypes: ConvertTypes specifies columns by their header (instead of by number), but HeaderRowNum has not been specified!!"
    TestRes = TestCSVRead(175, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=HStack(Array("Col A", "N"), Array("Col B", "N"), Array("Col C", "N")), _
        IgnoreEmptyLines:=False, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test175 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test176(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test various time formats"
    FileName = "test_various_time_formats.csv"
    Expected = ThrowIfError(CSVRead(Folder + FileName, ConvertTypes:="N", SkipToRow:=2, SkipToCol:=2, NumCols:=1))
    CastDoublesToDates Expected
    TestRes = TestCSVRead(176, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:="D", _
        DateFormat:="Y-M-D", _
        IgnoreEmptyLines:=False, _
        SkipToRow:=2, _
        SkipToCol:=1, _
        NumCols:=1, _
        AbsTol:=0.000000000000001, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test176 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test177(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test y-m-d dates with fractional seconds"
    FileName = "test_y-m-d_dates_with_fractional_seconds.csv"
    Expected = ThrowIfError(CSVRead(Folder + FileName, ConvertTypes:="N", SkipToRow:=2, SkipToCol:=2, NumCols:=1))
    CastDoublesToDates Expected
    TestRes = TestCSVRead(177, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:="D", _
        DateFormat:="Y-M-D", _
        IgnoreEmptyLines:=False, _
        SkipToRow:=2, _
        SkipToCol:=1, _
        NumCols:=1, _
        AbsTol:=0.0000000001, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test177 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test178(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "Used as example in README.md"
    FileName = "https://vincentarelbundock.github.io/Rdatasets/csv/carData/TitanicSurvival.csv"
    Expected = HStack( _
        Array(vbNullString, "Allen, Miss. Elisabeth Walton", "Allison, Master. Hudson Trevor", "Allison, Miss. Helen Loraine", "Allison, Mr. Hudson Joshua Crei", "Allison, Mrs. Hudson J C (Bessi", "Anderson, Mr. Harry", "Andrews, Miss. Kornelia Theodos", "Andrews, Mr. Thomas Jr", "Appleton, Mrs. Edward Dale (Cha", "Artagaveytia, Mr. Ramon", "Astor, Col. John Jacob", "Astor, Mrs. John Jacob (Madelei", "Aubart, Mme. Leontine Pauline", "Barber, Miss. Ellen Nellie", "Barkworth, Mr. Algernon Henry W", "Baumann, Mr. John D"), _
        Array("survived", "yes", "yes", "no", "no", "no", "yes", "yes", "no", "yes", "no", "no", "yes", "yes", "yes", "yes", "no"), _
        Array("sex", "female", "male", "female", "male", "female", "male", "female", "male", "female", "male", "male", "female", "female", "female", "male", "male"), _
        Array("age", 29#, 0.916700006, 2#, 30#, 25#, 48#, 63#, 39#, 53#, 71#, 47#, 18#, 24#, 26#, 80#, Empty), _
        Array("passengerClass", "1st", "1st", "1st", "1st", "1st", "1st", "1st", "1st", "1st", "1st", "1st", "1st", "1st", "1st", "1st", "1st"))
    TestRes = TestCSVRead(178, TestDescription, Expected, FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        IgnoreEmptyLines:=False, _
        NumRows:=17, _
        MissingStrings:="NA", _
        ShowMissingsAs:=Empty, _
        HeaderRowNum:=1#, _
        SkipToRow:=1, _
        ExpectedHeaderRow:=HStack(vbNullString, "survived", "sex", "age", "passengerClass"))
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test178 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test179(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test headers in fixed width file"
    Expected = HStack( _
        Array("Col1", 11970#, 2041#, 42004#, "xyz"), _
        Array("Col2", 37#, 26#, 0#, 137693#), _
        Array("Col3", 11721#, 36#, 33236#, 5175#), _
        Array("Col4", 135#, 85#, 115#, 3#), _
        Array("Col5", 79270#, 214#, 1#, 3650#), _
        Array("Col6", 9066#, 20#, 80821#, 203158#))
    FileName = "test_headers_in_fixed_width_file.csv"
    
    TestRes = TestCSVRead(179, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        Delimiter:=" ", _
        IgnoreRepeated:=True, _
        IgnoreEmptyLines:=False, _
        ShowMissingsAs:=Empty, _
        HeaderRowNum:=1#, _
        SkipToRow:=1, _
        ExpectedHeaderRow:=HStack("Col1", "Col2", "Col3", "Col4", "Col5", "Col6"))
        
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test179 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test180(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test array lower bounds"
    Expected = HStack(Array("Col1", 1#, 2#, 3#), Array("Col2", 4#, 5#, 6#), Array("Col3", 7#, 8#, 9#))
    FileName = "test_array_lower_bounds.csv"
    TestRes = TestCSVRead(180, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        IgnoreEmptyLines:=False, _
        ShowMissingsAs:=Empty)
        
    If TestRes Then
        If LBound(Observed, 1) <> 1 Then
            TestRes = False
            WhatDiffers = "Test 180 test array lowwer bounds FAILED, Test was that array lower bound should be 1"
        End If
    End If

    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test180 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test181(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test literal with no line feed"
    Expected = HStack(1#, 2#, 3#)
    FileName = "1,2,3"
    TestRes = TestCSVRead(181, TestDescription, Expected, FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        IgnoreEmptyLines:=False, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test181 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : TwoFiveFiveChars
' Purpose    : Returns the contents of files 255_characters-ANSI.csv, 255_characters-UTF-8.csv,
'              255_characters-UTF-8-BOM.csv, 255_characters-UTF-16-BE-BOM.csv, 255_characters-UTF-16-LE-BOM.csv
' -----------------------------------------------------------------------------------------------------------------------
Function TwoFiveFiveChars() As Variant
    Dim i As Long
    Dim Res As Variant

    On Error GoTo ErrHandler
    Res = Fill(vbNullString, 256, 2)
    Res(1, 1) = "N"
    Res(1, 2) = "Char"
    For i = 2 To 256
        Res(i, 1) = CStr(i - 1)
        Res(i, 2) = Chr$(i - 1)
    Next
    Res(11, 2) = vbCrLf 'Have CRLF here to counteract git's annoying habit of "correcting" mixed line endings when pushing and pulling from remote
    Res(14, 2) = vbCrLf

    TwoFiveFiveChars = Res

    Exit Function
ErrHandler:
    Throw "#TwoFiveFiveChars (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Sub Test182(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "255 characters-ANSI"
    FileName = "255_characters-ANSI.csv"
    Expected = TwoFiveFiveChars()
    TestRes = TestCSVRead(182, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=False, _
        IgnoreEmptyLines:=False, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test182 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test183(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "255 characters-UTF-8"
    Expected = TwoFiveFiveChars()
    FileName = "255_characters-UTF-8.csv"
    TestRes = TestCSVRead(183, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        IgnoreEmptyLines:=False, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test183 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test184(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "255 characters-UTF-8-BOM"
    Expected = TwoFiveFiveChars()
    FileName = "255_characters-UTF-8-BOM.csv"
    TestRes = TestCSVRead(184, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        IgnoreEmptyLines:=False, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test184 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test185(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "255 characters-UTF-16-BE-BOM"
    Expected = TwoFiveFiveChars()
    FileName = "255_characters-UTF-16-BE-BOM.csv"
    TestRes = TestCSVRead(185, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        IgnoreEmptyLines:=False, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures
    Exit Sub
ErrHandler:
    Throw "#Test185 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test186(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "255 characters-UTF-16-LE-BOM"
    Expected = TwoFiveFiveChars()
    FileName = "255_characters-UTF-16-LE-BOM.csv"
    TestRes = TestCSVRead(186, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        IgnoreEmptyLines:=False, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures
    Exit Sub
ErrHandler:
    Throw "#Test186 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test187(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test mac line endings"
    Expected = HStack(Array("1,2,3", "4,5,6"))
    FileName = "test_mac_line_endings.csv"
    TestRes = TestCSVRead(187, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        Delimiter:=False, _
        IgnoreEmptyLines:=False, _
        SkipToRow:=2, _
        NumRows:=2, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test187 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test188(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "FileName passed as CSV Contents, parse as not delimited"
    Expected = HStack(Array("2", "3", "4", "5", "6"))
    FileName = "1" + vbCrLf + "2" + vbCrLf + "3" + vbCrLf + "4" + vbCrLf + "5" + vbCrLf + "6" + vbCrLf + "7" + vbCrLf + "8" + vbCrLf + "9" + vbCrLf + "10" + vbCrLf + vbNullString
    TestRes = TestCSVRead(188, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        Delimiter:=False, _
        IgnoreEmptyLines:=False, _
        SkipToRow:=2, _
        NumRows:=5, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test188 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test189(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "FileName passed as CSV Contents, parse as delimited"
    Expected = HStack(Array(5#, 8#), Array(6#, 9#))
    FileName = "1,2,3" + vbCrLf + "4,5,6" + vbCrLf + "7,8,9" + vbCrLf + "10,11,12" + vbCrLf + "13,14,15" + vbCrLf + "16,17,18" + vbCrLf + vbNullString
    TestRes = TestCSVRead(189, TestDescription, Expected, FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        IgnoreEmptyLines:=False, _
        SkipToRow:=2, _
        SkipToCol:=2, _
        NumRows:=2, _
        NumCols:=2, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test189 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test190(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "norwegian data"
    Expected = Empty
    FileName = "norwegian_data.csv"
    TestRes = TestCSVRead(190, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        DateFormat:="Y-M-D", _
        IgnoreEmptyLines:=False, _
        MissingStrings:="NULL", _
        ShowMissingsAs:=Empty, _
        Encoding:="UTF-8", _
        HeaderRowNum:=1#, _
        NumRowsExpected:=1230, _
        NumColsExpected:=83, _
        ExpectedHeaderRow:=HStack("regine_area", "main_no", "point_no", "param_key", "version_no_end", "station_name", "station_status_name", "dt_start_date", "dt_end_date", "percent_missing_days", _
        "first_year_regulation", "start_year", "end_year", "aktuell_avrenningskart", "excluded_years", "tilgang", "latitude", "longitude", "utm_east_z33", "utm_north_z33", _
        "regulation_part_area", "regulation_part_reservoirs", "transfer_area_in", "transfer_area_out", "drainage_basin_key", "area_norway", "area_total", "comment", "drainage_dens", "dt_registration_date", _
        "dt_regul_date", "gradient_1085", "gradient_basin", "gradient_river", "height_minimum", "height_hypso_10", "height_hypso_20", "height_hypso_30", "height_hypso_40", "height_hypso_50", _
        "height_hypso_60", "height_hypso_70", "height_hypso_80", "height_hypso_90", "height_maximum", "length_km_basin", "length_km_river", "ocean_polar_angle", "ocean_polar_distance", "perc_agricul", _
        "perc_bog", "perc_eff_bog", "perc_eff_lake", "perc_forest", "perc_glacier", "perc_lake", "perc_mountain", "perc_urban", "prec_intens_max", "utm_zone_gravi", _
        "utm_east_gravi", "utm_north_gravi", "utm_zone_inlet", "utm_east_inlet", "utm_north_inlet", "br1_middelavrenning_1930_1960", "br2_Tilsigsberegning", "br3_Regional_flomfrekvensanalyse", "br5_Regional_lavvannsanalyse", "br6_Klimastudier", _
        "br7_Klimascenarier", "br9_Flomvarsling", "br11_FRIEND", "br12_GRDC", "br23_HBV", "br24_middelavrenning_1961_1990", "br26_TotalAvlop", "br31_FlomserierPrim", "br32_FlomserierSekundar", "br33_Flomkart_aktive_ureg", _
        "br34_Hydrologisk_referanseserier_klimastudier", "br38_Flomkart_aktive_ureg_periode", "br39_Flomkart_nedlagt_stasjon"))

    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test190 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test191(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "csv segfault.txt"
    Expected = Empty
    FileName = "csv_segfault.txt"
    TestRes = TestCSVRead(191, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        NumRowsExpected:=468, _
        NumColsExpected:=9, _
        IgnoreEmptyLines:=True, _
        ShowMissingsAs:=Empty, _
        Encoding:="UTF-8", _
        HeaderRowNum:=1#, _
        ExpectedHeaderRow:=HStack("Time (CEST)", "Latitude", "Longitude", "Course", "kts", "mph", "feet", "Rate", "Reporting Facility"))
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test191 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test192(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "Enedis smart meter data"
    Expected = Empty
    FileName = "Enedis_smart_meter_data.csv"
    TestRes = TestCSVRead(192, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        NumRowsExpected:=415, _
        NumColsExpected:=47, _
        ConvertTypes:=True, _
        DateFormat:="ISOZ", _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test192 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test193(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "Enedis smart meter data"
    Expected = HStack( _
        CDate("2020-Aug-01 22:00:00"), _
        "Arr?t? quotidien", _
        26575641#, _
        61358114#, _
        Empty, _
        Empty, _
        Empty, _
        Empty, _
        Empty, _
        Empty, _
        Empty, _
        Empty, _
        23340318#, _
        55115006#, _
        3235323#, _
        6243108#, _
        87933755#)
    FileName = "Enedis_smart_meter_data.csv"
    TestRes = TestCSVRead(193, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        DateFormat:="ISOZ", _
        NumRows:=1, _
        ShowMissingsAs:=Empty, _
        HeaderRowNum:=3#, _
        ExpectedHeaderRow:=HStack( _
        "Horodate", _
        "Type de releve", _
        "EAS F1", _
        "EAS F2", _
        "EAS F3", _
        "EAS F4", _
        "EAS F5", _
        "EAS F6", _
        "EAS F7", _
        "EAS F8", _
        "EAS F9", _
        "EAS F10", _
        "EAS D1", _
        "EAS D2", _
        "EAS D3", _
        "EAS D4", _
        "EAS T"))
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test193 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test194(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "utf 8 with bom"
    Expected = HStack(Array(1#, 5#, 9#), Array(2#, 6#, 10#), Array(3#, 7#, 11#), Array(4#, 8#, 12#))
    FileName = "utf_8_with_bom.csv"
    TestRes = TestCSVRead(194, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test194 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test195(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test bad D-M-Y"
    FileName = "test_bad_D-M-Y.csv"
    Expected = ThrowIfError(CSVRead(Folder + FileName, False))
    TestRes = TestCSVRead(195, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        NumRowsExpected:=776, _
        NumColsExpected:=2, _
        ConvertTypes:=True, _
        DateFormat:="D-M-Y", _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test195 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test196(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test bad M-D-Y"
    FileName = "test_bad_M-D-Y.csv"
    Expected = ThrowIfError(CSVRead(Folder + FileName, False))
    TestRes = TestCSVRead(196, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        NumRowsExpected:=776, _
        NumColsExpected:=2, _
        ConvertTypes:=True, _
        DateFormat:="M-D-Y", _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test196 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test197(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test bad Y-M-D"
    FileName = "test_bad_Y-M-D.csv"
    Expected = ThrowIfError(CSVRead(Folder + FileName, False))
    TestRes = TestCSVRead(197, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        NumRowsExpected:=776, _
        NumColsExpected:=2, _
        ConvertTypes:=True, _
        DateFormat:="Y-M-D", _
        ShowMissingsAs:=Empty)
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test197 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test198(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test good D-M-Y"
    FileName = "test_good_D-M-Y.csv"
    Expected = ThrowIfError(CSVRead(Folder + FileName, "N", SkipToRow:=2, SkipToCol:=2, NumCols:=1))
    CastDoublesToDates Expected
    
    TestRes = TestCSVRead(198, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        DateFormat:="D-M-Y", _
        SkipToRow:=2, _
        NumCols:=1, _
        ShowMissingsAs:=Empty, _
        AbsTol:=1 / 24 / 60 / 60 / 1000 / 100) '10 microsecond tolerance
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test198 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test199(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test good M-D-Y"
    FileName = "test_good_M-D-Y.csv"
    Expected = ThrowIfError(CSVRead(Folder + FileName, "N", SkipToRow:=2, SkipToCol:=2, NumCols:=1))
    CastDoublesToDates Expected
    
    TestRes = TestCSVRead(199, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        DateFormat:="M-D-Y", _
        SkipToRow:=2, _
        NumCols:=1, _
        ShowMissingsAs:=Empty, _
        AbsTol:=1 / 24 / 60 / 60 / 1000 / 100) '10 microsecond tolerance
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test199 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Test200(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Observed As Variant
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test good Y-M-D"
    FileName = "test_good_Y-M-D.csv"
    Expected = ThrowIfError(CSVRead(Folder + FileName, "N", SkipToRow:=2, SkipToCol:=2, NumCols:=1))
    CastDoublesToDates Expected
    
    TestRes = TestCSVRead(200, TestDescription, Expected, Folder + FileName, Observed, WhatDiffers, _
        ConvertTypes:=True, _
        DateFormat:="Y-M-D", _
        SkipToRow:=2, _
        NumCols:=1, _
        ShowMissingsAs:=Empty, _
        AbsTol:=1 / 24 / 60 / 60 / 1000 / 100) '10 microsecond tolerance
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures

    Exit Sub
ErrHandler:
    Throw "#Test200 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

'Non-standard test, since we are testing behaviour which is "From Excel sheet, not from VBA"
Private Sub Test201(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Formula As String
    Dim Observed As Variant
    Dim R As Range
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim wb As Workbook
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test 32K limit"
    Expected = "#CSVRead: The file has a field (row 2, column 2) of length 33,000. Excel cells cannot contain strings longer than 32,767!"
    FileName = "test_32K_limit.csv"
    
    Formula = "=CSVRead(""" & Folder & FileName & """, TRUE)"
    
    shHiddenSheet.Unprotect
    shHiddenSheet.UsedRange.Clear
    Set R = shHiddenSheet.Cells(1, 1)
    R.Formula2 = Formula
    Observed = R.value
    TestRes = Observed = Expected
    If Not TestRes Then WhatDiffers = "Observed = '" & Observed & "' Expected = '" & Expected & "'"
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures
    shHiddenSheet.UsedRange.Clear

    Exit Sub
ErrHandler:
    Throw "#Test201 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

'Non-standard test, since we are testing behaviour which is "From Excel sheet, not from VBA"
Private Sub Test202(Folder As String, ByRef NumPassed As Long, ByRef NumFailed As Long, ByRef Failures() As String)
    Dim Expected As Variant
    Dim FileName As String
    Dim Formula As String
    Dim Observed As Variant
    Dim R As Range
    Dim TestDescription As String
    Dim TestRes As Boolean
    Dim wb As Workbook
    Dim WhatDiffers As String

    On Error GoTo ErrHandler
    TestDescription = "test 32K limit"
    Expected = "#CSVRead: #ParseTextFile: Line 2 of the file is of length 33,002. Excel cells cannot contain strings longer than 32,767!!"
    FileName = "test_32K_limit.csv"
    
    Formula = "=CSVRead(""" & Folder & FileName & """, FALSE,FALSE)"
    
    shHiddenSheet.Unprotect
    shHiddenSheet.UsedRange.Clear
    Set R = shHiddenSheet.Cells(1, 1)
    R.Formula2 = Formula
    Observed = R.value
    TestRes = Observed = Expected
    If Not TestRes Then WhatDiffers = "Observed = '" & Observed & "' Expected = '" & Expected & "'"
    AccumulateResults TestRes, NumPassed, NumFailed, WhatDiffers, Failures
    shHiddenSheet.UsedRange.Clear

    Exit Sub
ErrHandler:
    Throw "#Test202 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

