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

    Const Unicode = False
    Dim Data As Variant
    Dim DataReread1
    Dim DataReread2
    Dim DataReread3
    Dim DataReread4
    Dim ExtraInfo As String
    Dim FileName As String
    Dim FnName As String
    Dim i As Long
    Dim j As Long
    Dim NumCols As Long
    Dim NumRows As Long
    Dim OS As String
    Dim QuoteAllStrings As Boolean
    Dim SmallFileName As String
    Dim StringLength As Double
    Dim t1 As Double
    Dim t2 As Double
    Dim t3 As Double
    Dim t4 As Double
    Dim tend As Double
    Dim tstart As Double

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
                Data = RandomDoubles(NumRows, NumCols)
                ExtraInfo = "random doubles"
                QuoteAllStrings = False
            Case 2
                Data = sFill("abcdefghij", NumRows, NumCols)
                ExtraInfo = "10 char Strings unquoted"
                QuoteAllStrings = False
            Case 3
                Data = sFill(String(StringLength, "x"), NumRows, NumCols)
                ExtraInfo = CStr(Len(Data(1, 1))) & " char strings quoted"
                QuoteAllStrings = True
            Case 4
                Data = sFill(String(StringLength / 2, "x") + vbCrLf + String((StringLength / 2) - 1, "y"), NumRows, NumCols)
                ExtraInfo = CStr(Len(Data(1, 1))) & " char strings quoted with line feeds"
                QuoteAllStrings = True
        End Select

        FileName = NameThatFile(m_FolderSpeedTest, OS, NumRows, NumCols, Replace(ExtraInfo, " ", "-"), Unicode, False)
        ThrowIfError CSVWrite(FileName, Data, QuoteAllStrings, , , , Unicode, OS)
        
        Debug.Print "FileName = " & FileName
        Debug.Print "Contains " + ExtraInfo + " " + _
            Format(NumRows, "###,##0") + " rows, " + Format(NumCols, "###,##0") + " cols. " '+ _
            "File size = " + Format(sFileInfo(FileName, "size"), "###,##0") + " bytes."
        For j = 1 To 4
            tstart = sElapsedTime
            Select Case j
                Case 1
                    DataReread1 = ThrowIfError(CSVRead(FileName, False, ",", , , , , , False))
                    FnName = "CSVRead       "
                Case 2
                    DataReread2 = ThrowIfError(CSVRead_sdkn104(FileName, Unicode))
                    FnName = "CSVRead_sdkn104  "
                Case 3
                    DataReread3 = ThrowIfError(CSVRead_ws_garcia(FileName, ",", vbCrLf))
                    FnName = "CSVRead_ws_garcia"
                Case 4
                    'DataReread4 = ThrowIfError(sFileShow(FileName, ",", False, , False, vbCrLf, , , , False, , , , , False))
                    'FnName = "sFileShow        "
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
        Debug.Print "v sdk104          " & CStr(t2 / t1) & "           >1 = CSVRead faster"
        Debug.Print "v garcia          " & CStr(t3 / t1) & "           >1 = CSVRead faster"
        Debug.Print "v sFileShow       " & CStr(t4 / t1) & "           >1 = CSVRead faster"

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
' Procedure  : TimeThreeParsers
' Author     : Philip Swannell
' Date       : 07-Aug-2021
' Purpose    : For use from sheet TimingResults - compares speed of 3 CSV parsing functions
' -----------------------------------------------------------------------------------------------------------------------
Function TimeThreeParsers(EachFieldContains As Variant, NumRows As Long, NumCols As Long, Optional Timeout As Double = 1, Optional WithHeaders As Boolean)

    Const Unicode = False
    Dim Data As Variant
    Dim DataReread1
    Dim DataReread2
    Dim DataReread3
    Dim DataRow As Long
    Dim ExtraInfo As String
    Dim FileName As String
    Dim FnName As String
    Dim FnName1 As String
    Dim FnName2 As String
    Dim FnName3 As String
    Dim i As Long
    Dim j As Long
    Dim k As Double
    Dim NumCalls1 As Long
    Dim NumCalls2 As Long
    Dim NumCalls3 As Long
    Dim OneEqThree As Boolean
    Dim OneEqTwo As Boolean
    Dim OS As String
    Dim Ret As Variant
    Dim t1 As Double
    Dim t2 As Double
    Dim t3 As Double
    Dim tend As Double
    Dim tstart As Double

    On Error GoTo ErrHandler
    OS = "Windows"
    
    If VarType(EachFieldContains) = vbDouble Then
        ExtraInfo = "Doubles"
    ElseIf VarType(EachFieldContains) = vbString Then
        If Left(EachFieldContains, 1) = """" & Right(EachFieldContains, 1) = """" Then
            ExtraInfo = "Quoted_Strings_length_" & Len(EachFieldContains)
        Else
            ExtraInfo = "Strings_length_" & Len(EachFieldContains)
        End If
    Else
        ExtraInfo = "Unknown"
    End If

    ThrowIfError CreatePath(m_FolderSpeedTest)

    Data = sFill(EachFieldContains, NumRows, NumCols)
    FileName = NameThatFile(m_FolderSpeedTest, OS, NumRows, NumCols, Replace(ExtraInfo, " ", "-"), Unicode, False)
    ThrowIfError Application.Run("sFileSave", FileName, Data, ",", , , , True)
        
    For j = 1 To 6
        tstart = sElapsedTime()
        k = 0
        Do
            k = k + 1
            Select Case j
                Case 1
                    FnName1 = "CSVRead"
                    DataReread1 = ThrowIfError(CSVRead(FileName, False, ",", , , , , , Unicode))
                Case 2
                    FnName2 = "CSVRead_sdkn104"
                    DataReread2 = ThrowIfError(CSVRead_sdkn104(FileName, Unicode))
                Case 3
                    FnName3 = "CSVRead_ws_garcia"
                    DataReread3 = ThrowIfError(CSVRead_ws_garcia(FileName, ",", vbCrLf))
            End Select
            If sElapsedTime() - tstart > Timeout Then Exit Do
        Loop

        tend = sElapsedTime()
        Select Case j
            Case 1
                NumCalls1 = k
                t1 = (tend - tstart) / k
            Case 2
                NumCalls2 = k
                t2 = (tend - tstart) / k
            Case 3
                NumCalls3 = k
                t3 = (tend - tstart) / k
        End Select
    Next j

    'Hook in to SolumAddin. TODO version of sArraysIdentical to TestDeps?
    OneEqTwo = Application.Run("sArraysIdentical", DataReread1, DataReread2)
    'Comparing arrays but allowing for different lower bounds
    OneEqThree = Application.Run("sArraysIdentical", DataReread1, DataReread3, True, True)
                

    ReDim Ret(1 To IIf(WithHeaders, 2, 1), 1 To 10) As Variant
    
    DataRow = IIf(WithHeaders, 2, 1)
    
    Ret(DataRow, 1) = t1: If WithHeaders Then Ret(1, 1) = FnName1
    Ret(DataRow, 2) = t2: If WithHeaders Then Ret(1, 2) = FnName2
    Ret(DataRow, 3) = t3: If WithHeaders Then Ret(1, 3) = FnName3
    Ret(DataRow, 4) = NumCalls1: If WithHeaders Then Ret(1, 4) = "NCalls " + FnName1
    Ret(DataRow, 5) = NumCalls2: If WithHeaders Then Ret(1, 5) = "NCalls " + FnName2
    Ret(DataRow, 6) = NumCalls3: If WithHeaders Then Ret(1, 6) = "NCalls " + FnName3

    Ret(DataRow, 7) = OneEqTwo: If WithHeaders Then Ret(1, 7) = "1 = 2?"
    Ret(DataRow, 8) = OneEqThree: If WithHeaders Then Ret(1, 8) = "1 = 3?"
    Ret(DataRow, 9) = FileName: If WithHeaders Then Ret(1, 9) = "File"
    Ret(DataRow, 10) = FileSize(FileName): If WithHeaders Then Ret(1, 10) = "Size"
    

    TimeThreeParsers = Ret

    Exit Function
ErrHandler:
    TimeThreeParsers = "#TimeThreeParsers: " & Err.Description & "!"
End Function

