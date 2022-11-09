Attribute VB_Name = "modCSVTestRDatasets"
Option Explicit

' returns largest 20 files in RDatasets
Function RDatesetsFiles()

RDatesetsFiles = Array("csv\openintro\military.csv", "csv\mosaicData\Birthdays.csv", _
        "csv\stevedata\wvs_justifbribe.csv", "csv\nycflights13\flights.csv", _
        "csv\stevedata\wvs_immig.csv", "csv\AER\Fertility.csv", _
        "csv\openintro\avandia.csv", "csv\Stat2Data\AthleteGrad.csv", _
        "csv\causaldata\mortgages.csv", "csv\openintro\mammogram.csv", _
        "csv\lme4\InstEval.csv", "csv\stevedata\gss_abortion.csv", _
        "csv\stevedata\TV16.csv", "csv\stevedata\gss_wages.csv", _
        "csv\AER\CPSSW8.csv", "csv\stevedata\eq_passengercars.csv", _
        "csv\ggplot2movies\movies.csv", "csv\ggplot2\diamonds.csv", _
        "csv\causaldata\gov_transfers_density.csv", "csv\openintro\seattlepets.csv")

End Function

'Tested 5 Oct 2021, CSVRead against ArrayFromCSVfile, Nigel Hefferman's code given as an answer at
'https://stackoverflow.com/questions/12259595/load-csv-file-into-a-vba-array-rather-than-excel-sheet

'With RemoveQuotes = True
'CSVRead     23.285191300005            ArrayFromCSVfile             281.768168099996,          True
'CSVRead     23.8145522999985           ArrayFromCSVfile             325.074338899998           True

'With RemoveQuotes = False
'CSVRead     24.9822067999921           ArrayFromCSVfile             18.9174217000109           False

Sub TestAgainstLargestFileInRDatasets()

    Const FileName As String = "C:\Projects\RDatasets\csv\openintro\military.csv"
    Dim res1
    Dim res2
    Dim t1 As Double
    Dim t2 As Double
    Dim t3 As Double
    Const RemoveQuotes As Boolean = True
    Dim WhatDiffers As String

    t1 = ElapsedTime
    res1 = CSVRead(FileName)
    t2 = ElapsedTime
    res2 = ArrayFromCSVfile(FileName, , , RemoveQuotes)
    t3 = ElapsedTime

    Debug.Print "CSVRead", t2 - t1, "ArrayFromCSVfile", t3 - t2, ArraysIdentical(res1, res2, , True, WhatDiffers)
    Debug.Print WhatDiffers

End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : TestAgainstRDatasets
' Purpose    : Test against the 20 largest files in Rdatasets https://github.com/vincentarelbundock/Rdatasets
'              To run the code, first clone the above repo to C:\Projects
' -----------------------------------------------------------------------------------------------------------------------
Sub TestAgainstRDatasets()

    Const DatasetsFolder As String = "C:\Projects\RDatasets\"
    
    Dim ResultsFile As String
    
    ResultsFile = Left$(ThisWorkbook.path, InStrRev(ThisWorkbook.path, "\")) & _
        "testresults\SpeedTestRDatasets.csv"
    
    Dim CSVResult As Variant
    Dim FileName As String
    Dim Files As Variant
    Dim i As Long
    Dim Result As Variant
    Dim sdkn104Result As Variant
    Dim t1 As Double
    Dim t2 As Double
    Dim t3 As Double
    Dim t4 As Double
    Dim ws_garciaResult As Variant

    On Error GoTo ErrHandler
    Files = RDatesetsFiles()

    Result = Fill(vbNullString, UBound(Files) - LBound(Files) + 2, 5)
    Result(1, 1) = "File Name"
    Result(1, 2) = "Size"
    Result(1, 3) = "CSVRead time"
    Result(1, 4) = "sdkn104 time"
    Result(1, 5) = "ws_garcia time"

    For i = LBound(Files) To UBound(Files)
        FileName = DatasetsFolder & Files(i)
        If Not FileExists(FileName) Then Throw "Cannot find file '" & FileName & "'"
    Next

    For i = LBound(Files) To UBound(Files)
        FileName = DatasetsFolder & Files(i)
        If Not FileExists(FileName) Then Throw "Cannot find file '" & FileName & "'"
        t1 = ElapsedTime()
        CSVResult = CSVRead(FileName)
        t2 = ElapsedTime()
        sdkn104Result = Wrap_sdkn104(FileName, False)
        t3 = ElapsedTime()
        ws_garciaResult = Wrap_ws_garcia(FileName, ",", vbCrLf)
        t4 = ElapsedTime()
        
        ThrowIfError CSVResult
        ThrowIfError ws_garciaResult
        ThrowIfError sdkn104Result
        
        Debug.Print i, t2 - t1, t3 - t2, t4 - t3
        Result(i + 2, 1) = FileName
        Result(i + 2, 2) = FileSize(FileName)
        Result(i + 2, 3) = t2 - t1
        Result(i + 2, 4) = t3 - t2
        Result(i + 2, 5) = t4 - t3

    Next

    ThrowIfError CSVWrite(Result, ResultsFile)

    Exit Sub
ErrHandler:
    MsgBox ReThrow("TestAgainstRDatasets", Err, True), vbCritical
End Sub
