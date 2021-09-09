Attribute VB_Name = "modCSVTestRDatasets"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : TestAgainstRDatasets
' Purpose    : Test against the 20 largest files in Rdatasets https://github.com/vincentarelbundock/Rdatasets
'              To run the code, first clone the above repo to C:\Projects
' -----------------------------------------------------------------------------------------------------------------------
Sub TestAgainstRDatasets()

    Const DatasetsFolder = "C:\Projects\RDatasets\"
    
    Dim ResultsFile As String
    
    ResultsFile = Left$(ThisWorkbook.path, InStrRev(ThisWorkbook.path, "\")) + _
        "testresults\SpeedTestRDatasets.csv"
    
    Dim CSVResult
    Dim FileName As String
    Dim Files
    Dim i As Long
    Dim Result
    Dim sdkn104Result
    Dim t1 As Double
    Dim t2 As Double
    Dim t3 As Double
    Dim t4 As Double
    Dim ws_garciaResult

    On Error GoTo ErrHandler
    Files = Array("csv\openintro\military.csv", "csv\mosaicData\Birthdays.csv", _
        "csv\stevedata\wvs_justifbribe.csv", "csv\nycflights13\flights.csv", _
        "csv\stevedata\wvs_immig.csv", "csv\AER\Fertility.csv", _
        "csv\openintro\avandia.csv", "csv\Stat2Data\AthleteGrad.csv", _
        "csv\causaldata\mortgages.csv", "csv\openintro\mammogram.csv", _
        "csv\lme4\InstEval.csv", "csv\stevedata\gss_abortion.csv", _
        "csv\stevedata\TV16.csv", "csv\stevedata\gss_wages.csv", _
        "csv\AER\CPSSW8.csv", "csv\stevedata\eq_passengercars.csv", _
        "csv\ggplot2movies\movies.csv", "csv\ggplot2\diamonds.csv", _
        "csv\causaldata\gov_transfers_density.csv", "csv\openintro\seattlepets.csv")

    Result = Fill("", UBound(Files) - LBound(Files) + 2, 5)
    Result(1, 1) = "File Name"
    Result(1, 2) = "Size"
    Result(1, 3) = "CSVRead time"
    Result(1, 4) = "sdkn104 time"
    Result(1, 5) = "ws_garcia time"

    For i = LBound(Files) To UBound(Files)
        FileName = DatasetsFolder & Files(i)
        If Not FileExists(FileName) Then Throw "Cannot find file '" + FileName + "'"
    Next

    For i = LBound(Files) To UBound(Files)
        FileName = DatasetsFolder & Files(i)
        If Not FileExists(FileName) Then Throw "Cannot find file '" + FileName + "'"
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
    MsgBox "#TestAgainstRDatasets (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

