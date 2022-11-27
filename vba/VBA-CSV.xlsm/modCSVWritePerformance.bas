Attribute VB_Name = "modCSVWritePerformance"
Option Explicit


'27/11/2022 15:14:08
'VersionNumber 206
'Elapsed time for CSVRead: 17.5810412999708 seconds
'Elapsed time for CSVWrite ANSI: 25.1498529999517 seconds
'Elapsed time for CSVWrite ANSI: 26.987809299957 seconds
'Elapsed time for CSVWrite ANSI: 27.0266166999936 seconds
'Elapsed time for CSVWrite ANSI: 27.2726858998649 seconds
'Elapsed time for CSVWrite ANSI: 26.7385879000649 seconds
'Elapsed time for CSVWrite UTF-8: 20.1281892000698 seconds
'Elapsed time for CSVWrite UTF-8: 17.2521635999437 seconds
'Elapsed time for CSVWrite UTF-8: 19.8307610999327 seconds
'Elapsed time for CSVWrite UTF-8: 16.9916775000747 seconds
'Elapsed time for CSVWrite UTF-8: 16.8933107000776 seconds
'Elapsed time for CSVWrite UTF-16: 27.9302846000064 seconds
'Elapsed time for CSVWrite UTF-16: 27.2358552000951 seconds
'Elapsed time for CSVWrite UTF-16: 26.9824278000742 seconds
'Elapsed time for CSVWrite UTF-16: 28.2433342998847 seconds
'Elapsed time for CSVWrite UTF-16: 27.3908504000865 seconds


Sub TestWriteSpeedUsingMilitary()

          Const FileToWrite = "c:\Temp\military.csv"
          Dim Encoding As String
          Dim Res As String
          Dim i As Long, j As Long
          Dim Data
          
          Debug.Print Now
          Debug.Print "VersionNumber", shAudit.Range("B6").value
1         tic
2         Data = CSVRead("C:\Projects\RDatasets\csv\openintro\military.csv", True)
3         toc "CSVRead"

4         For j = 1 To 3
5             Encoding = Choose(j, "ANSI", "UTF-8", "UTF-16")
6             For i = 1 To 5
7                 tic
8                 Res = CSVWrite(Data, FileToWrite, , , , , Encoding)
9                 toc "CSVWrite " & Encoding
10            Next i
11        Next j
End Sub

