Attribute VB_Name = "modCSVWritePerformance"
Option Explicit

'--------------------------------------------------------------------------------
'Time:                       27/11/2022 15:14:08
'ComputerName:               PHILIP-LAPTOP
'VersionNumber:               206
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
'--------------------------------------------------------------------------------

'--------------------------------------------------------------------------------
'Time:                       27/11/2022 15:46:11
'ComputerName:               PHILIP-LAPTOP
'VersionNumber:               208
'Elapsed time for CSVRead: 13.6756363001186 seconds
'Elapsed time for CSVWrite ANSI: 11.0888602000196 seconds
'Elapsed time for CSVWrite ANSI: 10.5916883000173 seconds
'Elapsed time for CSVWrite ANSI: 11.331025699852 seconds
'Elapsed time for CSVWrite ANSI: 11.0151514001191 seconds
'Elapsed time for CSVWrite ANSI: 11.0596908999141 seconds
'Elapsed time for CSVWrite UTF-8: 10.6685319000389 seconds
'Elapsed time for CSVWrite UTF-8: 10.5587436000351 seconds
'Elapsed time for CSVWrite UTF-8: 10.6614930001087 seconds
'Elapsed time for CSVWrite UTF-8: 10.7582542998716 seconds
'Elapsed time for CSVWrite UTF-8: 11.1152926001232 seconds
'Elapsed time for CSVWrite UTF-16: 11.0943058999255 seconds
'Elapsed time for CSVWrite UTF-16: 11.284852599958 seconds
'Elapsed time for CSVWrite UTF-16: 11.7004682000261 seconds
'Elapsed time for CSVWrite UTF-16: 11.5732845000457 seconds
'Elapsed time for CSVWrite UTF-16: 11.3086764998734 seconds
'--------------------------------------------------------------------------------

'--------------------------------------------------------------------------------
'Time:                       28/11/2022 13:38:49
'ComputerName:               DESKTOP-0VD2AF0
'VersionNumber:               212
'Elapsed time for CSVRead: 12.6903092000284 seconds
'Elapsed time for CSVWrite ANSI: 9.0200504999957 seconds
'Elapsed time for CSVWrite ANSI: 8.97261259995867 seconds
'Elapsed time for CSVWrite ANSI: 8.83159349998459 seconds
'Elapsed time for CSVWrite ANSI: 8.94957689999137 seconds
'Elapsed time for CSVWrite ANSI: 8.87800550000975 seconds
'Elapsed time for CSVWrite UTF-8: 9.05053659999976 seconds
'Elapsed time for CSVWrite UTF-8: 9.07855340000242 seconds
'Elapsed time for CSVWrite UTF-8: 8.94219269999303 seconds
'Elapsed time for CSVWrite UTF-8: 8.96188490005443 seconds
'Elapsed time for CSVWrite UTF-8: 8.95847179996781 seconds
'Elapsed time for CSVWrite UTF-16: 9.29401939996751 seconds
'Elapsed time for CSVWrite UTF-16: 9.64014640002279 seconds
'Elapsed time for CSVWrite UTF-16: 9.23568999994313 seconds
'Elapsed time for CSVWrite UTF-16: 9.18535220000194 seconds
'Elapsed time for CSVWrite UTF-16: 9.20971269998699 seconds
'--------------------------------------------------------------------------------


Sub TestWriteSpeedUsingMilitary()

          Const FileToWrite = "c:\Temp\military.csv"
          Dim Encoding As String
          Dim Res As String
          Dim i As Long, j As Long
          Dim Data
1         On Error GoTo ErrHandler
2         Debug.Print String(80, "-")
3         Debug.Print "Time:         ", Now
4         Debug.Print "ComputerName: ", Environ$("ComputerName")
5         Debug.Print "VersionNumber:", shAudit.Range("B6").value
6         tic
7         Data = ThrowIfError(CSVRead("C:\Projects\RDatasets\csv\openintro\military.csv", True))
8         toc "CSVRead"

9         For j = 1 To 3
10            Encoding = Choose(j, "ANSI", "UTF-8", "UTF-16")
11            For i = 1 To 5
12                tic
13                Res = CSVWrite(Data, FileToWrite, , , , , Encoding)
14                toc "CSVWrite " & Encoding
15            Next i
16        Next j
17        Debug.Print String(80, "-")

18        Exit Sub
ErrHandler:
19        MsgBox ReThrow("TestWriteSpeedUsingMilitary", Err, True), vbCritical
End Sub

