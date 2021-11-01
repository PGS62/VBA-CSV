Attribute VB_Name = "Module1"
Option Explicit

Sub MakeGIF()
    Dim Keys
    Dim i As Long
    Dim t1 As Double

    AppActivate "VBA-CSV.xlsm"

    ActiveSheet.Calculate
    Keys = sExpandDown(ActiveSheet.Range("keys"))

    For i = 1 To sNRows(Keys)

        Application.SendKeys Keys(i, 1)
        t1 = sElapsedTime()
        While sElapsedTime() < t1 + ActiveSheet.Range("Delays").Offset(i - 1).Value
            DoEvents
        Wend
    Next i

Debug.Print Now()


End Sub


Sub MakeFirstGIF()

CopyStringToClipboard "https://vincentarelbundock.github.io/Rdatasets/csv/carData/TitanicSurvival.csv"
MakeGIF

End Sub
