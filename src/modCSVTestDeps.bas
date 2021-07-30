Attribute VB_Name = "modCSVTestDeps"
'Functions in this module are called from modCSVTest, but not called from modCSV

Option Explicit
Option Private Module


#If VBA7 Then
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
#Else
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
#End If

'---------------------------------------------------------------------------------------------------------
' Procedure : CreatePath
' Author    : Philip Swannell
' Date      : 29-Jun-2018
' Purpose   : Creates a folder on disk. FolderPath can be passed in as C:\This\That\TheOther even if the
'             folder C:\This does not yet exist. If successful returns the name of the
'             folder. If not successful returns an error string.
' Arguments
' FolderPath: Path of the folder to be created. For example C:\temp\My_New_Folder. It does not matter if
'             this path has a terminating backslash or not. This argument may be an array.
'---------------------------------------------------------------------------------------------------------
Function CreatePath(ByVal FolderPath As String)

          Dim F As Scripting.Folder
          Dim FSO As Scripting.FileSystemObject
          Dim i As Long
          Dim isUNC As Boolean
          Dim ParentFolderName

1         On Error GoTo ErrHandler

2         If Left$(FolderPath, 2) = "\\" Then
3             isUNC = True
4         ElseIf Mid$(FolderPath, 2, 2) <> ":\" Or Asc(UCase$(Left$(FolderPath, 1))) < 65 Or Asc(UCase$(Left$(FolderPath, 1))) > 90 Then
5             Throw "First three characters of FolderPath must give drive letter followed by "":\"" or else be""\\"" for " & _
                  "UNC folder name"
6         End If

7         FolderPath = Replace(FolderPath, "/", "\")

8         If Right$(FolderPath, 1) <> "\" Then
9             FolderPath = FolderPath + "\"
10        End If

11        Set FSO = New FileSystemObject
12        If FolderExists(FolderPath) Then
13            GoTo EarlyExit
14        End If

          'Go back until we find parent folder that does exist
15        For i = Len(FolderPath) - 1 To 3 Step -1
16            If Mid$(FolderPath, i, 1) = "\" Then
17                If FolderExists(Left$(FolderPath, i)) Then
18                    Set F = FSO.GetFolder(Left$(FolderPath, i))
19                    ParentFolderName = Left$(FolderPath, i)
20                    Exit For
21                End If
22            End If
23        Next i

24        If F Is Nothing Then Throw "Cannot create folder " + Left$(FolderPath, 3)

          'now add folders one level at a time
25        For i = Len(ParentFolderName) + 1 To Len(FolderPath)
26            If Mid$(FolderPath, i, 1) = "\" Then
                  Dim ThisFolderName As String
27                ThisFolderName = Mid$(FolderPath, InStrRev(FolderPath, "\", i - 1) + 1, i - 1 - InStrRev(FolderPath, "\", i - 1))
28                F.SubFolders.Add ThisFolderName
29                Set F = FSO.GetFolder(Left$(FolderPath, i))
30            End If
31        Next i

EarlyExit:
32        Set F = FSO.GetFolder(FolderPath)
33        CreatePath = F.Path
34        Set F = Nothing: Set FSO = Nothing

35        Exit Function
ErrHandler:
36        CreatePath = "#CreatePath (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
'---------------------------------------------------------------------------------------
' Procedure : FolderExists
' Author    : Philip Swannell
' Date      : 07-Oct-2013
' Purpose   : Returns True or False. Does not matter if FolderPath has a terminating backslash or not.
'---------------------------------------------------------------------------------------
Function FolderExists(ByVal FolderPath As String)
          Dim F As Scripting.Folder
          Dim FSO As Scripting.FileSystemObject
1         On Error GoTo ErrHandler
2         Set FSO = New FileSystemObject
3         Set F = FSO.GetFolder(FolderPath)
4         FolderExists = True
5         Exit Function
ErrHandler:
6         FolderExists = False
End Function
'
'---------------------------------------------------------------------------------------------------------
' Procedure : sElapsedTime
' Author    : Philip Swannell
' Date      : 16-Jun-2013
' Purpose   : Retrieves the current value of the performance counter, which is a high resolution (<1us)
'             time stamp that can be used for time-interval measurements.
'
'             See http://msdn.microsoft.com/en-us/library/windows/desktop/ms644904(v=vs.85).aspx
'---------------------------------------------------------------------------------------------------------
Function sElapsedTime() As Double
          Dim a As Currency
          Dim b As Currency
1         On Error GoTo ErrHandler

2         QueryPerformanceCounter a
3         QueryPerformanceFrequency b
4         sElapsedTime = a / b

5         Exit Function
ErrHandler:
6         Throw "#sElapsedTime (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : sNCols
' Author    : Philip Swannell
' Date      : 19-Jun-2013
' Purpose   : Number of columns in an array. Missing has zero rows, 1-dimensional arrays
'             have one row and the number of columns returned by this function.
'---------------------------------------------------------------------------------------
Function sNCols(Optional TheArray) As Long
1         If TypeName(TheArray) = "Range" Then
2             sNCols = TheArray.Columns.Count
3         ElseIf IsMissing(TheArray) Then
4             sNCols = 0
5         ElseIf VarType(TheArray) < vbArray Then
6             sNCols = 1
7         Else
8             Select Case NumDimensions(TheArray)
                  Case 1
9                     sNCols = UBound(TheArray, 1) - LBound(TheArray, 1) + 1
10                Case Else
11                    sNCols = UBound(TheArray, 2) - LBound(TheArray, 2) + 1
12            End Select
13        End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : sNRows
' Author    : Philip Swannell
' Date      : 19-Jun-2013
' Purpose   : Number of rows in an array. Missing has zero rows, 1-dimensional arrays have one row.
'---------------------------------------------------------------------------------------
Function sNRows(Optional TheArray) As Long
1         If TypeName(TheArray) = "Range" Then
2             sNRows = TheArray.Rows.Count
3         ElseIf IsMissing(TheArray) Then
4             sNRows = 0
5         ElseIf VarType(TheArray) < vbArray Then
6             sNRows = 1
7         Else
8             Select Case NumDimensions(TheArray)
                  Case 1
9                     sNRows = 1
10                Case Else
11                    sNRows = UBound(TheArray, 1) - LBound(TheArray, 1) + 1
12            End Select
13        End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : sFill
' Author     : Philip Swannell
' Date       : 30-Jul-2021
' Purpose    : Creates an array filled with the value x
' -----------------------------------------------------------------------------------------------------------------------
Function sFill(ByVal x As Variant, ByVal NumRows As Long, ByVal NumCols As Long)

1         On Error GoTo ErrHandler

          Dim i As Long
          Dim j As Long
          Dim Result() As Variant

2         ReDim Result(1 To NumRows, 1 To NumCols)

3         For i = 1 To NumRows
4             For j = 1 To NumCols
5                 Result(i, j) = x
6             Next j
7         Next i

8         sFill = Result

9         Exit Function
ErrHandler:
10        sFill = "#sFill (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
