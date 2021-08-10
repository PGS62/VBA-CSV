Attribute VB_Name = "modParsingWithStreaming"
Option Explicit


Function RawFileContents(FileName As String)
            Dim FSO As New FileSystemObject, F As Scripting.File, T As Scripting.TextStream
            Set F = FSO.GetFile(FileName)
            Set T = F.OpenAsTextStream()
            RawFileContents = T.ReadAll
            T.Close


End Function


'Sub TestSearchInBuffer()
'          Dim T As scripting.TextStream
'          Dim FSO As New scripting.FileSystemObject
'          Dim FileName As String
'          Dim res1 As Long
'          Dim res2
'          Dim i As Long
'
'          Const Delimiter = ":::"
'          Const QuoteChar = """"
'          Dim Buffer As String
'          Dim BufferUpdatedTo As Long
'          Dim Streaming As Boolean
'          Dim ReadAll As String
'          Dim SearchFor As String
'          Dim SearchForArray() As String
'          Dim Which As Long
'
'          Const startover = True
'
'1         On Error GoTo ErrHandler
'2         FileName = "c:\temp\anothertest.csv"
'
'3         Set T = FSO.GetFile(FileName).OpenAsTextStream(ForReading)
'4         ReadAll = T.ReadAll
'5         T.Close
'
'6         ReDim SearchForArray(1 To 1)
'
'7         Set T = FSO.GetFile(FileName).OpenAsTextStream(ForReading)
'
'8         For i = 30 To 100
'9             SearchFor = Chr(i)
'10            SearchForArray(1) = SearchFor
'11            res1 = SearchInBuffer(SearchForArray, 1, T, Delimiter, QuoteChar, Which, Buffer, BufferUpdatedTo)
'12            res2 = InStr(ReadAll, SearchFor)
'13            If res2 = 0 Then res2 = Len(ReadAll) + 1
'
'14            Debug.Print SearchFor, res1, res2, res1 = res2
'15            If res1 <> res2 Then Stop
'
'16            If startover Then
'17                T.Close
'18                Set T = FSO.GetFile(FileName).OpenAsTextStream(ForReading)
'19                Buffer = ""
'20                BufferUpdatedTo = 0
'
'21            End If
'
'
'22        Next i
'23        T.Close
'24        Debug.Print String(50, "=")
'25        Debug.Print ReadAll
'26        Debug.Print (Len(ReadAll))
'
'
'27        Exit Sub
'ErrHandler:
'28        SomethingWentWrong "#TestSearchInBuffer (line " & CStr(Erl) + "): " & Err.Description & "!"
'End Sub
'
'
'
'Sub TestInStrMulti()
'          Dim SearchFor() As String
'          Const SearchWithin = "????????????This?????That???????????Another"
'          Dim StartingAt As Long
'          Dim EndingAt As Long
'          Dim Which As Long
'          Dim Res As Long
'
'1         StartingAt = 14
'2         EndingAt = Len(SearchWithin)
'
'
'3         ReDim SearchFor(1 To 3)
'
'4         SearchFor(1) = "This"
'5         SearchFor(2) = "That"
'6         SearchFor(3) = "Another"
'
'
'7         Res = InStrMulti(SearchFor, SearchWithin, StartingAt, EndingAt, Which)
'
'8         Debug.Print Res
'9         Debug.Print Which
'
'
'End Sub






















