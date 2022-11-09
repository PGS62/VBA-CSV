Attribute VB_Name = "modAdHoc2"
Option Explicit

Function TestSpeed()
          Dim t1 As Double, t2 As Double, N As Long, Res As Variant, i As Long

          Const FileName = "C:\Projects\VBA-CSV\testfiles\single_character.csv"

1         N = 10000

2         t1 = ElapsedTime()

3         For i = 1 To N
              '~3.8E-04
4             Res = CSVRead(FileName, False, ",", , , , , , , , , , , , , , "ANSI")
              
              '~6.5E-04
              'res = CSVRead(FileName, False)
5         Next i
6         t2 = ElapsedTime()

7         TestSpeed = (t2 - t1) / N
8         Debug.Print TestSpeed

End Function



Function Foo()
1         On Error GoTo ErrHandler
2         Foo = Bar()
3         Exit Function
ErrHandler:
4         Foo = ReThrow("Foo", Err, True)
End Function

Private Function Bar()
1         On Error GoTo ErrHandler
2         Bar = Baz()
3         Exit Function
ErrHandler:
4         ReThrow "Bar", Err
End Function

Private Function Baz()
1         On Error GoTo ErrHandler
2         Baz = Grumpy()
3         Exit Function
ErrHandler:
4         ReThrow "Baz", Err
End Function

Private Function Grumpy()
1         On Error GoTo ErrHandler
2         Throw "WTF"
3         Exit Function
ErrHandler:
4         ReThrow "Grumpy", Err
End Function



