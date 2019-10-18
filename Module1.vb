Module Module1
   

   
    
    Function MNI(ByVal file As String, ByRef Less As Double, ByRef Discount As Double, ByRef Sales As Double) As Boolean
        MNI = True
        Less = 0
        Discount = 0
        Sales = 0
        Try
            FileOpen(2, file, OpenMode.Input, OpenAccess.Read)
            Dim reder As String = ""
            While EOF(2) = False
                Input(2, reder)
                Input(2, reder)
                Input(2, reder)
                Input(2, reder)
                Input(2, reder)
                Input(2, reder)
                Input(2, reder)
                Input(2, reder)
                Input(2, reder)
                Input(2, reder)
                Input(2, reder) 'less
                Less += Val(Replace(asc100(reder), ",", ""))
                Input(2, reder) 'disc
                Discount += Val(Replace(asc100(reder), ",", ""))
                Input(2, reder) 'coh
                Sales += Val(Replace(asc100(reder), ",", ""))
            End While
            FileClose(2)
        Catch ex As Exception
            MNI = False
        End Try

    End Function
    Function asc100(ByVal str As String) As String
        asc100 = ""
        Dim i As Integer
        For i = 0 To str.Length - 1
            asc100 = asc100 & Chr(Asc(str.Chars(i)) - 100)
        Next i
    End Function
    

End Module
