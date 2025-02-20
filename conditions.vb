Module VBModule
    Sub Main()
        If [Condition] Then
            [Code to Execute if Condition is True]
        Else
            [Code to Execute if Condition is False]
            End If
    End Sub
    
    'Checking a Cell’s Value'
    Sub CheckCellValue()
        If Range("A1").Value > 10 Then
            MsgBox "The value in A1 is greater than 10."
        Else
            MsgBox "The value in A1 is 10 or less."
        End If
    End Sub
    
    'Nested If Statements'
    Sub NestedIfExample()
        If Range("A1").Value > 20 Then
            MsgBox "The value in A1 is greater than 20."
        ElseIf Range("A1").Value > 10 Then
            MsgBox "The value in A1 is between 11 and 20."
        Else
            MsgBox "The value in A1 is 10 or less."
        End If
    End Sub
    
    'Looping Trough A range of Cells'
    Sub ForNextExample()
        Dim i As Integer
        For i = 1 To 10
            Cells(i, 1).Value = i 'Fills cells A1 to A10 with values 1 to 10
        Next i
    End Sub
    
    'Looping Until a Condition is False'
    Sub DoWhileExample()
        Dim i As Integer
        i = 1
        Do While Cells(i, 1).Value <> ""
            Cells(i, 2).Value = Cells(i, 1).Value * 2 'Doubles the value in column A and places it in column B
            i = i + 1
        Loop
    End Sub

    'Looping Until a Cell is Empty'
    Sub DoUntilExample()
        Dim i As Integer
        i = 1
        Do Until Cells(i, 1).Value = ""
            Cells(i, 2).Value = Cells(i, 1).Value & " Processed" 'Appends "Processed" to the value in column A
            i = i + 1
        Loop
    End Sub
    
End Module
