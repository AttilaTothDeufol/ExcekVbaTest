Attribute VB_Name = "Module1"
Sub IfMacroButton()
    If Range("C2").value = "M" Then
        MsgBox "Its a Boy"
    Else
        MsgBox "Its a Girl"
    End If
End Sub
Sub CheckForFreeShipping()
    Dim Row As Integer
    Dim value As Integer
    If IsEmpty(Range("J6").value) Then
        value = InputBox("Please enter number between 2 and 13")
        Range("J6").value = value
    ElseIf Not (IsNumeric(Range("J6").value)) Or Range("J6").value > 13 Or Range("J6").value < 2 Then
        value = InputBox("Please enter number between 2 and 13")
        Range("J6").value = value
    End If
    
    Row = Range("J6").value
    If (Cells(Row, 2).value < 5000 And Cells(Row, 4).value = "Y" <> 0) Then
        MsgBox Cells(Row, 1).value & " " & Cells(Row, 2).value & " " & Cells(Row, 3).value & " " & Cells(Row, 4).value
        Cells(Row, 5).value = "Free Shipping"
    Else
        MsgBox Cells(Row, 1).value & " " & Cells(Row, 2).value & " " & Cells(Row, 3).value & " " & Cells(Row, 4).value
        Cells(Row, 5).value = "Charged Shipping"
    End If
    
   
End Sub
