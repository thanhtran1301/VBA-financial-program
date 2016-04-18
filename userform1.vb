rivate Sub btn_Cancel_Click()
Unload Me
End Sub

Private Sub btn_OK_Click()
'Get the input for the spot price,
' volatility and risk-free rate

If TextBox1.Value = "" Or IsNumeric(TextBox1.Value) = False Then
MsgBox "Please enter a valid number"
TextBox1.SetFocus
Exit Sub
End If

If TextBox2.Value = "" Or IsNumeric(TextBox2.Value) = False Then
TextBox2.SetFocus
MsgBox "Please enter a valid number"
Exit Sub
End If

If TextBox3.Value = "" Or IsNumeric(TextBox3.Value) = False Then
TextBox3.SetFocus
MsgBox "Please enter a valid number"
Exit Sub
End If


Unload Me

Cells(4, 2).Value = TextBox1.Value
Cells(5, 2).Value = TextBox2.Value
Cells(9, 2).Value = TextBox3.Value

End Sub