Private Sub btn_Cancel_Click()
Unload Me
End Sub

Private Sub btn_OK_Click()

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



Cells(4, 12).Value = TextBox1.Value
Cells(7, 12).Value = TextBox2.Value
Unload Me
End Sub