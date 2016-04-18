Private Sub btn_Cancel_Click()
Unload Me

End Sub

Private Sub btn_OK_Click()
If TextBox1.Value = "" Or IsNumeric(TextBox1.Value) = False Then
MsgBox "Please enter a valid number"
TextBox1.SetFocus
Exit Sub
End If
Cells(1, 2).Value = TextBox1.Value
Unload Me

End Sub