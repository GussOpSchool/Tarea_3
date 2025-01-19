Sub VerificarMayorIgualMenor()
    Dim oDoc As Object
    Dim oSheet As Object
    Dim oCell As object
    Dim valor As Double
    
    oDoc = ThisComponent
    oSheet = oDoc.Sheets(0)
    oCell = oDoc.CurrentSelection
    
    valor = oCell.Value
    
    If valor > 10 Then
        MsgBox "El valor es mayor que 10"
    ElseIf valor = 10 Then
        MsgBox "El valor es igual a 10"
    ElseIf valor < 10 Then
        MsgBox "El valor es menor que 10"
    End If
End Sub