Sub VerificarParImpar()
    Dim oDoc As Object
    Dim oSheet As Object
    Dim oCell As object
    Dim valor As Double
    
    oDoc = ThisComponent
    oSheet = oDoc.Sheets(0)
    oCell = oDoc.CurrentSelection
    
    valor = oCell.Value
    
    If valor Mod 2 = 0 Then
    	MsgBox "El valor '" & valor & "' es PAR"
    Else
    	MsgBox "El valor '" & valor & "' es IMPAR"
    End If
End Sub