Sub ConvertirBases()
    Dim valor As Variant
    Dim binario As String
    Dim octal As String
    Dim hexadecimal As String
    Dim decimal As Double
    Dim celda As Object

    Set celda = ThisComponent.CurrentSelection
    
    If IsNumeric(celda.Value) Then
        valor = celda.Value
        decimal = valor
      
        binario = ConvertirABase(decimal, 2)
        octal = ConvertirABase(decimal, 8)
        hexadecimal = ConvertirABase(decimal, 16)
        
        ThisComponent.Sheets(0).getCellByPosition(1, 2).String = binario
        ThisComponent.Sheets(0).getCellByPosition(2, 2).String = octal
        ThisComponent.Sheets(0).getCellByPosition(3, 2).String = hexadecimal
        ThisComponent.Sheets(0).getCellByPosition(4, 2).String = decimal
    Else
        MsgBox "Por favor ingresa un valor numérico."
    End If
End Sub

Function ConvertirABase(decimal, basenum) As String
    Dim resultado As String
    Dim residuo As Integer
    Dim digitos As String
    Dim decifuncion As Integer
    Dim digito As String

    digitos = "0123456789ABCDEF"
    resultado = ""

    If basenum < 2 Or basenum > 16 Then
        ConvertirABase = "Base no válida"
        Exit Function
    End If

    decifuncion = decimal

    Do While decifuncion > 0
        residuo = decifuncion Mod basenum
        decifuncion = Int(decifuncion / basenum) 
        digito = Mid(digitos, residuo + 1, 1)
        resultado = digito & resultado
    Loop

    If resultado = "" Then
        resultado = "0"
    End If
    
    ConvertirABase = resultado
    
End Function