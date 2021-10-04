Dim v1, v2, v3, maior, meio, menor, resp

Call entrada_valores

Sub entrada_valores ()
v1 = CDbl(InputBox("Valor 1: ", "Maior Valor | Início"))
v2 = CDbl(InputBox("Valor 2: ", "Maior Valor"))
v3 = CDbl(InputBox("Valor 3: ", "Maior Valor"))

' Maior que v3
If v3 > v1 and v3 > v2 Then
    MsgBox("O maior valor é: " & "" & v3 & ""), vbInformation + vbOKOnly, "Maior Valor | Resultado"

' Maior que v2
ElseIf v2 > v1 and v2 > v3 Then
    MsgBox("O maior valor é: " & "" & v2 & ""), vbInformation + vbOKOnly, "Maior Valor | Resultado"

' Maior que v1
ElseIf v1 > v2 and v1 > v3 Then
    MsgBox("O maior valor é: " & "" & v1 & ""), vbInformation + vbOKOnly, "Maior Valor | Resultado"

' Todos iguais
ElseIf v1 = v2 and v2 = v3 Then
    MsgBox("Todos tem o mesmo valor!" & ""), vbInformation + vbOKOnly, "Maior Valor | Resultado"

' Erros
Else
    MsgBox("Erro!" & ""), vbExclamation + vbOKOnly, "Maior Valor | Erro"
End If

Call continuar
End Sub

Sub continuar()
resp = MsgBox("Deseja continuar?", vbQuestion + vbYesNo, "Maior Valor | Confirmação")
If resp = vbYes Then
    Call entrada_valores
Else
    WScript.Quit
End If
End Sub