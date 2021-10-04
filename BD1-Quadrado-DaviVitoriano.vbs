Dim lado, area, perimetro, resp

Call entrada_lado

Sub entrada_lado()
lado = CDbl(InputBox("Digite a medida do lado do quadrado: ", "Quadrado | Início"))

area = CDbl(lado * lado)
perimetro = CDbl(lado + lado + lado + lado)

MsgBox("Área: " & area & "" & vbNewLine &_
       "Perí­metro: " & perimetro & ""), vbInformation + vbOKOnly, "Quadrado | Resultado"

Call continuar
End Sub

Sub continuar()
resp = MsgBox("Deseja continuar?", vbQuestion + vbYesNo, "Quadrado | Confirmação")

If resp = vbYes Then
    Call entrada_lado
Else
    WScript.Quit
End If

End Sub