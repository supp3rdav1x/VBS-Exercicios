Dim lado, area, perimetro, resp

Call entrada_lado

Sub entrada_lado()
lado = CDbl(InputBox("Digite a medida do lado do quadrado: ", "Quadrado | In�cio"))

area = CDbl(lado * lado)
perimetro = CDbl(lado + lado + lado + lado)

MsgBox("�rea: " & area & "" & vbNewLine &_
       "Per�metro: " & perimetro & ""), vbInformation + vbOKOnly, "Quadrado | Resultado"

Call continuar
End Sub

Sub continuar()
resp = MsgBox("Deseja continuar?", vbQuestion + vbYesNo, "Quadrado | Confirma��o")

If resp = vbYes Then
    Call entrada_lado
Else
    WScript.Quit
End If

End Sub