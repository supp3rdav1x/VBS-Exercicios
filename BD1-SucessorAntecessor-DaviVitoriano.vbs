Dim numero, antecessor, sucessor, resp

Call entrada_numero

Sub entrada_numero()
numero = CDbl(InputBox("Digite um número: ", "Antecessor e Sucessor | Início"))

antecessor = CDbl(numero - 1)
sucessor = CDbl(numero + 1)

MsgBox("Antecessor: " & antecessor & "" & vbnewline &_
       "Sucessor: " & sucessor & ""), vbInformation + vbOKOnly, "Antecessor e Sucessor | Resultado"

Call continuar
End Sub

Sub continuar()
resp = MsgBox("Deseja continuar?", vbQuestion + vbYesNo, "Antecessor e Sucessor | Confirmação")

If resp = vbYes Then
    Call entrada_numero
Else
    WScript.Quit
End If

End Sub 