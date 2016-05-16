'Entrada do texto cifrado pelo usuario
strTexto = InputBox("Entre com o texto:", "Decifrador de Cesar")

'Entrada do usuario da quantidade de rotacoes(pra esquerda) que serao dadas para decifrar o texto
numRotacoes = InputBox("Quantidade de Rotacoes a esquerda para decifrar?:", "Rotacoes")

'Texto decriptado
decriptado = ""

'Pega cada letra do texto....
For i=1 To Len(strTexto)
	'Captura a letra
	c = Mid(strTexto,i,1)

	'Converte ela inteiro (ASCII)
	numChar = Asc(c)

	'Se ela for minuscula...
	If numChar>=65 And numChar<=90 Then
		numChar = numChar - numRotacoes
		'Tratamento do Underflow
		If numChar<65 Then
			numChar = numChar + 26
		End If
	ElseIf Asc(c)>=97 And Asc(c)<=122 Then  'Se ela for maiuscula...
    	numChar = numChar - numRotacoes
		'Tratamento do Underflow
		If numChar<97 Then
			numChar = numChar + 26
		End If
	End If

	'WScript.Echo Chr(numChar)
	'Concatena a letra decriptada no resultado
	decriptado = decriptado & CStr(Chr(numChar))
Next

'Exibe o resultado para o Usuario
WScript.Echo decriptado
