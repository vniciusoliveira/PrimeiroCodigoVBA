Attribute VB_Name = "M�dulo1"
Sub primeiro()
'O comando DIM(Dimension) � utilizado para declarar variavel
'A variav�l nome foi tipada como string(texto)

Dim nome As String

'O comando inputbox abre uma caixa de entrada de dados
'Assim o usu�rio digita o nome e aloca na variavel nome
nome = InputBox("Digite o seu nome")

'O comando range permite selecionar uma c�lula na planilha,
'assim selecionamos a c�lula A1 e adicionamos o valor que
'foi digitado na caixa de entrada usando a variavel nome
Range("A1").Value = nome

End Sub
