Attribute VB_Name = "Módulo1"
Sub primeiro()
'O comando DIM(Dimension) é utilizado para declarar variavel
'A variavél nome foi tipada como string(texto)

Dim nome As String

'O comando inputbox abre uma caixa de entrada de dados
'Assim o usuário digita o nome e aloca na variavel nome
nome = InputBox("Digite o seu nome")

'O comando range permite selecionar uma célula na planilha,
'assim selecionamos a célula A1 e adicionamos o valor que
'foi digitado na caixa de entrada usando a variavel nome
Range("A1").Value = nome

End Sub
