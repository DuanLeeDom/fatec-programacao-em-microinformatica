Attribute VB_Name = "Módulo1"
Sub aula1vba()
    MsgBox "Oi, tudo bem com vocês!!!"
End Sub

Sub mensagem()
    MsgBox "Qual é o seu nome: " & Range("C2").Value
End Sub

Sub limpar()
    If MsgBox("Você realmente deseja limpar os dados da planilha?", vbYesNo + vbQuestion, "Limpar Planilha") = vbYes Then
        Range("C2").ClearContents
    End If
End Sub

Sub mensagem3()

    Nome = InputBox("Por favor, digite seu nome.")
    MsgBox "Muito Obrigado, " & Nome & "!"

End Sub

Sub mensagem4()

    Nome = InputBox("Por favor, digite seu nome.")
    MsgBox "Muito Obrigado, " & Nome & "!"

End Sub

Sub soma()

    Dim numero
    Dim resultado
    numeros = Worksheets("Planilha1").Range("B8", "B20")
    resultado = WorksheetFunction.Sum(numeros)
    Range("C9") = resultado
    
End Sub

Sub soma2()

    Dim numero As Variant
    Dim resultado As Double
    numero = Worksheets("planilha1").Range("B8:B20")
    resultado = WorksheetFunction.Sum(numero)
    Range("C9").Value = resultado

End Sub

Sub somas()

    Dim Numeros1
    Dim Numeros2
    Dim Numeros3
    
    Dim Resultado1
    Dim Resultado2
    Dim Resultado3
    
    Numeros1 = Worksheets("Planilha2").Range("A1", "A10")
    Numeros2 = Worksheets("Planilha2").Range("B1", "B10")
    Numeros3 = Worksheets("Planilha2").Range("C1", "C10")
    
    Resultado1 = WorksheetFunction.Sum(Numeros1)
    Resultado2 = WorksheetFunction.Sum(Numeros2)
    Resultado3 = WorksheetFunction.Sum(Numeros3)

    Range("E1") = Resultado1 + Resultado2 + Resultado3

End Sub

Sub calculos()

    Dim num1 As Integer
    Dim num2 As Integer
    Dim resposta1, resposta2, resposta3, resposta4 As String
    
    num1 = Range("B1").Value
    num2 = Range("B2").Value
    
    Range("D1") = num1 + num2
    Range("D2") = num1 - num2
    Range("D3") = num1 * num2
    Range("D4") = num1 / num2

End Sub

