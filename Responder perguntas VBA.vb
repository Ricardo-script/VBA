Private Sub CommandButton1_Click()
Sheets("Plan1").Select
Range("A1").Select
ActiveCell.Offset(1, 0).Select

If TextBox1.Value <> "" Then
    ActiveCell.Value = TextBox1.Value
End If

Call Procura
Procurar.Hide
End Sub
Sub Procura()

With Worksheets(1).Range("a2:a500")
Set c = .Find(what:="KPI", LookIn:=xlValues)
If Not c Is Nothing Then
MsgBox " O Relatório de Kpi se encontra na pasta de rede, Basf é gerado com 30 dias e FMC do dia 20 do dia ante-anterior até o dia 21 do mês anterior!"
End If

Set c = .Find(what:="Provisão", LookIn:=xlValues)
If Not c Is Nothing Then
MsgBox " O Relatório de Provisão deve se inserir os numeros de Fatura e numeros de chave, pegar o modelo na pasta de rede!"

End If

End With
'MsgBox " Não Encontrado"
'Call GeraMacro
End Sub

