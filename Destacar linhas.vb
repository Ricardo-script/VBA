Dim LinhaSelecAnterior As Range

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

   Select Case ActiveCell.Row

       Case 1, 2, 3, 4, 5, 6, 7, 13, 15, 19, 20, 23, 24, 28, 34, 37, 41, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 59, 60, 61, 66, 74, 82, 83, 84, 85, 86, 87, 88, 89
           'Coloque neste ‘case’ as linhas que não devem ser
           'destacadas na seleção
           'Exemplo: Linhas de título; Aqui eu defini como as linhas 1 e 2

           'Remove cor de fundo da linha selecionada anteriormente
           Select Case LinhaSelecAnterior.Row

               Case Is <> 1, 2

                   Rows(LinhaSelecAnterior.Row).Interior.ColorIndex = 0

               End Select

       Case Else

           'Altera a cor de fundo da linha selecionada
           Rows(ActiveCell.Row).Interior.ColorIndex = 15

           'Remove a cor de fundo quando a linha perde a seleção
           If Not LinhaSelecAnterior Is Nothing Then

               'Verifica se a linha atual já estava selecionada
               'neste momento, caso seja uma nova linha selecionada
               'remove a cor de fundo.
               If ActiveCell.Row <> LinhaSelecAnterior.Row Then

                   Rows(LinhaSelecAnterior.Row).Interior.ColorIndex = 0

               End If

           End If

           'Inicializa a variavel informando a seleção atual
           'que será utilizada no inicio do procedimento
           'como sendo a seleção anterior
           Set LinhaSelecAnterior = ActiveCell

   End Select

End Sub