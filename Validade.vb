‘PLANILHA É EXCLUÍDA ANTES DE ABRIR
‘Pois bem, vamos ao código: Abra a janela de edição (atalho alt + f11) e dê um ‘clique em 'EstaPasta_de_trabalho'. Lá vamos colar o seguinte código:


Private Sub Workbook_Open()
Dim dtexp As Date
  
          'Escolha a data que deverá expirar
     dtexp = ("29/04/2011")
  
     If Date >= #1/11/2010# Then
     If Date >= dtexp Then
  
     ThisWorkbook.Saved = True
  
          'personalize a mensagem na linha abaixo
     MsgBox "Este arquivo está expirado, se autoexcluirá!"         
  
     ThisWorkbook.ChangeFileAccess xlReadOnly
  
     Kill ThisWorkbook.FullName
     ThisWorkbook.Close
          
     End If
     End If
  
End Sub

