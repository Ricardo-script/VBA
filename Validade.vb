�PLANILHA � EXCLU�DA ANTES DE ABRIR
�Pois bem, vamos ao c�digo: Abra a janela de edi��o (atalho alt + f11) e d� um �clique em 'EstaPasta_de_trabalho'. L� vamos colar o seguinte c�digo:


Private Sub Workbook_Open()
Dim dtexp As Date
  
          'Escolha a data que dever� expirar
     dtexp = ("29/04/2011")
  
     If Date >= #1/11/2010# Then
     If Date >= dtexp Then
  
     ThisWorkbook.Saved = True
  
          'personalize a mensagem na linha abaixo
     MsgBox "Este arquivo est� expirado, se autoexcluir�!"         
  
     ThisWorkbook.ChangeFileAccess xlReadOnly
  
     Kill ThisWorkbook.FullName
     ThisWorkbook.Close
          
     End If
     End If
  
End Sub

