Private Sub Workbook_Open()
Dim cont, auxcont As Integer
Dim senha, enter_snh As String
  senha = "1234"
  Sheets("Validade").Select
  Range("c1") = Date ' data atual
  If Range("b1") = nill Then ' inserir na célula a data de validade
  Range("b1") = Date
  End If
  cont = Format(Range("c1") - Range("b1"), "0") ' data atual menos data de validade
  auxcont = 0 - cont
  If cont >= 0 Then
  enter_snh = InputBox("ESSE PROGRAMA EXPIROU, INFORME A SENHA DE ACESSO:", "Informe...")
  If enter_snh <> senha Then
  'MsgBox "Senha incorreta!", vbInformation, "Erro"
  
 
 ThisWorkbook.Saved = True
  
         
     MsgBox "Este arquivo está expirado, se autoexcluirá!"
  
     ThisWorkbook.ChangeFileAccess xlReadOnly
  
     Kill ThisWorkbook.FullName
     ThisWorkbook.Close


  End If
  Else
  Sheets("Du Pont").Select
  MsgBox "Faltam " & auxcont & " dias  para expirar."
  End If
  
End Sub

