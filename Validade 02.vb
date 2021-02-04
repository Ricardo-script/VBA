Experiação de Planilha exemplo  :

Private Sub Workbook_Open()
Dim cont, auxcont As Integer
Dim senha, enter_snh As String
  senha = "1234"
  Range("c10000") = Date
  If Range("b10000") = nill Then
  Range("b10000") = Date
  End If
  cont = Format(Range("c10000") - Range("b10000"), "0")
  auxcont = 30 - cont
  If cont >= 30 Then
  enter_snh = InputBox("ESSE PROGRAMA EXPIROU, INFORME A SENHA DE ACESSO:", "Informe...")
  If enter_snh <> senha Then
  MsgBox "Senha incorreta!", vbInformation, "Erro"
  Saved = True
  ActiveWorkbook.Close
  End If
  Else
  MsgBox "Faltam " & auxcont & " dias  para expirar."
  End If
  
End Sub

Exemplo 2:


Sub workbook_open()
If Date <= #10/11/2014# Then Exit Sub
MsgBox "Planilha fora da validade"
With ThisWorkbook
    .Saved = True
    .ChangeFileAccess xlReadOnly
    Kill .FullName
    .Close False
End With
End Sub
