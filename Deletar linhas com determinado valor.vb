Sub DeleteRowsWithWord()
 Dim Col As Variant, Word As String
 On Error GoTo fim

 'Let Col = InputBox("Em qual coluna devo manter o foco da busca da palavra?")
 Let Col = ("E")

 If Len(Col) > 0 And Not Col Like "*[!0-9]*" Then Col = Val(Col)

 'Let Word = InputBox("Que palavra devo encontrar nas Linhas para apagá-las?")
 Let Word = "Em deposito"

 With Columns(Col)
     .Replace Word, "#N/A", xlWhole
     .SpecialCells(xlCellTypeConstants, xlErrors).EntireRow.Delete
 End With
fim:
End Sub