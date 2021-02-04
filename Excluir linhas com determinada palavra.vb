Sub DeleteRowsWithWord()
Dim Col As Variant, Word As String

Let Col = InputBox("Em qual coluna devo manter o foco da busca da palavra?")

If Len(Col) > 0 And Not Col Like "*[!0-9]*" Then Col = Val(Col)

Let Word = InputBox("Que palavra devo encontrar nas Linhas para apagá-las?")

With Columns(Col)
    .Replace Word, "#N/A", xlWhole
    .SpecialCells(xlCellTypeConstants, xlErrors).EntireRow.Delete
End With
End Sub

