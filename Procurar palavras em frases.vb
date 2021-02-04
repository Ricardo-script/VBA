Sub ProcurarOcorrencia()

    Dim linhas1 As Long
    Dim linhas2 As Long
    Dim i As Long
    Dim j As Long

    
    'linhas1 = Worksheets("Nomes").Cells(Rows.Count, 7).End(xlUp).Row
    linhas1 = Workbooks("Novo Projeto").Worksheets("Nomes").Cells(Rows.Count, 7).End(xlUp).Row
    linhas2 = Worksheets("Relatório de Nf em depósito").Cells(Rows.Count, 36).End(xlUp).Row '36 ONDE VAI SER PROCURADO

    For i = 1 To linhas1
        For j = 3 To linhas2

            If (Worksheets("Relatório de Nf em depósito").Cells(j, 23) <> "SIM") Then ' 23 sera a coluna preenchida a formula

                a = InStr(1, Worksheets("Relatório de Nf em depósito").Cells(j, 36), Workbooks("Novo Projeto").Worksheets("Nomes").Cells(i, 7), 7)
                If (a = 0) Then
                    Worksheets("Relatório de Nf em depósito").Cells(j, 23) = "NAO"
                Else: Worksheets("Relatório de Nf em depósito").Cells(j, 23) = "SIM"
                End If

            End If

        Next j
    Next i

End Sub

