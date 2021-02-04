 Dim L As Long
    Dim LF As Long
    
    Dim AD As Long
    ar = (Range("AR1048576").End(xlUp).Row + 1)
    ap = (Range("AP1048576").End(xlUp).Row + 1)
    an = (Range("AN1048576").End(xlUp).Row + 1)
    al = (Range("AL1048576").End(xlUp).Row + 1)
    aj = (Range("AJ1048576").End(xlUp).Row + 1)
    ag = (Range("AG1048576").End(xlUp).Row + 1)
    
    
    LF = (Range("AA1048576").End(xlUp).Row + 1)
    For L = 2 To LF
        If L >= AD Then Sheets("Filiais").Cells(L, "AR").FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-1],'[Novo Projeto.xlsm]Nomes'!C7,1,0),"""")"
        If L >= AD Then Sheets("Filiais").Cells(L, "AP").FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-1],'[Novo Projeto.xlsm]Nomes'!C7,1,0),"""")"
        If L >= AD Then Sheets("Filiais").Cells(L, "AN").FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-1],'[Novo Projeto.xlsm]Nomes'!C7,1,0),"""")"
        If L >= AD Then Sheets("Filiais").Cells(L, "AL").FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-1],'[Novo Projeto.xlsm]Nomes'!C7,1,0),"""")"
        If L >= AD Then Sheets("Filiais").Cells(L, "AJ").FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-1],'[Novo Projeto.xlsm]Nomes'!C7,1,0),"""")"
        If L >= AD Then Sheets("Filiais").Cells(L, "AG").FormulaR1C1 = "=RC[3]&"" - ""&RC[5]&"" - ""&RC[7]&"" - ""&RC[9]&"" - ""&RC[11]"
       
    Next L