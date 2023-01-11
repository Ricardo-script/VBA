Dim L As Long
    Dim LF As Long
    LF = (Range("C1048576").End(xlUp).Row + 1)
    For L = 4 To LF
        Sheets("Dados").Cells(L, "AD").FormulaR1C1 = "=IFERROR(IF(VLOOKUP(RC[-27],Geral!C[-27]:C[-17],8,0)="""",""SEM ""&R3C,VLOOKUP(RC[-27],Geral!C[-27]:C[-17],8,0)),"""")"
        Sheets("Dados").Cells(L, "AE").FormulaR1C1 = "=IFERROR(RC[-1]-RC[-29]-1,""-"")"
        Sheets("Dados").Cells(L, "AI").FormulaR1C1 = "=IFERROR(IF(VLOOKUP(R[5]C[-32],Geral!C[-32]:C[-22],9,FALSE)="""",""SEM ""&R8C,VLOOKUP(R[5]C[-32],Geral!C[-32]:C[-22],9,FALSE)),"""")"
        Sheets("Dados").Cells(L, "AK").FormulaR1C1 = "=IFERROR(IF(IFERROR(VALUE(R[5]C[-1])>1,"""")="""","""",IF(R[5]C[-1]=R[5]C[-3],""No Prazo"",IF(R[5]C[-1]>R[5]C[-3],""Atrasado"",IF(R[5]C[-1]="""",""Em Trânsito"",IF(R[5]C[-1]<R[5]C[-3],""Entrega Antecipada"",""ERRO""))))),""ERRO"")"
        Sheets("Dados").Cells(L, "AM").FormulaR1C1 = "=IF(RC[-35]=R[1]C[-35],0,1)"
        Sheets("Dados").Cells(L, "AO").FormulaR1C1 = "=VLOOKUP(R[5]C[-38],Geral!C[-38]:C[-28],11,0)"
        Sheets("Dados").Cells(L, "AP").FormulaR1C1 = "=IF(RC[-39]="""","""",RC[-39]=R[-1]C[-39])"
 
    Next L
‘__________________________________________________________
Range("AL4").Select
For e = 4 To Range("C65000").End(xlUp).Row
 Range("AL" & e).FormulaR1C1 = "=IF(RC[-35]=R[1]C[-35],0,1)"
Next e
‘_____________________________________________________________
