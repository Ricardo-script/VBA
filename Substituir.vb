Dim Col As Variant, Word As String
Let Col = ("J")
If Len(Col) > 0 And Not Col Like "*[!0-9]*" Then Col = Val(Col)
Let Word = "NUFARM - LJ 13 COLINAS TOCANTI"
With Columns(Col)
.Replace Word, "NUFARM - LJ 13 COLINAS DO TOCANTINS", xlWhole
End With