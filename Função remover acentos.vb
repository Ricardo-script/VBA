Function Remover_os_Acentos(vtexto As String)
   vCom_Acento = "ÀÁÂÃÄÅÇÈÉÊËÌÍÎÏÒÓÔÕÖÙÚÛÜàáâãäåçèéêëìíîïòóôõöùúûü"
   vSem_Acento = "AAAAAACEEEEIIIIOOOOOUUUUaaaaaaceeeeiiiiooooouuuu"
   For i = 1 To Len(vtexto)
       vposicao = InStr(vCom_Acento, Mid(vtexto, i, 1))
       If vposicao > 0 Then
           vtexto = Replace(vtexto, Mid(vCom_Acento, vposicao, 1), Mid(vSem_Acento, vposicao, 1))
       End If
   Next
       Remover_os_Acentos = vtexto
End Function
