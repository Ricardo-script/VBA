Sub fMain()
    Dim lngBD As Long
    Dim lngLast As Long
    Dim wksBD As Worksheet
    Dim wks As Worksheet
    
    Set wksBD = ThisWorkbook.Sheets("Relatório de Nf em depósito")
    With wksBD
        For lngBD = 2 To .Cells(.Rows.Count, "Q").End(xlUp).Row
            Set wks = Nothing
            On Error Resume Next
            Set wks = ThisWorkbook.Sheets(CStr(.Cells(lngBD, "Q")))
            On Error GoTo 0
            If wks Is Nothing Then
                Set wks = ThisWorkbook.Sheets.Add
                wks.Name = CStr(.Cells(lngBD, "Q"))
                wksBD.Rows(1).Copy wks.Rows(1)
            End If
            lngLast = wks.Cells(wks.Rows.Count, "Q").End(xlUp).Row + 1
            wksBD.Rows(lngBD).Copy wks.Rows(lngLast)
        Next lngBD
    End With
End Sub
