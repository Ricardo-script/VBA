Public Function OpenFileDialog() As String
    Dim Filter As String, Title As String
    Dim FilterIndex As Integer
    Dim Filename As Variant
    ' Define o filtro de procura dos arquivos
    Filter = "Arquivo do Excel (*.xls; *.xlsx), *.xls;*.xlsx"
    ' O filtro padrão é *.*
    FilterIndex = 3
    ' Define o Título (Caption) da Tela
    Title = "Selecione um arquivo"
    ' Define o disco de procura
    ChDrive ("C")
    ChDir ("C:\")
    With Application
        ' Abre a caixa de diálogo para seleção do arquivo com os parâmetros
        Filename = .GetOpenFilename(Filter, FilterIndex, Title)
        ' Reseta o Path
        ChDrive (Left(.DefaultFilePath, 1))
        ChDir (.DefaultFilePath)
    End With
    ' Abandona ao Cancelar
    If Filename = False Then
        MsgBox "Nenhum arquivo foi selecionado."
        Exit Function
    End If
    ' Retorna o caminho do arquivo
    OpenFileDialog = Filename
End Function
