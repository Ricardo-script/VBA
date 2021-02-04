Põe o seguinte código num module:
Código (Text):
Public pasta As String
Sub carrega_form()
    Load UserForm
    UserForm.Show nomodal
End Sub
e o seguinte código no Form:

Código (Text):
Private Sub UserForm_Initialize()
    pasta = ActiveWindow.Caption
    ActiveWindow.Visible = False
End Sub
Private Sub UserForm_Terminate()
    Windows(pasta).Visible = True
End Sub
Quando você rodar a sub "carrega_form" a pasta de trabalho ativa vai ficar invisível e o form vai aparecer, enquanto as demais pastas de trabalho permanecem visíveis e editáveis.


Minimizar excel, somente mostrar form:

Private Sub UserForm_Initialize()
Application.WindowState = xlMinimized
End Sub

OU:
Private Sub UserForm_Initialize() 
Application.WindowState = xlMinimized
End Sub
Private Sub UserForm_Terminate()
Application.WindowState = xlMinimized
End Sub

OU:
Private Sub Workbook_Open()
Application.WindowState = xlMinimized
Menu_Principal.Show
End Sub
