P�e o seguinte c�digo num module:
C�digo (Text):
Public pasta As String
Sub carrega_form()
    Load UserForm
    UserForm.Show nomodal
End Sub
e o seguinte c�digo no Form:

C�digo (Text):
Private Sub UserForm_Initialize()
    pasta = ActiveWindow.Caption
    ActiveWindow.Visible = False
End Sub
Private Sub UserForm_Terminate()
    Windows(pasta).Visible = True
End Sub
Quando voc� rodar a sub "carrega_form" a pasta de trabalho ativa vai ficar invis�vel e o form vai aparecer, enquanto as demais pastas de trabalho permanecem vis�veis e edit�veis.


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
