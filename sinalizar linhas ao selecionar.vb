Private Sub Worksheet_SelectionChange(ByVal Target As Excel.Range)

   

'*** Definição de variáveis ***

  

h = ActiveCell.Height

w2 = ActiveCell.Width

t = ActiveCell.Top

w = ActiveCell.Left

 

'Testa se os retangulos shapes são existentes.

 

  On Error Resume Next

  ActiveSheet.Shapes("RectangleV").Delete

  On Error Resume Next

 

ActiveSheet.Shapes("RectangleH").Delete

 

'Ajuste dos shapes retangulos

 

ActiveSheet.Shapes.AddShape(msoShapeRectangle, 0, t, w, h).Name = "RectangleV"

 

  With ActiveSheet.Shapes("RectangleV")

     .Fill.Visible = msoFalse

     .Fill.Transparency = 20#

     .Line.Weight = 2#

     .Line.ForeColor.SchemeColor = 10

     .PrintObject = False

  End With

 

ActiveSheet.Shapes.AddShape(msoShapeRectangle, w, 0, w2, t).Name = "RectangleH"

 

  With ActiveSheet.Shapes("RectangleH")

     .Fill.Visible = msoFalse

     .Fill.Transparency = 20#

     .Line.Weight = 2#

     .Line.ForeColor.SchemeColor = 10

   End With

 

End Sub