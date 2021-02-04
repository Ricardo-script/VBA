'A função ARRUMAR com VBA
Sub RetiraEspaco()
  Application.ScreenUpdating = False
  'With Selection
  '  .Value = Trim(Evaluate("IF(IsText(" & .Address & "), TRIM(" & .Address & "), REPT(" & .Address & ",1)) "))
  'End With
  Selection = Evaluate("if(" & Selection.Address & "="""","""",substitute(clean(trim(" & Selection.Address & ")),char(160),""""))")
  
  Application.ScreenUpdating = True
End Sub