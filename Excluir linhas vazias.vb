  'Excluir notas em aberto-----------------------------------------------
           Columns("AI:AI").Select 'Adapte para a coluna que quiser
Selection.SpecialCells(xlCellTypeBlanks).Select
Selection.EntireRow.Delete
'-------------------------------------------------------------------------
