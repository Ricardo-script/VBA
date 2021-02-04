
 Dim wkbk As Workbook
 Dim NewFile As Variant
 NewFile = Application.GetOpenFilename("microsoft excel files (*.xlsx*), *.xlsx*")
 Workbooks.Open (NewFile)