Sub CopyFormattedRange(sheetName As String, rngName As String)
    Dim rng As Range
    On Error GoTo ErrHandler
    
    ' Référence la feuille passée en paramètre
    Set rng = ThisWorkbook.Sheets(sheetName).Range(rngName)
    
    ' Copier valeurs + formats
    rng.Copy
    Exit Sub

ErrHandler:
    MsgBox "Erreur: impossible de copier la plage " & rngName & " sur la feuille " & sheetName, vbCritical
End Sub


Sub CopyFormattedCell(sheetName As String, cellAddress As String)
    Dim rng As Range
    On Error GoTo ErrHandler
    
    ' Référence la feuille passée en paramètre
    Set rng = ThisWorkbook.Sheets(sheetName).Range(cellAddress)
    
    ' Copier valeur + format
    rng.Copy
    Exit Sub

ErrHandler:
    MsgBox "Erreur: impossible de copier la cellule " & cellAddress & " sur la feuille " & sheetName, vbCritical
End Sub
