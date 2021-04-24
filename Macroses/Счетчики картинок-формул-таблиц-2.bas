Attribute VB_Name = "Counters"
Private stylePictureVal As String
Private styleFormulaVal As String
Private styleTableVal As String
Private styleOriginLiteratureVal As String

Private styleNamePicture  As String
Private styleNameFormula  As String
Private styleNameTable  As String
Private styleNameOriginLiterature  As String

Private countPicture As Integer
Private countFormula As Integer
Private countTable As Integer
Private countOriginLiterature As Integer

Public Sub autoopen()
    stylePictureVal = "stylePicture"
    styleFormulaVal = "styleFormula"
    styleTableVal = "styleTable"
    styleOriginLiteratureVal = "styleOriginLiterature"
    
    styleNamePicture = "К. Название рисунка"
    styleNameFormula = "К. Формула №"
    styleNameTable = "К. Название таблицы"
    styleNameOriginLiterature = "К. Список литературы"
    
    CounterStyle
    
    On Error Resume Next
    CountStylePicture
    CountStyleFormula
    CountStyleTable
    CountStyleOriginLiterature
    
     ActiveDocument.Fields.Update
End Sub

Private Sub CountStylePicture()
    ActiveDocument.Variables(stylePictureVal).Value = countPicture
    ChecErr5825AndAddVariable stylePictureVal
End Sub

Private Sub CountStyleFormula()
    ActiveDocument.Variables(styleFormulaVal).Value = countFormula
    ChecErr5825AndAddVariable styleFormulaVal
End Sub

Private Sub CountStyleTable()
    ActiveDocument.Variables(styleTableVal).Value = countTable
    ChecErr5825AndAddVariable styleTableVal
End Sub

Private Sub CountStyleOriginLiterature()
    ActiveDocument.Variables(styleOriginLiteratureVal).Value = countOriginLiterature
    ChecErr5825AndAddVariable styleOriginLiteratureVal
End Sub

' Проверяет на ошибку отсутсвия переменной (5825) , и содает переменную varName если ее нет
Private Sub ChecErr5825AndAddVariable(varName As String)
    If Err.Number = 5825 Then
        ActiveDocument.Variables.Add varName, 0
    End If
End Sub

' считает количество абзацев со стилями
Private Sub CounterStyle()
    
    countPicture = 0
    countFormula = 0
    countTable = 0
    countOriginLiterature = 0
        
    For Each para In ActiveDocument.Paragraphs
        If InStr(styleNamePicture, para.Style) > 0 Then
            countPicture = countPicture + 1
        End If
        
        If InStr(styleNameFormula, para.Style) > 0 Then
            countFormula = countFormula + 1
        End If
        
        If InStr(styleNameTable, para.Style) > 0 Then
            countTable = countTable + 1
        End If
        
        If InStr(styleNameOriginLiterature, para.Style) > 0 Then
            countOriginLiterature = countOriginLiterature + 1
        End If
    Next para
End Sub


