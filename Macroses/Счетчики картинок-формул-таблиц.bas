Attribute VB_Name = "Counters"
Private stylePicture As String
Private styleFormula As String
Private StyleTable As String
Private StyleOriginLiterature As String

Public Sub autoopen()
    stylePicture = "stylePicture"
    styleFormula = "styleFormula"
    StyleTable = "styleTable"
    StyleOriginLiterature = "styleOriginLiterature"
    
    On Error Resume Next
    CountStylePicture
    CountStyleFormula
    CountStyleTable
    CountStyleOriginLiterature
    
End Sub

Private Sub CountStylePicture()
    ActiveDocument.Variables(stylePicture).Value = CounterStyle("К. Название рисунка")
    ChecErr5825AndAddVariable stylePicture
End Sub

Private Sub CountStyleFormula()
    ActiveDocument.Variables(styleFormula).Value = CounterStyle("К. Формула №")
    ChecErr5825AndAddVariable styleFormula
End Sub

Private Sub CountStyleTable()
    ActiveDocument.Variables(StyleTable).Value = CounterStyle("К. Название таблицы")
    ChecErr5825AndAddVariable StyleTable
End Sub

Private Sub CountStyleOriginLiterature()
    ActiveDocument.Variables(StyleOriginLiterature).Value = CounterStyle("К. Список литературы")
    ChecErr5825AndAddVariable StyleOriginLiterature
End Sub

' Проверяет на ошибку отсутсвия переменной (5825) , и содает переменную varName если ее нет
Private Sub ChecErr5825AndAddVariable(varName As String)
    If Err.Number = 5825 Then
        ActiveDocument.Variables.Add varName, 0
    End If
    ActiveDocument.Fields.Update
End Sub

' считает количество абзацев со стилем nameStyle
Private Function CounterStyle(nameStyle As String) As Integer
    Dim counter As Integer
    counter = 0
    For Each para In ActiveDocument.Paragraphs
        If InStr(nameStyle, para.Style) > 0 Then
            counter = counter + 1
        End If
    Next para
   ' MsgBox "Стилей: " & "К. Название рисунка" & ": " & counter
    CounterStyle = counter
End Function


