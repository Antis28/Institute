Attribute VB_Name = "CleanerText"
Sub CleanerText()
' Очистка текста от двойных пробелов,
' переносов слов, двойных абзатцев.


    Dim myRange As Range
    Dim counSpaces As Integer

    Set myRange = Selection.Range
'
' Замена двойных пробелов на одинарные
'
'
    Do
        SearchAndReplaceInRange "  ", " "
        counSpaces = SearchInRange("  ")
        myRange.Select
    Loop While counSpaces > 0
    
    SearchAndReplaceInRange "- ", ""
    SearchAndReplaceInRange " ^p", " "
    SearchAndReplaceInRange "^p^p", "^p"

    'myRange.Select
    
Exit Sub
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "  "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        End With
    Selection.Find.Execute Replace:=wdReplaceOne
    
'
' Замена '-_' на слитное написание
'
'
    myRange.Select
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
 With Selection.Find
        .Text = "- "
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
'
' Замена переносов на пробелы
'
'
    myRange.Select
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
 With Selection.Find
        .Text = "^p"
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Function SearchInRange(value As String)
  Dim bSel As Long
  Dim eSel As Long
  Dim count As Long
  count = 0
  bSel = Selection.Start
  eSel = Selection.End
  Do While True
    With Selection.Find
      .Text = value
      If .Execute Then
            count = count + 1
        bSel = Selection.End
        Selection.Start = bSel
        Selection.End = eSel
      Else
        'MsgBox count
        SearchInRange = count
        Exit Function
      End If
    End With
  Loop
End Function



Sub SearchAndReplaceInRange(findVal As String, replaceVal As String)
    Dim bSel As Long
    Dim eSel As Long
    bSel = Selection.Start
    eSel = Selection.End
  
  
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findVal
        .Replacement.Text = replaceVal
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
