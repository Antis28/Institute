Attribute VB_Name = "SpaceTransferReplacement"
Option Explicit

Sub Макрос1()
Dim x
    Selection.HomeKey Unit:=wdStory
    With Selection.Find
    .ClearFormatting
    .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
 
    Selection.Find.Execute findtext:=vbCrLf, replacewith:=" ", Replace:=wdReplaceAll

End Sub

'For Each x In Split("; : ! ? . ,") 'здесь список знаков ЧЕРЕЗ ПРОБЕЛ
    'Selection.Find.Execute findtext:=" " & x, replacewith:=x, Replace:=wdReplaceAll
'Next
