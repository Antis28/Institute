Attribute VB_Name = "SpaceTransferReplacement"
Option Explicit

Sub ������1()
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

'For Each x In Split("; : ! ? . ,") '����� ������ ������ ����� ������
    'Selection.Find.Execute findtext:=" " & x, replacewith:=x, Replace:=wdReplaceAll
'Next
