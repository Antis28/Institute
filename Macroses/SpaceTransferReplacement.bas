Attribute VB_Name = "SpaceTransferReplacement"
Option Explicit

Sub Макрос1()
    'Selection.HomeKey Unit:=wdLine, Extend:=wdMove
    'Selection.Expand Unit:=wdSection
    
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
 
    Selection.Find.Execute findtext:=wdCRLF, replacewith:=" ", Replace:=wdReplaceAll
    Selection.Find.Execute findtext:=wdLFCR, replacewith:=" ", Replace:=wdReplaceAll
    Selection.Find.Execute findtext:=wdLFOnly, replacewith:=" ", Replace:=wdReplaceAll
    Selection.Find.Execute findtext:=wdCROnly, replacewith:=" ", Replace:=wdReplaceAll
    Selection.Find.Execute findtext:=vbNewLine, replacewith:=" ", Replace:=wdReplaceAll
    Selection.Find.Execute findtext:=Chr(13), replacewith:=" ", Replace:=wdReplaceAll
    Selection.Find.Execute findtext:=Chr(10), replacewith:=" ", Replace:=wdReplaceAll
    
    Selection.Find.Execute findtext:="  ", replacewith:=" ", Replace:=wdReplaceAll
    

End Sub

'For Each x In Split("; : ! ? . ,") 'здесь список знаков ЧЕРЕЗ ПРОБЕЛ
    'Selection.Find.Execute findtext:=" " & x, replacewith:=x, Replace:=wdReplaceAll
'Next
