Attribute VB_Name = "Task_1"
Option Explicit

Public Sub ВepreciationСalculation()
    Dim cost As Integer
    Dim deduction As Integer
    Dim result As Integer
    
    cost = InputBox("Введите стоимость оборудования в тыс. руб.", "Введите число")
    deduction = InputBox("Введите годовую норму амортизационных отчислений ", "Введите число")
    result = cost * deduction / 100
    
    MsgBox ("Стоимость оборудования составляет: " & "тыс. руб")
End Sub

Public Sub ВepreciationСalculation2()
    Dim fixedAssets As Integer
    Dim cost As Integer
    Dim retiredEquipment As Integer
    
    
    Dim result As Integer
    
    fixedAssets = InputBox("Стоимость основных средств составляет: ", "Введите число")
    cost = InputBox("Приобретено оборудование на сумму: ", "Введите число")
    retiredEquipment = InputBox("Списано оборудование на сумму: ", "Введите число")
    
    result = fixedAssets + cost * 8 / 12 - retiredEquipment * 3 / 12
    
    MsgBox ("" & result & "тыс. руб")
End Sub

