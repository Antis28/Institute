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
    
    MsgBox ("Среднегодовая стоимость составляет: " & result & "тыс. руб")
End Sub

Public Sub ВepreciationСalculation3()
    Dim costBuildings As Integer
    Dim costVehicle As Integer
    Dim costEquipment As Integer
    
    Dim result As Single
    
    costBuildings = InputBox("Cтоимость зданий составляет: ", "Введите число")
    costVehicle = InputBox("Cтоимость транспортных средств: ", "Введите число")
    costEquipment = InputBox("Cтоимость оборудования: ", "Введите число")
    MsgBox ("Средняя норма амортизационных отчислений по видам основных средств составила соответственно 5, 10 и 12 %")
    
    result = costBuildings * 0.05 + costVehicle * 0.1 + costEquipment * 0.12
    
    MsgBox ("Сумма амортизационных отчислений равна: " & result & " млн. руб")
End Sub

