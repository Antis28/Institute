Attribute VB_Name = "Task_1"
Option Explicit

Public Sub �epreciation�alculation()
    Dim cost As Integer
    Dim deduction As Integer
    Dim result As Integer
    
    cost = InputBox("������� ��������� ������������ � ���. ���.", "������� �����")
    deduction = InputBox("������� ������� ����� ��������������� ���������� ", "������� �����")
    result = cost * deduction / 100
    
    MsgBox ("��������� ������������ ����������: " & "���. ���")
End Sub

Public Sub �epreciation�alculation2()
    Dim fixedAssets As Integer
    Dim cost As Integer
    Dim retiredEquipment As Integer
    
    
    Dim result As Integer
    
    fixedAssets = InputBox("��������� �������� ������� ����������: ", "������� �����")
    cost = InputBox("����������� ������������ �� �����: ", "������� �����")
    retiredEquipment = InputBox("������� ������������ �� �����: ", "������� �����")
    
    result = fixedAssets + cost * 8 / 12 - retiredEquipment * 3 / 12
    
    MsgBox ("������������� ��������� ����������: " & result & "���. ���")
End Sub

Public Sub �epreciation�alculation3()
    Dim costBuildings As Integer
    Dim costVehicle As Integer
    Dim costEquipment As Integer
    
    Dim result As Single
    
    costBuildings = InputBox("C�������� ������ ����������: ", "������� �����")
    costVehicle = InputBox("C�������� ������������ �������: ", "������� �����")
    costEquipment = InputBox("C�������� ������������: ", "������� �����")
    MsgBox ("������� ����� ��������������� ���������� �� ����� �������� ������� ��������� �������������� 5, 10 � 12 %")
    
    result = costBuildings * 0.05 + costVehicle * 0.1 + costEquipment * 0.12
    
    MsgBox ("����� ��������������� ���������� �����: " & result & " ���. ���")
End Sub

