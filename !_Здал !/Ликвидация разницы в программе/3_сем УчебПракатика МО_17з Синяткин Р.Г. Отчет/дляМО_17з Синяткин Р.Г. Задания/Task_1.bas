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
    
    MsgBox ("" & result & "���. ���")
End Sub

