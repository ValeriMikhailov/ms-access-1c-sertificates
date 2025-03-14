Attribute VB_Name = "mdlSplitRow"
Option Compare Database
Option Explicit

Public Function SplitRow(fldSplit As Variant, numSplit As Byte) As Variant
    Dim A
    If Not IsNull(fldSplit) Then
        A = Split(fldSplit, Chr(32), 9) '��������� �� ������ ����� �� "�������"
        If numSplit <= UBound(A) Then
            SplitRow = A(numSplit)
        End If
    End If
End Function

Public Function SplitOtkaz(fldSplit As Variant, numSplit As Byte) As Variant
    Dim A
    If Not IsNull(fldSplit) Then
        A = Split(fldSplit, "����� ", 2) '��������� �� ��� ���� �� "����� "
        If numSplit <= UBound(A) Then
            SplitOtkaz = A(numSplit)
        End If
    End If
End Function

Public Function SplitArticul(fldSplit As Variant, numSplit As Byte) As Variant
    Dim A
    If Not IsNull(fldSplit) Then
        A = Split(fldSplit, "_����", 2) '��������� �� ��� ���� �� "_����"
        If numSplit <= UBound(A) Then
            SplitArticul = A(numSplit)
        End If
    End If
End Function
