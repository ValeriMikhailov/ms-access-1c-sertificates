Attribute VB_Name = "mdlSplitRow"
Option Compare Database
Option Explicit

Public Function SplitRow(fldSplit As Variant, numSplit As Byte) As Variant
    Dim A
    If Not IsNull(fldSplit) Then
        A = Split(fldSplit, Chr(32), 9) 'разбиение на ДЕВЯТЬ полей по "пробелу"
        If numSplit <= UBound(A) Then
            SplitRow = A(numSplit)
        End If
    End If
End Function

Public Function SplitOtkaz(fldSplit As Variant, numSplit As Byte) As Variant
    Dim A
    If Not IsNull(fldSplit) Then
        A = Split(fldSplit, "отказ ", 2) 'разбиение на ДВА поля по "отказ "
        If numSplit <= UBound(A) Then
            SplitOtkaz = A(numSplit)
        End If
    End If
End Function

Public Function SplitArticul(fldSplit As Variant, numSplit As Byte) As Variant
    Dim A
    If Not IsNull(fldSplit) Then
        A = Split(fldSplit, "_СНЯТ", 2) 'разбиение на ДВА поля по "_СНЯТ"
        If numSplit <= UBound(A) Then
            SplitArticul = A(numSplit)
        End If
    End If
End Function
