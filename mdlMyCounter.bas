Attribute VB_Name = "mdlMyCounter"
Option Compare Database

Private curNum As Long

Public Function startNum() As Boolean
  curNum = 0
  startNum = True
End Function

Public Function GetNextNum(anyField) As Long
  curNum = curNum + 1
  GetNextNum = curNum
End Function
