Attribute VB_Name = "mdlReferencesPlug"
Option Compare Database
Option Explicit

'VBA    GUID: {000204EF-0000-0000-C000-000000000046}    C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\VBE7.DLL  Version: 4.2
'Access    GUID: {4AFFC9A0-5F99-101B-AF4E-00AA003F0F07} C:\Program Files\Microsoft Office\root\Office16\MSACC.OLB   Version: 9.0
'stdole    GUID: {00020430-0000-0000-C000-000000000046} C:\Windows\System32\stdole2.tlb Version: 2.0
'DAO   GUID: {4AC9E1DA-5BAD-4AC7-86E3-24F4CDCECA28}    C:\Program Files\Common Files\Microsoft Shared\OFFICE16\ACEDAO.DLL  Version: 12.0
'VBIDE    GUID: {0002E157-0000-0000-C000-000000000046}  C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB   Version: 5.3
'ADODB  GUID: {B691E011-1797-432E-907A-4D8C69339129}    C:\Program Files\Common Files\System\ado\msado15.dll    Version: 6.1
'Office GUID: {2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}    C:\Program Files\Common Files\Microsoft Shared\OFFICE16\MSO.DLL Version: 2.8
'Excel  GUID: {00020813-0000-0000-C000-000000000046}    C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE   Version: 1.9
'Word    GUID: {00020905-0000-0000-C000-000000000046}    C:\Program Files\Microsoft Office\root\Office16\MSWORD.OLB  Version: 8.7

Sub getRef() ' добавление библиотеки из файла
Dim dbs As Database
Dim ref As Reference

Set dbs = CurrentDb()
'Set dbs = OpenDatabase("ExcAcc.accdb")
Set ref = References!Access

'Set ref = References.AddFromFile("C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB")
'Set ref = References.AddFromGuid("{4AC9E1DA-5BAD-4AC7-86E3-24F4CDCECA28}", 12, 0) ' DAO
With dbs
Set ref = References.AddFromGuid("{0002E157-0000-0000-C000-000000000046}", 5, 3) ' VBIDE
Set ref = References.AddFromGuid("{B691E011-1797-432E-907A-4D8C69339129}", 6, 1) ' ADODB
Set ref = References.AddFromGuid("{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}", 2, 8) ' Office
Set ref = References.AddFromGuid("{00020813-0000-0000-C000-000000000046}", 1, 9) ' Excel
Set ref = References.AddFromGuid("{00020905-0000-0000-C000-000000000046}", 0, 0) ' Word..,0,0<- нули - Access сам выбрал последнюю версию
End With
Set ref = Nothing
Set dbs = Nothing
End Sub

'ThisWorkbook.VBProject.References.AddFromGuid GUID:="{0002E157-0000-0000-C000-000000000046}", Major:=5, Minor:=3

