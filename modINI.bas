Attribute VB_Name = "modINI"
Private Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpRetunedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lplFileName As String) As Long

Public Function INIRead(ByVal Section As String, ByVal Key As String) As String
   Dim RetStr As String
   RetStr = String(255, Chr(0))
   INIRead = Left(RetStr, GetPrivateProfileString(Section, ByVal Key, "", RetStr, Len(RetStr), INIFile))
End Function

Public Function INIWrite(ByVal Section As String, ByVal Key As String, ByVal Value As String) As Boolean
    INIWrite = CBool(WritePrivateProfileString(Section, Key, Value, INIFile) = 1)
End Function

Public Function INIDeleteSection(ByVal Section As String) As Boolean
    'Delete a Section, including keynames and values
    INIDeleteSection = CBool(WritePrivateProfileString(Section, ByVal 0&, ByVal 0&, INIFile))
End Function

Public Function INIDeleteKey(ByVal Section As String, ByVal Key As String) As Boolean
    'Deletes a keyname and it's value
    INIDeleteKey = CBool(WritePrivateProfileString(Section, Key, ByVal 0&, INIFile))
End Function
