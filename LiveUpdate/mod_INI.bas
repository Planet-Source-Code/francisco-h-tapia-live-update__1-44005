Attribute VB_Name = "mod_INI"
'Example by Antti Häkkinen (antti@theredstar.f2s.com)
'Visit his website at http://www.theredstar.f2s.com/
'require variable declaration
Option Explicit

'declares for ini controlling
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long



'// INI CONTROLLING PROCEDURES

'reads ini string
Public Function ReadIni(Filename As String, Section As String, Key As String) As String
Dim RetVal As String * 255, v As Long
On Local Error GoTo 100

    v = GetPrivateProfileString(Section, Key, "", RetVal, 255, Filename)
    ReadIni = Left(RetVal, v) '- 1)
Exit Function
100 If WriteIni(Filename, Section, Key, "") = True Then
        Resume
    End If

End Function

'reads ini section
Public Function ReadIniSection(Filename As String, Section As String) As String
Dim RetVal As String * 255, v As Long
On Local Error GoTo 100

v = GetPrivateProfileSection(Section, RetVal, 255, Filename)
ReadIniSection = Left(RetVal, v - 1)
Exit Function
100 If WriteIniSection(Filename, Section, "") = True Then
        Resume
    End If
End Function

'writes ini
Public Function WriteIni(Filename As String, Section As String, Key As String, Value As String) As Boolean
 On Local Error GoTo 100
    WriteIni = False
    WritePrivateProfileString Section, Key, Value, Filename
    WriteIni = True
100 'Exit as False
End Function

'writes ini section
Public Function WriteIniSection(Filename As String, Section As String, Value As String) As Boolean
 On Local Error GoTo 100
    WriteIniSection = False
    WritePrivateProfileSection Section, Value, Filename
    WriteIniSection = True
100 'Exit as False
End Function


