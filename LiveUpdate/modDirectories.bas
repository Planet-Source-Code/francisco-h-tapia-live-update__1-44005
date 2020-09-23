Attribute VB_Name = "modDirectories"
'Author: Unknown...
'******************** Code Start **************************
Private Const MAX_PATH As Integer = 255
Private Declare Function apiGetSystemDirectory& Lib "kernel32" _
        Alias "GetSystemDirectoryA" _
        (ByVal lpBuffer As String, ByVal nSize As Long)

Private Declare Function apiGetWindowsDirectory& Lib "kernel32" _
        Alias "GetWindowsDirectoryA" _
        (ByVal lpBuffer As String, ByVal nSize As Long)

Private Declare Function apiGetTempDir Lib "kernel32" _
        Alias "GetTempPathA" (ByVal nBufferLength As Long, _
        ByVal lpBuffer As String) As Long
Function PathExists(pname) As Boolean
'   Author: Unknown
'   Returns TRUE if the path or iniFILE exists
    Dim x As String
    On Error Resume Next
    x = GetAttr(pname) And 0
    If Err = 0 Then PathExists = True Else: PathExists = False
End Function
Function fReturnDir(ByVal strDir As String)
'Returns Requested folder, compiled the original functions to 1 function with a string parameter
    Dim lngx As Long
    Dim strDirectory As String
    strDirectory = strDir
    
    strDir = String$(MAX_PATH, 0)
    Select Case strDirectory
        Case "Temp"
            lngx = apiGetTempDir(MAX_PATH, strDir)
            If lngx <> 0 Then
                fReturnDir = Left$(strDir, lngx)
            Else
                fReturnDir = ""
            End If
        Case "System"
                lngx = apiGetSystemDirectory(strDir, MAX_PATH)
                If lngx <> 0 Then
                    fReturnDir = Left$(strDir, lngx)
                Else
                    fReturnDir = ""
                End If
        Case "Windows"
                lngx = apiGetWindowsDirectory(strDir, MAX_PATH)
                If lngx <> 0 Then
                    fReturnDir = Left$(strDir, lngx)
                Else
                    fReturnDir = ""
                End If
        Case Else
            fReturnDir = "Invalid Parameter. Valid Types: Temp; System; Windows"
    End Select
End Function


