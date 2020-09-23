Attribute VB_Name = "modUpdates"
'****************************************************************
'
' Live Update Code
'
' Written by:  Francisco H Tapia
'              3/11/2003


' Special Thanks to Blake Pell <blakepell@hotmail.com> for inspiring this
' http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=13413&lngWId=1

' This code is open source, I would appreciate that anybody using
' this is a released application to e-mail or get in contact with
' me.
' The original concept of this code and much of the Form code is very similar to that of Blake's
' I wanted to use an INI FILE source for the URL paths of the web and destination files because
' I did not want to keep re-compiling the project for other programs...
' Also the space where Version.ver downloads into is sperate from the program so that if there
' are errors the program will re-download the update instead of staying broken.
' This version of the Live Update code, does away with user interaction, except to ask the
' user if they want to continue w/ the D/L.
' I hope this makes someone's day easier or helps them learn
' a bit as it did for me.

' WININET.DLL SOURCE:  http://support.microsoft.com/default.aspx?scid=KB;EN-US;Q232194&
'
'
'****************************************************************


Option Explicit

Public Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Public Const INTERNET_OPEN_TYPE_DIRECT = 1
Public Const INTERNET_OPEN_TYPE_PROXY = 3

Public Const scUserAgent = "VB OpenUrl"
Public Const INTERNET_FLAG_RELOAD = &H80000000

Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
(ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, _
ByVal sProxyBypass As String, ByVal lFlags As Long) As Long

Public Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" _
(ByVal hOpen As Long, ByVal sUrl As String, ByVal sHeaders As String, _
ByVal lLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long

Public Declare Function InternetReadFile Lib "wininet.dll" _
(ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, _
lNumberOfBytesRead As Long) As Integer

Public Declare Function InternetCloseHandle Lib "wininet.dll" _
(ByVal hInet As Long) As Integer

Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal dwReserved As Long) As Long


Global myVer As String
Global status$
Global UpdateTime As Integer

Public Function IsConnected() As Boolean
    If InternetGetConnectedState(0&, 0&) = 1 Then
        IsConnected = True
    Else
        IsConnected = False
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetUpdatedFile
' DateTime  : 3/11/2003 13:12
' Author    : Francisco H Tapia <fhtapia@hotmail.com>
' Purpose   : OPENURL alternative...
'SOURCE     : http://support.microsoft.com/default.aspx?scid=KB;EN-US;Q232194&
'---------------------------------------------------------------------------------------
'
Public Function GetUpdatedFile(sUrl As String, DestDIR As String) As Boolean
    Dim hOpen               As Long
    Dim hOpenUrl            As Long
    Dim bDoLoop             As Boolean
    Dim bRet                As Boolean
    Dim sReadBuffer         As String * 2048
    Dim lNumberOfBytesRead  As Long
    Dim sBuffer             As String
    Dim myData() As Byte, x As Integer, MyFile As String, RealFile As String

    
    On Error GoTo GetUpdatedFile_Error

    For x = Len(sUrl) To 1 Step -1
        If Left$(Right$(sUrl, x), 1) = "/" Then RealFile$ = Right$(sUrl, x - 1)
    Next x
    MyFile$ = DestDIR + RealFile$

    If PathExists(MyFile$) = True Then
        Kill MyFile$ 'delete file before the download helps avoid file mismatch errors
    End If
    
    hOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    hOpenUrl = InternetOpenUrl(hOpen, sUrl, vbNullString, 0, INTERNET_FLAG_RELOAD, 0)

    bDoLoop = True
    While bDoLoop
        sReadBuffer = vbNullString
        bRet = InternetReadFile(hOpenUrl, sReadBuffer, Len(sReadBuffer), lNumberOfBytesRead)
        sBuffer = sBuffer & Left$(sReadBuffer, lNumberOfBytesRead)
        If Not CBool(lNumberOfBytesRead) Then bDoLoop = False
    Wend
    
    

    Open MyFile$ For Binary Access Write As #1
    Put #1, , sBuffer
    Close #1
    
    If hOpenUrl <> 0 Then InternetCloseHandle (hOpenUrl)
    If hOpen <> 0 Then InternetCloseHandle (hOpen)
    


    GetUpdatedFile = True

GetUpdatedFile_Exit:
    On Error Resume Next

    On Error GoTo 0
    Exit Function

GetUpdatedFile_Error:
    Select Case Err
        Case 4000
            MsgBox "There is no connection to the internet.  Please try again later.", vbInformation
        Case Else
            MsgBox "An error has occured in the file transfer or write.  Please try again later. " & vbCrLf & _
            "Error " & Err.Number & " (" & Err.Description & ") in procedure GetUpdatedFile of Module modUpdates, Line#" & Erl
    End Select
    GetUpdatedFile = False
    Resume GetUpdatedFile_Exit
    Resume

End Function




