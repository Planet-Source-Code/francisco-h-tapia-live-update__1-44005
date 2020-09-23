VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUPDATE 
   Caption         =   "Software Live Update"
   ClientHeight    =   1515
   ClientLeft      =   840
   ClientTop       =   990
   ClientWidth     =   5535
   Icon            =   "liveu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   5535
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Max             =   3
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   1140
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Live Update Control Panel"
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      Begin VB.Timer DownloadTimer 
         Left            =   4680
         Top             =   0
      End
      Begin VB.Timer StatusTimer 
         Interval        =   1
         Left            =   5040
         Top             =   0
      End
      Begin VB.Label lblNotify 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Preparing to connect..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   5295
      End
   End
End
Attribute VB_Name = "frmUPDATE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Dim Initilized As Boolean
Dim iniFILE As String, sVersionURL As String, sFileURL As String

'---------------------------------------------------------------------------------------
' Procedure : CheckVersion
' DateTime  : 3/3/2003 13:31
' Modified by : Francisco H Tapia
' Purpose   : Check for a new version on a local network share or on the internet...
'---------------------------------------------------------------------------------------
'
Private Function CheckVersion() As Boolean
    On Local Error GoTo CheckVersion_Error
    Dim x As Integer, RealFile As String, MyFile As String, UpdateVer As String
    Dim TransferSuccess As Boolean, Temp As String, iResponse As Integer

    UpdateTime = 0
    DownloadTimer.Interval = 1000
    ProgressBar1.Value = 1
    status$ = "Checking for updated version."
    
    'Determine local path to save file to
        For x = Len(sVersionURL) To 1 Step -1
            If Left$(Right$(sVersionURL, x), 1) = "/" Then RealFile$ = Right$(sVersionURL, x - 1)
        Next x
        MyFile$ = App.Path + "\" + RealFile$
        
    'Get New Version Number
    TransferSuccess = GetUpdatedFile(sVersionURL, App.Path & "\")
    If TransferSuccess = False Then Err.Raise 4050 'Trasfer Failed
    
    ProgressBar1.Value = 2
    status$ = "Version check success."
    
    'Read Downloaded Version.ver File
    Open MyFile$ For Input As #1
        Input #1, UpdateVer$
        Input #1, UpdateVer$
    Close #1
      
    If UpdateVer$ > myVer Then
        Err.Raise 4070
    Else
        Err.Raise 4060 'No Updates
    End If

    status$ = "Getting updated file."
    
   'Determine local path to save file to
        Temp$ = ReadIni(iniFILE, "Local", "File")
        For x = Len(Temp$) To 1 Step -1
            If Right$(Left$(Temp$, x), 1) = "\" Then
                MyFile$ = Left(Temp$, x)
                Exit For
            End If
        Next x
        
   'GETUPDATED FILE
   TransferSuccess = GetUpdatedFile(sFileURL, MyFile$)
   If TransferSuccess = False Then Err.Raise 4050 'Trasfer Failed
    
    ProgressBar1.Value = 3
    DownloadTimer.Interval = 0
    
    DoEvents
    x = MsgBox("Live Update Complete!", vbInformation + vbSystemModal)
    Shell ReadIni(iniFILE, "LOCAL", "File"), vbNormalFocus


CheckVersion_Exit:
    On Error Resume Next

        Unload Me
        Exit Function
        
    On Error GoTo 0
    Exit Function

CheckVersion_Error:
    Select Case Err
        Case 4050
            MsgBox "Transfer Failed!", vbCritical + vbOKOnly
            ProgressBar1.Value = 3
            DownloadTimer.Interval = 0
            status$ = "Transfer Failed!"
        Case 4060
            lblNotify.Caption = "There are no updates available."
            ProgressBar1.Value = 3
            DownloadTimer.Interval = 0
        Case 4070
            lblNotify.Caption = "There is an update available to version " + UpdateVer
            DoEvents
            iResponse = MsgBox("Do want to download the new version " + UpdateVer + " Now?", vbYesNoCancel + vbSystemModal)
            If iResponse = vbYes Then
                Resume Next
            Else
                MsgBox "Update ABORTED!", vbSystemModal + vbCritical + vbOKOnly
            End If
        Case Else
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CheckVersion of Form Form1, Line#" & Erl
            
    End Select
    Resume CheckVersion_Exit
    Resume
End Function

Private Sub DownloadTimer_Timer()
    UpdateTime = UpdateTime + 1
    StatusBar1.Panels(2).Text = "Download Time:" & Str$(UpdateTime) & " Seconds"
End Sub

Private Sub Form_Load()
Dim status As String, UpdateTime As Integer

' myVer = App.Major & "." & App.Minor & "." & App.Revision


' this is where the updated program needs to write it's current version
' number to.  The above commented out line puts the version number in
' the correct format.
    
    
    'Check for INI Files
    
    On Error GoTo Form_Load_Error

    If PathExists(App.Path & "\liveu.ini") = False Then
        'Check for INI settings file, if Missing Exit application
        Err.Raise 4000
    Else
        iniFILE = App.Path & "\liveu.ini"
    End If
    'Populate Version Location variables
    If PathExists(ReadIni(iniFILE, "FileServer", "Version")) = True Then
        sVersionURL = "file://" + Replace(ReadIni(iniFILE, "FileServer", "Version"), "\", "/")
    Else
        If IsConnected = True Then
            sVersionURL = ReadIni(iniFILE, "URL", "Version")
        Else
            Err.Raise 4010 'No Internet detected
        End If
    End If
    'Populate Updated File Location variables
    If PathExists(ReadIni(iniFILE, "FileServer", "File")) = True Then
        sFileURL = "file://" + Replace(ReadIni(iniFILE, "FileServer", "File"), "\", "/")
    Else
        If IsConnected = True Then
            sFileURL = ReadIni(iniFILE, "URL", "File")
        Else
            Err.Raise 4010 'No Internet detected
        End If
    End If
    

status$ = "Idle"
UpdateTime = 0


Open ReadIni(iniFILE, "Local", "Version") For Input As #1
    'The 2nd line is the one that is useful in this app, it's the Build Number of your application
    Input #1, myVer 'Pickup Version No.
    Input #1, myVer 'Pickup Build No.
Close #1

Exit Sub



Form_Load_Exit:
    On Error Resume Next
    
    
    On Error GoTo 0
    Exit Sub

Form_Load_Error:
    
    Select Case Err
        Case 53
            myVer = "1.0.0"
            MsgBox "Version information has not been found, Live Update will assume it's Version 1.0.0"
        Case 4000
            MsgBox "INI File is Missing, Please re-install your application", vbCritical
            Unload Me
        Case 4010
            MsgBox "An Internet connection could not be detected.", vbCritical
            Unload Me
        Case Else
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form frmUPDATE, Line#" & Erl
            
    End Select
    Resume Form_Load_Exit
    Resume

End Sub
Private Sub Form_Unload(Cancel As Integer)
    StatusBar1.Panels(1).Text = "Status: Closing..."
End Sub

Private Sub StatusTimer_Timer()
    If Initilized = False Then
        StatusBar1.Panels(1).Text = "Status: Idle"
        Call CheckVersion
        Initilized = True
    Else
        StatusBar1.Panels(1).Text = "Status: " & status$
    End If
End Sub

