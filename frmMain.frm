VERSION 5.00
Object = "{84926CA3-2941-101C-816F-0E6013114B7F}#1.0#0"; "imgscan.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   Caption         =   "VBCam 2002"
   ClientHeight    =   5895
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8850
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   393
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   590
   StartUpPosition =   3  'Windows Default
   Begin ScanLibCtl.ImgScan ImgScan 
      Left            =   3600
      Top             =   780
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      PageType        =   6
      CompressionType =   6
      CompressionInfo =   4096
   End
   Begin VB.ListBox lstHistory 
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   1020
      IntegralHeight  =   0   'False
      ItemData        =   "frmMain.frx":030A
      Left            =   0
      List            =   "frmMain.frx":030C
      TabIndex        =   2
      Top             =   4155
      Width           =   8775
   End
   Begin MSWinsockLib.Winsock sckWebCam 
      Index           =   0
      Left            =   3330
      Top             =   2295
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   2000
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   5520
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   661
      Style           =   1
      SimpleText      =   "Welcome to VBCam2000!"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2070
      Left            =   0
      ScaleHeight     =   134
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   167
      TabIndex        =   0
      Top             =   0
      Width           =   2565
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileTakePicture 
         Caption         =   "&Take Picture"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePictureSettings 
         Caption         =   "&Picture settings..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileWebServerSettings 
         Caption         =   "&Web Server settings..."
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Constants
Private Const CurrentModule As String = "Form1"

'Types
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONUP = &H202       'Button up
Private Const WM_LBUTTONDBLCLK = &H203   'Double-click
Private Const WM_RBUTTONUP = &H205       'Button up

'Declarations
Private Declare Function SetForegroundWindow Lib "user32" _
    (ByVal hwnd As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" _
    Alias "Shell_NotifyIconA" _
    (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

'Variables
Private LocalPort As Long
Private Ctr As Long, LastImage() As Byte
Private nid As NOTIFYICONDATA

'--------------------------------------------------------------------------------
' User Interface code
'--------------------------------------------------------------------------------

Private Sub Form_Load()
    'Load the old settings
    On Error GoTo Err_Init
    LocalPort = GetSetting(App.Title, "Preferences", "Web Port", "2000")
    
    'Initialize the scanner
    InitScanner
    
    'Initialize the Winsock port
    InitWebServer
    
    'Initialize the system tray
    With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = Me.Caption & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, nid
    
    'Take the first picture
    Show
    mnuFileTakePicture_Click

    Exit Sub

Err_Init:
    HandleError CurrentModule, "Form_Load", Err.Number, Err.Description
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'this procedure receives the callbacks from the System Tray icon.
    On Error GoTo Err_Init
    Dim msg As Long
    
    'the value of X will vary depending upon the scalemode setting
    If Me.ScaleMode = vbPixels Then
        msg = X
    Else
        msg = X / Screen.TwipsPerPixelX
    End If
    nid.szTip = Me.Caption & " - " & Ctr & " requests" & vbNullChar

    Shell_NotifyIcon NIM_MODIFY, nid
    
    Select Case msg
        Case WM_LBUTTONUP, WM_LBUTTONDBLCLK        '514,515 restore form window
            Me.WindowState = vbNormal
            SetForegroundWindow Me.hwnd
            Me.Show
        Case WM_RBUTTONUP        '517 display popup menu
            SetForegroundWindow Me.hwnd
            Me.PopupMenu Me.mnuFile
    End Select
    
    Exit Sub

Err_Init:
    HandleError CurrentModule, "Form_MouseMove", Err.Number, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Save the modified settings
    On Error GoTo Err_Init
    SaveSetting App.Title, "Preferences", "Web Port", LocalPort
    'Remove the icon from the system tray
    Shell_NotifyIcon NIM_DELETE, nid
    Exit Sub

Err_Init:
    HandleError CurrentModule, "Form_Unload", Err.Number, Err.Description
End Sub

Private Sub mnuFilePictureSettings_Click()
    'Modify the scan settings
    On Error GoTo Err_Init
    ShowScanSettings
    Exit Sub

Err_Init:
    HandleError CurrentModule, "mnuFilePictureSettings_Click", Err.Number, Err.Description
End Sub

Private Sub mnuFileWebServerSettings_Click()
    'Modify the port the web server listens in on.
    On Error GoTo Err_Init
    Dim s As String
    s = InputBox("Enter the port to listen on: ", "WebServer Settings", LocalPort)
    If Len(s) = 0 Then
        Exit Sub
    End If
    LocalPort = Val(s)
    InitWebServer
    Exit Sub

Err_Init:
    HandleError CurrentModule, "mnuFileWebServerSettings_Click", Err.Number, Err.Description
End Sub

Private Sub Picture1_Resize()
    'Adjust the controls to accommodate
    On Error GoTo Err_Init
    Form_Resize
    Exit Sub

Err_Init:
    HandleError CurrentModule, "Picture1_Resize", Err.Number, Err.Description
End Sub

Private Sub Form_Resize()
    'Adjust the history box size
    Dim NewHeight As Long, NewTop As Long
    On Error Resume Next
    If WindowState <> vbMinimized Then
        NewHeight = ScaleHeight - StatusBar1.Height - Picture1.ScaleHeight
        NewTop = Picture1.ScaleHeight
        If NewHeight < 10 Then
            NewHeight = (ScaleHeight - StatusBar1.Height) / 2
            NewTop = NewHeight
        End If
        lstHistory.Move 0, NewTop, ScaleWidth, NewHeight
    Else
        Hide
    End If
End Sub

Private Sub mnuFileTakePicture_Click()
    'Take a picture
    On Error GoTo Err_Init
    Dim s As String
    Screen.MousePointer = vbHourglass
    s = App.Path & "\twain.jpg"
    DeleteFile s
    TakePicture s
    Status "Picture stored under " & s
    Screen.MousePointer = vbNormal
    Exit Sub

Err_Init:
    HandleError CurrentModule, "mnuFileTakePicture_Click", Err.Number, Err.Description
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuFileExit_Click()
    'Quit the program
    On Error GoTo Err_Init
    Unload Me
    Exit Sub

Err_Init:
    HandleError CurrentModule, "mnuFileExit_Click", Err.Number, Err.Description
End Sub

'--------------------------------------------------------------------------------
' General Purpose Routines
'--------------------------------------------------------------------------------

Private Sub Status(ByVal s As String)
    'Display the status
    On Error GoTo Err_Init
    s = Now & "    " & s
    StatusBar1.SimpleText = s
    StatusBar1.Refresh
    lstHistory.AddItem s
    lstHistory.ListIndex = lstHistory.ListCount - 1
    Exit Sub

Err_Init:
    HandleError CurrentModule, "Status", Err.Number, Err.Description
End Sub

Private Sub DeleteFile(ByVal s As String)
    'Delete a file
    On Error GoTo Err_Init
    If Dir(s, vbNormal Or vbArchive) > "" Then
        Kill s
    End If
    Exit Sub

Err_Init:
    HandleError CurrentModule, "DeleteFile", Err.Number, Err.Description
End Sub

Private Function FileExists(ByVal s As String) As Boolean
    'Return whether or not a file exists
    On Error GoTo Err_Init
    If Dir(s, vbNormal Or vbArchive) = "" Then
        FileExists = False
    Else
        FileExists = True
    End If
    Exit Function

Err_Init:
    HandleError CurrentModule, "FileExists", Err.Number, Err.Description
End Function

Private Function LoadFile(ByVal s As String, ByRef b() As Byte) As Boolean
    'Load a file into a byte array
    On Error GoTo Err_Init
    Dim FileNo As Integer
    If Dir(s, vbNormal Or vbArchive) = "" Then
        Exit Function
    End If
    FileNo = FreeFile
    Open s For Binary Access Read As #FileNo
    ReDim b(1 To LOF(FileNo))
    Get #FileNo, , b
    Close #FileNo
    LoadFile = True
    Exit Function

Err_Init:
    HandleError CurrentModule, "LoadFile", Err.Number, Err.Description
End Function

Private Sub HandleError(ByVal CurrentModule As String, ByVal CurrentProcedure As String, ByVal ErrNum As Long, ByVal ErrDescription As String)
    Dim s As String
    On Error Resume Next
    s = CurrentModule & "_" & CurrentProcedure & ": " & ErrNum & " - " & ErrDescription
    MsgBox s, vbCritical
    Status s
End Sub

'--------------------------------------------------------------------------------
' Picture Scanning Routines
'--------------------------------------------------------------------------------

Private Sub InitScanner()
    'Initialize the scanner control
    On Error GoTo Err_Init
    With ImgScan
        If .ScannerAvailable = False Then
            Status "No TWAIN scanner detected!"
            End
        End If
        .ScanTo = FileOnly
        .FileType = JPG_File
        .ShowSetupBeforeScan = False
    End With
    Exit Sub

Err_Init:
    HandleError CurrentModule, "InitScanner", Err.Number, Err.Description
End Sub

Private Sub ShowScanSettings()
    'Let them choose the size of the picture
    Dim s As String
    On Error GoTo Err_Init
    s = App.Path & "\settings.jpg"
    ImgScan.ShowSetupBeforeScan = True
    TakePicture s
    ImgScan.ShowSetupBeforeScan = False
    DeleteFile s
    Status "Modified settings at " & Now
    Form_Resize
    Exit Sub

Err_Init:
    HandleError CurrentModule, "ShowScanSettings", Err.Number, Err.Description
End Sub

Private Function TakePicture(ByVal FileName As String) As Boolean
    'Take the picture, and store it under the desired filename.
    On Error GoTo Err_Init
    DeleteFile FileName
    ImgScan.Image = FileName
    ImgScan.StartScan
    DoEvents
    ImgScan.CloseScanner
    If FileExists(FileName) = True Then
        'Display the picture
        Picture1.Picture = LoadPicture(FileName)
        TakePicture = True
    End If
    Exit Function

Err_Init:
    HandleError CurrentModule, "TakePicture", Err.Number, Err.Description
End Function

'--------------------------------------------------------------------------------
'Web Server Routines
'--------------------------------------------------------------------------------

Private Sub InitWebServer()
    'Initialize the web server - set the port, start it listening.
    On Error GoTo Err_Init
    
    sckWebCam.Item(0).Close
    sckWebCam.Item(0).LocalPort = LocalPort
    sckWebCam.Item(0).Listen
    Status "Listening for picture requests on port " & LocalPort
    Exit Sub

Err_Init:
    HandleError CurrentModule, "InitWebServer", Err.Number, Err.Description
End Sub

Private Sub sckWebCam_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    'New incoming picture request. Hook them up to an unused connection.
    On Error GoTo Err_Init
    
    Dim i As Long, FoundOne As Boolean
    
    'Look for an unused slot
    For i = 1 To sckWebCam.Count - 1
        If sckWebCam(i).Tag = "" Then
            FoundOne = True
            Exit For
        End If
    Next i
    
    'If we didn't find one, load up a new one
    If FoundOne = False Then
        i = sckWebCam.Count
        Load sckWebCam(i)
    End If
    
    'Accept the connection
    sckWebCam(i).Tag = "INUSE"
    sckWebCam(i).Accept requestID
    
    Exit Sub

Err_Init:
    HandleError CurrentModule, "sckWebCam_ConnectionRequest", Err.Number, Err.Description
End Sub

Private Sub sckWebCam_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    On Error GoTo Err_Init
    'Get the data. Totally ignore it, because it doesn't matter what they
    ' requested, they're gonna get a WebCam shot back anyway.
    'We're using the existence of incoming data, as our indicator that it's time
    'to take a picture and send back the result.
    Dim s As String, FileName As String
    
    'Get their incoming data, but throw it away.
    sckWebCam(Index).GetData s
    
    'Take the picture
    FileName = App.Path & "\" & Format$(Ctr, "0000") & " " & Format$(Now, "YYYYMMDDHHMMSS") & ".jpg"
    TakePicture FileName
    If FileExists(FileName) = False Then
        'Send the old file, because we're currently scanning
    Else
        'Load the picture into the byte array
        Ctr = Ctr + 1
        If LoadFile(FileName, LastImage) = False Then
            Status "Unable to load picture " & FileName & "!"
            Exit Sub
        End If
        DeleteFile FileName
    End If
    
    'Send the picture via Winsock
    If sckWebCam(Index).State = sckConnected Then
        s = "HTTP/1.0 200 OK" & vbCrLf & _
            "Content-Length: " & UBound(LastImage, 1) & vbCrLf & _
            "Content-Type: image/jpeg" & vbCrLf & vbCrLf
        sckWebCam(Index).Tag = Ctr
        sckWebCam(Index).SendData s
        sckWebCam(Index).SendData LastImage
    End If
    
    Exit Sub

Err_Init:
    HandleError CurrentModule, "sckWebCam_DataArrival", Err.Number, Err.Description
    Resume Next
End Sub

Private Sub sckWebCam_Close(Index As Integer)
    'Close the connection, and mark the slot as unused.
    On Error GoTo Err_Init
    sckWebCam(Index).Close
    sckWebCam(Index).Tag = ""
    Exit Sub

Err_Init:
    HandleError CurrentModule, "sckWebCam_Close", Err.Number, Err.Description
End Sub

Private Sub sckWebCam_SendComplete(Index As Integer)
    'Done sending the picture.
    On Error GoTo Err_Init
    Status "Sent picture " & sckWebCam(Index).Tag & " to slot " & Index & " - " & sckWebCam(Index).RemoteHostIP & " " & sckWebCam(Index).RemoteHost
    sckWebCam_Close Index
    Exit Sub

Err_Init:
    HandleError CurrentModule, "sckWebCam_SendComplete", Err.Number, Err.Description
End Sub

Private Sub sckWebCam_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'Display an error
    On Error GoTo Err_Init
    Status "sckWebCam_Error on slot " & Index & "! Err num " & Number & " - " & Description
    sckWebCam_Close Index
    Exit Sub

Err_Init:
    HandleError CurrentModule, "sckWebCam_Error", Err.Number, Err.Description
End Sub

