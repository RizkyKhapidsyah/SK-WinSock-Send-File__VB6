VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Winsock SendFile Demo"
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7140
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   7140
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame fraSending 
      Caption         =   "Frame1"
      Height          =   1965
      Left            =   90
      TabIndex        =   13
      Top             =   2670
      Visible         =   0   'False
      Width           =   6885
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   585
         Left            =   3480
         TabIndex        =   15
         Top             =   300
         Width           =   3105
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "&Send File"
         Height          =   585
         Left            =   360
         TabIndex        =   14
         Top             =   300
         Width           =   3075
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   495
         Left            =   330
         TabIndex        =   16
         Top             =   1320
         Width           =   6285
         _ExtentX        =   11086
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label lblTransferStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   360
         TabIndex        =   17
         Top             =   960
         Width           =   6255
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1560
      Top             =   1890
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSWinsockLib.Winsock wskClient 
      Left            =   1080
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1000
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "&Disconnect"
      Height          =   405
      Left            =   4500
      TabIndex        =   12
      Top             =   870
      Width           =   2535
   End
   Begin MSWinsockLib.Winsock wskServer 
      Left            =   630
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   1000
   End
   Begin VB.TextBox txtLocalIP 
      BackColor       =   &H8000000A&
      Height          =   315
      Left            =   4500
      TabIndex        =   11
      Top             =   450
      Width           =   2505
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6450
      TabIndex        =   8
      Top             =   2310
      Width           =   585
   End
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   60
      TabIndex        =   7
      Top             =   2310
      Width           =   6345
   End
   Begin VB.CommandButton cmdHost 
      Caption         =   "&Host"
      Height          =   405
      Left            =   2220
      TabIndex        =   5
      Top             =   870
      Width           =   2235
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Connect"
      Enabled         =   0   'False
      Height          =   405
      Left            =   60
      TabIndex        =   4
      Top             =   870
      Width           =   2115
   End
   Begin VB.OptionButton optClient 
      BackColor       =   &H8000000C&
      Caption         =   "Client"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   90
      TabIndex        =   3
      Top             =   480
      Width           =   2085
   End
   Begin VB.OptionButton optServer 
      BackColor       =   &H8000000C&
      Caption         =   "Server"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   405
      Left            =   90
      TabIndex        =   2
      Top             =   90
      Value           =   -1  'True
      Width           =   2085
   End
   Begin VB.TextBox txtRemoteIP 
      Height          =   315
      Left            =   4500
      TabIndex        =   0
      Top             =   90
      Width           =   2505
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Press F1 for help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   2670
      TabIndex        =   18
      Top             =   1380
      Width           =   1785
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Local IP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   2220
      TabIndex        =   10
      Top             =   450
      Width           =   2205
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "File to send:"
      Height          =   195
      Left            =   90
      TabIndex        =   9
      Top             =   2100
      Width           =   855
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   60
      TabIndex        =   6
      Top             =   1650
      Width           =   6975
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Remote IP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   2220
      TabIndex        =   1
      Top             =   90
      Width           =   2205
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This boolean indicates if this side if currently sending a file
Dim SendingFile As Boolean
'This boolean indicates if the other side denies to recieve the file
Dim AbortFile As Boolean
'This variable contains the file that is being sent from the other side
Dim RecievedFile As String
'This variable is used in keeping track of the transfer rate.
Dim BeginTransfer As Single


Private Sub cmdCancel_Click()

   Dim Answer As VbMsgBoxResult
   Answer = MsgBox("Do you really want to abort the transfer ?", vbInformation Or vbYesNo)
   If Answer = vbYes Then
      Unload Me
   End If

End Sub

Private Sub cmdConnect_Click()

   On Error GoTo ErrorHandler

   With wskClient
      .Close
      .RemoteHost = Trim(txtRemoteIP.Text)
      .Connect
   End With
   
   lblStatus.Caption = "Connection Status : CONNECTING"
   
   Exit Sub
   
ErrorHandler:

   MsgBox "Unable to connect.", vbInformation
   
End Sub

Private Sub cmdDisconnect_Click()
   
   Dim Reply As String
   Dim WinsockControl As Winsock
   
   If optServer.Value = True Then
      Set WinsockControl = wskServer
   Else
      Set WinsockControl = wskClient
   End If
   
   MsgBox WinsockControl.State
   
   If WinsockControl.State = sckConnected Then
      Reply = MsgBox("You are currently connected." & vbCrLf & _
      "Do you wish to disconnect?", vbYesNo Or vbInformation)
      If Reply = vbNo Then Exit Sub
   End If
            
   WinsockControl.Close
   lblStatus.Caption = "Connection Status : INACTIVE"

End Sub

Private Sub cmdHost_Click()

   With wskServer
      .Close
      .Listen
   End With
   
   lblStatus.Caption = "Connection Status : LISTENING"

End Sub

Private Sub cmdSend_Click()

   SendFile txtFileName.Text

End Sub

Private Sub Command1_Click()

   If lblStatus.Caption <> "Connection Status : HOSTING - READY TO SEND FILE" Then
      MsgBox "You must be hosting a connection in order to send a file.", vbInformation
      Exit Sub
   End If

   On Error GoTo ErrorHandler
   
   CommonDialog1.ShowOpen
   txtFileName = CommonDialog1.FileName
   
   fraSending.Caption = "Sending File..."
   fraSending.Visible = True
   
   Exit Sub
   
ErrorHandler:

   Exit Sub

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

   If KeyCode = vbKeyF1 Then
      ChDir App.Path
      Shell "notepad.exe readme.txt", vbNormalFocus
   End If

End Sub

Private Sub Form_Load()

   lblStatus.Caption = "Connection Status : INACTIVE"
   txtLocalIP.Text = wskServer.LocalIP
   
End Sub

Private Sub Form_Unload(Cancel As Integer)

   wskClient.Close
   wskServer.Close

End Sub

Private Sub optClient_Click()

   cmdHost.Enabled = Not optClient.Value
   cmdConnect.Enabled = optClient.Value

End Sub

Private Sub optServer_Click()

   cmdConnect.Enabled = Not optServer.Value
   cmdHost.Enabled = optServer.Value
   
End Sub

Private Sub wskClient_Close()

   MsgBox "The remote computer has closed the connection.", vbInformation
   wskClient.Close

End Sub

Private Sub wskClient_Connect()

   lblStatus.Caption = "Connection Status : CONNECTED - WAITING FOR FILE"

End Sub

Private Sub wskClient_DataArrival(ByVal bytesTotal As Long)
   
   'The data that comes from the server is either:
   '   a) the string "S_######_$$$$$$" where
   '      ##### and $$$$$$$$ are the size and filename of the
   '      file that the server wishes to send over.
   '   b) data from the actual file transfer that is taking place

   On Error GoTo ErrorHandler
   
   Dim NewData As String
   'Store the new data in to the variable NewData
   wskClient.GetData NewData
   
   Static FileName As String
   Static FileSize As Long
   
   'If this is a request from the server ,to send a new file
   If Left(NewData, 2) = "S_" Then
      
      FileSize = Val(Mid(NewData, 3, 3 + InStr(4, NewData, "_")))
      FileName = Mid(NewData, InStr(4, NewData, "_") + 1)
      
      Dim Question As String
      Dim Answer As VbMsgBoxResult
      
      Question = "The remote computer wishes to send you this file:" & vbCrLf & _
               FileName & " (" & FileSize & " bytes)" & vbCrLf & vbCrLf & _
               "Recieve this file? "
      Answer = MsgBox(Question, vbInformation Or vbYesNo)
      
      'then prompt the user with the options to accept or decline
      'the file transfer.
      If Answer = vbYes Then
         'Prepare for the file transfer
         fraSending.Caption = "Recieving file " & FileName
         cmdSend.Visible = False
         lblTransferStatus.Caption = "Recieved 0 bytes (0%)"
         ProgressBar1.Max = FileSize
         ProgressBar1.Value = 0
         fraSending.Visible = True
         RecievedFile = ""
         'The string "R_" means that this side accepts the file
         'transfer
         wskClient.SendData "R_"
         BeginTransfer = Timer
      Else
         'The string "N_" means that this side doesnt accept the file
         'transfer
         wskClient.SendData "N_"
      End If
      
   Else
      'if this is data from the actual file transfer then
      'add it to the variable that contains the data already sent.
      RecievedFile = RecievedFile & NewData
      lblTransferStatus.Caption = "Recieved " & Len(RecievedFile) & " bytes (" & Format((Len(RecievedFile) * 100) / FileSize, "00.0") & "%) - " & Format(Len(RecievedFile) / (Timer - BeginTransfer) / 1000, "0.0") & " kbps"
      ProgressBar1.Value = Len(RecievedFile)
      ProgressBar1.Refresh
      lblTransferStatus.Refresh
      
      'check if the file transfer is complete
      If Len(RecievedFile) = FileSize Then
         CommonDialog1.FileName = FileName
         CommonDialog1.ShowSave
         'prompt the user for a path to save the file in
         Open CommonDialog1.FileName For Binary As #1
         Put #1, 1, RecievedFile
         Close
      End If
      DoEvents
   
   End If

   Exit Sub
   
ErrorHandler:

End Sub

Private Sub wskServer_Close()

   MsgBox "The remote computer has closed the connection.", vbInformation
   wskServer.Close

End Sub

Private Sub WskServer_ConnectionRequest(ByVal requestID As Long)

   If wskServer.State <> sckClosed Then wskServer.Close
   wskServer.Accept requestID
   lblStatus.Caption = "Connection Status : HOSTING - READY TO SEND FILE"
   
End Sub

'procedure   :SendFile
'inptuts     :The full path & filename of the file to send
'what it does:
'              a)It reads the file into a string variable
'              b)It sends the file information (filename and size)
'                to the other side and it waits for a response.
'              c)If the response is a yes,it sends the file
Private Sub SendFile(ByVal FileName As String)

   Dim FileData As String
   Dim ByteData As Byte
   Dim Counter As Long
   
   Open FileName For Binary As #1
   
   ProgressBar1.Max = LOF(1)
   ProgressBar1.Value = 0
   
   lblTransferStatus.Caption = "Reading file into memmory...Please be patient..."
   DoEvents
   'Read the file into the variable FileData
   FileData = Input(LOF(1), 1)
   lblTransferStatus.Caption = "Initiating file transfer..."
   
   Close
   
   SendingFile = False
   AbortFile = False
   
   If MsgBox(FileTitle(FileName) & " (" & Len(FileData) & " bytes)" & vbCrLf & _
   "Begin the file transfer?", vbInformation Or vbYesNo) <> vbYes Then
      Exit Sub
   End If
   wskServer.SendData "S_" & Len(FileData) & "_" & FileTitle(FileName)
   
   'This loop suspends the program until the other side
   Do Until SendingFile Or AbortFile Or DoEvents = 0
      DoEvents
   Loop
   
   lblTransferStatus.Caption = "Sent 0 bytes (0%)"
   
   'This command begins the file transfer.The whole file is stored
   'in the string variable FileData.
   BeginTransfer = Timer
   wskServer.SendData FileData

End Sub

'function  : FileTitle
'inputs    : A string containing a full filename (path & file title)
'returns   : The file title
'example   : FileTitle("c:\windows\desktop\readme.txt")
'            returns "readme.txt"
Private Function FileTitle(ByVal FileName As String) As String

   Dim I As Integer
   Dim Temp As String
   
   'if the string includes a path
   If InStr(FileName, "\") <> 0 Then
      'then begin the proccess of parsing the file title.
      I = Len(FileName)
      Do Until Left(Temp, 1) = "\"
         I = I - 1
         Temp = Mid(FileName, I)
      Loop
      FileTitle = Mid(Temp, 2)
   Else
      'If it's already a file title,just return the same string.
      FileTitle = FileName
   End If

End Function

Private Sub wskServer_DataArrival(ByVal bytesTotal As Long)

   Dim NewData As String
   wskServer.GetData NewData
   
   'The string "R_" means that the other side accepts the transfer.
   If NewData = "R_" Then SendingFile = True
   'The string "N_" means that the other side doesn't accept the transfer.
   If NewData = "N_" Then
       MsgBox "The remote computer refuses to accept the file.", vbInformation
       AbortFile = True
   End If
   
End Sub

Private Sub wskServer_SendComplete()
   'This event is raised every time a send operation is
   'completed.

   'Check if this is the end of the send operation that sent
   'the actual file.
   If SendingFile Then
      lblTransferStatus.Caption = "Transfer Complete"
      SendingFile = False
   End If

End Sub

Private Sub wskServer_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)

   'This event is raised whenever data is being sent.It is
   'used to refresh the progress bar and the status label.
   
   If SendingFile Then
      Dim BytesAlreadySent As Long
      BytesAlreadySent = Val(Mid(lblTransferStatus.Caption, 5, InStr(6, lblTransferStatus.Caption, "b")))
      BytesAlreadySent = BytesAlreadySent + bytesSent
      ProgressBar1.Value = BytesAlreadySent
      lblTransferStatus.Caption = "Sent " & BytesAlreadySent & " bytes (" & Format((BytesAlreadySent * 100) / (BytesAlreadySent + bytesRemaining), "00.0") & "%) - " & Format(BytesAlreadySent / (Timer - BeginTransfer) / 1000, "0.0") & " kbps"
      ProgressBar1.Refresh
      lblTransferStatus.Refresh
   End If
   
End Sub






