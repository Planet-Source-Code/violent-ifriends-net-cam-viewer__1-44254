VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCam 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5460
   ClientLeft      =   3435
   ClientTop       =   2715
   ClientWidth     =   4785
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5460
   ScaleWidth      =   4785
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3960
      ScaleHeight     =   255
      ScaleWidth      =   795
      TabIndex        =   12
      Top             =   5160
      Width           =   830
      Begin VB.CheckBox chkPrivate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Private"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   20
         MouseIcon       =   "Form1.frx":0CCA
         MousePointer    =   99  'Custom
         TabIndex        =   13
         ToolTipText     =   "Check this box if you want to talk in Private."
         Top             =   -20
         Width           =   855
      End
   End
   Begin RichTextLib.RichTextBox txtRoom 
      Height          =   1320
      Left            =   0
      TabIndex        =   11
      Top             =   3845
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   2328
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Form1.frx":0E1C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Username 
      Height          =   285
      Left            =   6120
      TabIndex        =   9
      Top             =   1560
      Width           =   1815
   End
   Begin MSWinsockLib.Winsock Winsock5 
      Left            =   7560
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox ChatINFO1 
      Height          =   285
      Left            =   6600
      TabIndex        =   8
      Text            =   "<PARAM name=""timeseq"" value="""
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox ValueChar 
      Height          =   285
      Left            =   6360
      TabIndex        =   7
      Text            =   """><input type=""submit"" value="" Begin Guest Chat"
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox ValueHOST 
      Height          =   285
      Left            =   6120
      TabIndex        =   6
      Text            =   """CUSTSCREENNAME"" value=""""><input type=""hidden"" name=""recordcode"" value="""
      Top             =   1200
      Width           =   255
   End
   Begin MSWinsockLib.Winsock Winsock4 
      Left            =   7080
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock3 
      Left            =   6600
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   6120
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtChat 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Top             =   5160
      Width           =   3945
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   -20
      ScaleHeight     =   225
      ScaleWidth      =   4770
      TabIndex        =   1
      Top             =   3590
      Width           =   4800
      Begin VB.Image Image2 
         Height          =   240
         Left            =   0
         Picture         =   "Form1.frx":0E97
         Top             =   -20
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   1080
         Picture         =   "Form1.frx":1221
         Top             =   0
         Width           =   240
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Disconnected"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2520
         TabIndex        =   5
         Top             =   30
         Width           =   2295
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   " Connect to chat"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1320
         MouseIcon       =   "Form1.frx":15AB
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   30
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   " SnapShot"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         MouseIcon       =   "Form1.frx":16FD
         MousePointer    =   99  'Custom
         TabIndex        =   2
         ToolTipText     =   "Click here to take a picture"
         Top             =   30
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   3555
      Left            =   20
      MouseIcon       =   "Form1.frx":184F
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":19A1
      ScaleHeight     =   3555
      ScaleWidth      =   4755
      TabIndex        =   0
      ToolTipText     =   "Add to favorites."
      Top             =   0
      Width           =   4755
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   4080
         Top             =   600
      End
      Begin Project1.ChatBubble Bubble 
         Height          =   1935
         Left            =   600
         TabIndex        =   10
         Top             =   720
         Visible         =   0   'False
         Width           =   3615
         _ExtentX        =   3836
         _ExtentY        =   2355
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Timer tmrGET 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1080
      Top             =   4320
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1560
      Top             =   4320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmCam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Coded by SNiiP3R or some people know me as Violent in Visual Basic 6 in about 2 days :P
'-----------------------------------------------------
'We'll need 5 Winsock Controls
'Winsock1 = Connects to CAM Server ( requests / downloads images )
'Winsock2 = Connects to Girl's profile and grabs SESSID that you will need later
'Winsock3 = Connects to Girls's Chat page and Grabs Access Cookie
'Winsock4 = Connects to Girls Chat Page and grabs ( Chat server address and port )
'Winsock5 = This socket connects to Chat using all the info that other Socks got!
'-------------------------------------------------------
Dim intVAL As Integer
Dim nickVal As Integer
Dim TempBuffer As String
Dim CurrentIP, SESSID, BigString, PORT, ServerIP, seq() As String

'////// THIS FUNCTION Connects us to CAMS ////////////
Public Sub CamConnect(Server As String, PORT As String)
CurrentIP = Server            'Cam number
frmMain.Status.Caption = "CAM" & CAMNUM & " " & "connecting to " & Server
If Winsock1.State <> sckConnected Then
   Winsock1.Connect Server, PORT
Else: Exit Sub
End If
End Sub

Private Sub Bubble_Click()

End Sub

Private Sub Form_Load()
TempBuffer = ""
intVAL = Int((100 * Rnd) + 1) ' This generates random numbers from 1 to 100
End Sub

Private Sub Form_Resize()
If Me.WindowState <> vbMinimized Then
With Me
   .Left = frmMain.Width - 5500
   .Top = frmMain.Picture1.Top + 200
End With
End If
End Sub

Private Sub Label1_Click()
Dim a As String
'//////// SAVES Pictures ( SnapShot )
If Winsock1.State <> sckConnected Then Exit Sub
PlayWav (App.Path & "\TakePic.wav")
a = (Time)
a = Replace(a, ":", ".")
MoveFile App.Path & "\" & "temp" & intVAL & ".jpg", App.Path & "\Saved Gallery\" & "temp" & a & ".jpg"
End Sub

Public Sub MoveFile(Source As String, Destination As String)
On Error Resume Next
    FileCopy Source, Destination
    Exit Sub
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label1.ForeColor = vbWhite
End Sub

Private Sub Label2_Click()
If Winsock5.State = sckConnected Then MsgBox "Your already connected.", vbInformation, ""
If IsNumeric(seq(0)) Then
   Winsock4.Close
   Winsock5.Connect "chat.iFriends.net", PORT
End If
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label2.ForeColor = vbWhite
End Sub

Private Sub Picture1_Click()
frmFavorites.AddIP Trim(Username.Text), Me.Caption
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
With Me
    .ForeColor = vbBlack
    .ForeColor = vbBlack
End With
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
With Me
    .ForeColor = vbBlack
    .ForeColor = vbBlack
End With
End Sub

Private Sub Timer1_Timer()
With Me
   .Enabled = False
   .Visible = False
End With
End Sub

Private Sub txtChat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If txtChat.Text = "" Then Exit Sub
If Winsock5.State = sckConnected Then
Do Until Len(SESSID) = 8
SESSID = SESSID + Space(1)
Loop
If chkPrivate.Value = 1 Then
   With Winsock5                              ' "gUeSt"
      .SendData "@" & SESSID & Username.Text & "gUeSt" & nickVal & "       " & "/" & Trim(Username.Text) & " " & txtChat.Text
      .SendData vbCrLf
   End With
   txtChat.Text = ""
   Else
   With Winsock5                              ' "gUeSt"
      .SendData "@" & SESSID & Username.Text & "gUeSt" & nickVal & "       " & txtChat.Text
      .SendData vbCrLf
   End With
   txtChat.Text = ""
   End If
Else: Exit Sub
End If
End If
End Sub

Private Sub txtRoom_DblClick()
txtRoom.Text = ""
End Sub

Private Sub txtRoom_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label2.ForeColor = vbBlack
Label1.ForeColor = vbBlack
End Sub

Private Sub Winsock2_Connect()
Winsock2.SendData "GET /~wsapi/ifbrowse.dll?type=L&kw=" & Trim(Username.Text) & " " & "HTTP/1.1" & vbCrLf & _
"Accept: */*" & vbCrLf & "Referer: http://www.ifriends.net/livewebcamviewer/if/55/index.htm" & vbCrLf & _
"Accept-Language: en-us" & vbCrLf & "Accept-Encoding: gzip, deflate" & vbCrLf & _
"User-Agent: Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; .NET CLR 1.0.3705)" & vbCrLf & _
"Host: apps7.ifriends.net" & vbCrLf & "Connection: Keep-Aliv" & vbCrLf & vbCrLf
Label3.Caption = "Connected. Requesting Profile"
End Sub

Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim strID As String

Winsock2.PeekData strID

SESSID = GrabItBetween(strID, "http://access.iFriends.net/cgi/iReqFeed.exe?screenname=" & Trim(Username.Text) & "&sessionid=", Chr(34) & ">Guest Chat</A>")
Label3.Caption = "Getting Session ID"
If Len(SESSID) <> 0 And IsNumeric(SESSID) Then
   Winsock2.Close
   Winsock3.Connect "205.246.203.39", 80
   Label3.Caption = "Connecting to server."
End If

End Sub
Private Sub Winsock3_Connect()
Label3.Caption = "Requesting access cookie."
GetHeaderF "cgi/iReqFeed.exe?screenname=" & Trim(Username.Text) & "&sessionid=" & SESSID, "http://apps7.ifriends.net/~wsapi/ifbrowse.dll?type=L&kw=" & Trim(Username.Text) & "&_dummyduh=true", "access.ifriends.net", Winsock3
End Sub
Private Sub Winsock3_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim strACCESS As String

Winsock3.PeekData strACCESS

strACCESS = Replace(strACCESS, Chr(10), "")
strACCESS = Replace(strACCESS, Chr(13), "")

BigString = GrabItBetween(strACCESS, ValueHOST.Text, ValueChar.Text)

If Mid$(BigString, 1, 7) = "lhost__" Then
   Winsock3.Close
   Winsock4.Connect "205.246.203.39", 80
   Label3.Caption = "Connecting to server."
End If

End Sub
Private Sub Winsock4_Connect()
PostHeader "/cgi/showcam.exe?", "http://access.ifriends.net/cgi/iReqFeed.exe?screenname=" & Trim(Username.Text) & "&sessionid=" & SESSID, "access.ifriends.net", "SCREENNAME=" & Trim(Username.Text) & "&PARM5=AHPLA&SESSIONID=" & SESSID & "&CUSTSESA=0&CUSTSCREENNAME=&recordcode=" & BigString, Winsock4
Label3.Caption = "Requesting server ip and port."
End Sub

Private Sub Winsock4_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim strChatINFO As String
Dim prLoc As Integer
Dim tmLoc As Integer

Winsock4.PeekData strChatINFO

prLoc = InStr(1, strChatINFO, Chr(34) & "port" & Chr(34) & " value=")
tmLoc = InStr(1, strChatINFO, ChatINFO1.Text)
PORT = Mid(strChatINFO, prLoc + 13, 4)
seq = Split(Mid(strChatINFO, tmLoc + Len(ChatINFO1), 100), Chr(34))

If IsNumeric(seq(0)) Then
   Label3.Caption = "Recieved chat ID(s)."
   With Bubble
      .Add Username.Text, SESSID, seq(0), PORT, BigString
      .Visible = True
   End With
   Label2.Enabled = True
   Timer1.Enabled = True
End If
End Sub
Private Function GetProperLenght(Username As String)
'Neeeded for ifriends Chat , without this function u wont be able to connect to some chats
Dim TempHolder As String
TempHolder = Username
If Len(TempHolder) = 0 Then Exit Function
Do Until Len(TempHolder) = 15
TempHolder = TempHolder + Space(1)
Loop
GetProperLenght = TempHolder
End Function
Private Sub Winsock5_Connect()
Username.Text = GetProperLenght(Username.Text)
Do Until Len(SESSID) = 8
SESSID = SESSID + Space(1)
Loop
Winsock5.SendData "!" & SESSID & Username.Text & "GUEST" & nickVal & "       " & seq(0) & vbCrLf
Label3.Caption = "Connected."
End Sub

Private Sub Winsock5_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim strChat As String

Winsock5.GetData strChat
txtRoom.Text = txtRoom.Text + strChat
txtRoom.SelStart = Len(txtRoom.Text)

End Sub

Private Sub tmrGET_Timer()
Winsock1.SendData "GET /java.jpg HTTP/1.0  Host: " & CurrentIP & vbCrLf & vbCrLf
tmrGET.Enabled = False
End Sub

Private Sub Winsock1_Connect()
Winsock1.SendData "GET /speed.ifg HTTP/1.0  Host: " & CurrentIP & vbCrLf & vbCrLf
tmrGET.Enabled = True
frmMain.Status.Caption = "CAM" & CAMNUM & " is " & "connected. Sending request..."
    Winsock2.Connect "205.246.203.26", 80
    Label3.Caption = "Connecting to Server."
    nickVal = Int((300 * Rnd) + 1) ' This Generates Random #s. For you Chat nickname :P
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim strData As String

Winsock1.GetData strData

TempBuffer = TempBuffer + strData
frmMain.strBytes.Caption = Val(frmMain.strBytes) + bytesTotal & " bytes"
frmMain.KB.Caption = Int(Val(frmMain.strBytes) / 1000) & " k"
 
If Right(TempBuffer, 2) = Chr(&HFF) & Chr(&HD9) Then
Winsock1.SendData "GET /java.jpg HTTP/1.0  Host: " & CurrentIP & vbCrLf & vbCrLf
String2File TempBuffer, App.Path & "\" & "temp" & intVAL & ".jpg"
TempBuffer = ""
frmMain.Status.Caption = "CAM" & CAMNUM & " is " & "connected to server. Enjoy the show."
End If

End Sub

Public Function String2File(ByRef Data As String, ByVal FileName As String)
  
If Dir$(FileName) <> "" Then

End If
  
Dim iFile As Integer
iFile = FreeFile()
  
Open FileName For Binary As #iFile
Put #iFile, 1, Data
Close #iFile

Picture1.Picture = LoadPicture(App.Path & "\" & "temp" & intVAL & ".jpg")

End Function

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Winsock1.Close
Winsock2.Close
Winsock3.Close
Winsock4.Close
Winsock5.Close
Kill (App.Path & "\" & "temp" & intVAL & ".jpg")
frmMain.Status.Caption = "CAM" & CAMNUM & " : " & "connection Closed"
If Me.Tag = 1 Then
   frmMain.LittleCam(1).Visible = False
Else: Unload frmMain.LittleCam(Me.Tag)
End If
CAMNUM = CAMNUM - 1
If CAMNUM = 0 Then
   frmMain.KB.Caption = "0"
   frmMain.strBytes.Caption = "0"
End If
End Sub

