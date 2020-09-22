VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmServers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Servers"
   ClientHeight    =   4425
   ClientLeft      =   1095
   ClientTop       =   1920
   ClientWidth     =   3960
   BeginProperty Font 
      Name            =   "Small Fonts"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmServers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4425
   ScaleWidth      =   3960
   Begin VB.TextBox String2 
      Height          =   255
      Left            =   4680
      TabIndex        =   4
      Text            =   """ archive=""vvhp.jar"" width=""40"" height=""30"">"
      Top             =   2160
      Width           =   2295
   End
   Begin VB.TextBox String1 
      Height          =   255
      Left            =   4680
      TabIndex        =   3
      Text            =   "<applet code=""vvhp.class"" codebase=""http://"
      Top             =   1800
      Width           =   2295
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   3600
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer tmrSend 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   120
      ScaleHeight     =   240
      ScaleWidth      =   3675
      TabIndex        =   1
      Top             =   3960
      Width           =   3735
      Begin VB.ComboBox strPage 
         Height          =   285
         ItemData        =   "frmServers.frx":038A
         Left            =   2320
         List            =   "frmServers.frx":03AF
         TabIndex        =   8
         Text            =   "5"
         Top             =   -30
         Width           =   550
      End
      Begin VB.ComboBox Combo1 
         Height          =   285
         ItemData        =   "frmServers.frx":03D6
         Left            =   2850
         List            =   "frmServers.frx":03E6
         TabIndex        =   7
         Text            =   "Index"
         Top             =   -30
         Width           =   855
      End
      Begin VB.ComboBox srch 
         Height          =   285
         ItemData        =   "frmServers.frx":040F
         Left            =   1800
         List            =   "frmServers.frx":0437
         TabIndex        =   5
         Text            =   "5"
         Top             =   -30
         Width           =   550
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Reconnect"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         MouseIcon       =   "frmServers.frx":0461
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   30
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         MouseIcon       =   "frmServers.frx":05B3
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   30
         Width           =   615
      End
   End
   Begin MSComctlLib.TreeView ServerList 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Right Click for more options"
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   6376
      _Version        =   393217
      Indentation     =   37
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmServers.frx":0705
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4800
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServers.frx":0867
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServers.frx":0C01
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   " Room:  Page:    Cams:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      Top             =   3750
      Width           =   1935
   End
End
Attribute VB_Name = "frmServers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShowScrollBar Lib "user32" _
(ByVal hwnd As Long, ByVal wBar As Long, _
ByVal bShow As Long) As Long
Dim nodex As Node


Private Sub Combo1_Click()
Select Case Combo1.ListIndex

Case "0"
    srch.Text = "5"
Case "1"
    srch.Text = "10"
Case "2"
    srch.Text = "7"
Case "3"
    srch.Text = "9"

End Select
End Sub


Private Sub Form_Load()
On Error Resume Next
Dim z As Integer
If Winsock1.State <> sckClosed Then Winsock1.Close
   Winsock1.Connect "65.212.89.51", 8080
nodex = ServerList.Nodes.Add(, , "S", "Servers Found", 1)
z = "1"
strPage.Text = "55"
Do Until Val(z) = 99
   strPage.AddItem z
   z = z + 1
Loop
End Sub

Private Sub Form_Resize()
If Me.WindowState <> vbMinimized Then
With Me
  .Left = 50
  .Top = 20
End With
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Winsock1.Close
Me.Hide
End Sub

Private Sub Label1_Click()
If tmrSend.Enabled = False Then
   Label1.Caption = "Stop"
     tmrSend.Enabled = True
     Else
    Label1.Caption = "Update"
  tmrSend.Enabled = False
End If
End Sub

Public Sub Label3_Click()
If Winsock1.State <> sckClosed Then Winsock1.Close
Winsock1.Connect "65.212.89.51", 8080
End Sub

Private Sub ServerList_DblClick()
On Error Resume Next
Dim X As String
Dim TempServer As String
ShowScrollBar ServerList.hwnd, 0, False
tmrSend.Enabled = False

Label1.Caption = "Update"
resources.Username.Text = ServerList.SelectedItem.Text
If InStr(1, ServerList.SelectedItem, "Servers Found") Then Exit Sub

X = Inet1.OpenURL("http://205.246.203.22/~wsapi/VCHDriveway.dll?screenname=" & ServerList.SelectedItem)
If InStr(X, "<html><HEAD><TITLE>iFriends Live Browse</TITLE></HEAD>") Then
MsgBox "Cam was not found. Please try another server.", vbInformation, "Not found"
Exit Sub
End If
TempServer = GrabItBetween(X, String1, String2)

If TempServer <> "" Then
frmMain.StartCAM TempServer
Else: Exit Sub
End If
X = ""

With frmMain.txtAddress
    .AddItem TempServer
    .Text = TempServer
End With
frmMain.RemoveDoubles
End Sub

Private Sub ServerList_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 2 Then
resources.PopupMenu resources.Menu
End If
End Sub

Private Sub Winsock1_Connect()
Me.Caption = "Connected"
Label3.Enabled = False
Label1.Enabled = True
End Sub
Private Sub tmrSend_Timer()
If Winsock1.State = sckConnected Then
Select Case Combo1.Text

Case "Index"
   '"5"
   Winsock1.SendData "GET /vch.info?room=" & srch.Text & "&doc=http://www.ifriends.net/livewebcamviewer/if/" & strPage & "/index.htm HTTP/1.0" & vbCrLf
Case "Hot Couples"
   '"9"
   Winsock1.SendData "GET /vch.info?room=" & srch.Text & "&doc=http://www.ifriends.net/livewebcamviewer/if/" & strPage & "/cpl.htm HTTP/1.0" & vbCrLf
Case "Group"
   '"10"
   Winsock1.SendData "GET /vch.info?room=" & srch.Text & "&doc=http://www.ifriends.net/livewebcamviewer/if/" & strPage & "/grp.htm HTTP/1.0" & vbCrLf
Case "Lesbians"
   '"7"
   Winsock1.SendData "GET /vch.info?room=" & srch.Text & "&doc=http://www.ifriends.net/livewebcamviewer/if/" & strPage & "/les.htm HTTP/1.0" & vbCrLf
End Select

Else
Label3.Enabled = True
Label1.Enabled = False
tmrSend.Enabled = False
Label1.Caption = "Update"
Me.Caption = "Disconnected"
End If
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim strData As String
Dim Names() As String
Winsock1.GetData strData

Names = Split(strData, "*")
nodex = ServerList.Nodes.Add("S", tvwChild, Names(0), Names(0), 2)
nodex.EnsureVisible
ServerList.Nodes.Item(1).Text = "Servers Found" & " ( " & ServerList.Nodes.Count - 1 & " ) "

End Sub

