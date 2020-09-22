VERSION 5.00
Begin VB.MDIForm frmMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H00808080&
   Caption         =   "ISPYBETA v3.0a"
   ClientHeight    =   7815
   ClientLeft      =   1020
   ClientTop       =   615
   ClientWidth     =   10080
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      BackColor       =   &H00808080&
      Height          =   520
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   10020
      TabIndex        =   8
      Top             =   7035
      Width           =   10080
      Begin VB.Label Label1 
         BackColor       =   &H00808080&
         Caption         =   $"frmMain.frx":0CCA
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   495
         Left            =   30
         TabIndex        =   9
         Top             =   -20
         Width           =   10095
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   10020
      TabIndex        =   4
      Top             =   7560
      Width           =   10080
      Begin VB.Image LittleCam 
         Height          =   240
         Index           =   1
         Left            =   2880
         MouseIcon       =   "frmMain.frx":0E92
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":0FE4
         Top             =   -30
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label KB 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   9120
         TabIndex        =   7
         Top             =   15
         Width           =   75
      End
      Begin VB.Label strBytes 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   7680
         TabIndex        =   6
         Top             =   15
         Width           =   75
      End
      Begin VB.Label Status 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   0
         TabIndex        =   5
         Top             =   15
         Width           =   30
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   10020
      TabIndex        =   0
      Top             =   0
      Width           =   10080
      Begin VB.ComboBox txtAddress 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmMain.frx":112E
         Left            =   3840
         List            =   "frmMain.frx":1130
         TabIndex        =   3
         Top             =   420
         Width           =   3135
      End
      Begin Project1.LaVolpeButton cmdServers 
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1085
         BTNICON         =   "frmMain.frx":1132
         BTYPE           =   3
         TX              =   "Servers"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   13160660
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmMain.frx":14CC
         ALIGN           =   0
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   2
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
         OPTVAL          =   0   'False
         OPTMOD          =   0   'False
         GStart          =   0
         GStop           =   16711680
         GStyle          =   0
      End
      Begin Project1.LaVolpeButton cmdFavorites 
         Height          =   615
         Left            =   1440
         TabIndex        =   2
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1085
         BTNICON         =   "frmMain.frx":14E8
         BTYPE           =   3
         TX              =   "Favorites"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   13160660
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmMain.frx":1882
         ALIGN           =   0
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   2
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
         OPTVAL          =   0   'False
         OPTMOD          =   0   'False
         GStart          =   0
         GStop           =   16711680
         GStyle          =   0
      End
      Begin VB.Image Image2 
         Height          =   465
         Left            =   3760
         Picture         =   "frmMain.frx":189E
         Top             =   90
         Width           =   2625
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   3240
         Picture         =   "frmMain.frx":201E
         Top             =   260
         Width           =   480
      End
      Begin VB.Image imgICO 
         Height          =   825
         Left            =   7680
         Picture         =   "frmMain.frx":2CE8
         Top             =   0
         Width           =   2505
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===================
'ISPY beta 3 Written by SNiiP3R
'March 03
'in VB6
'===================

Private Sub cmdFavorites_Click()
With frmFavorites
    .Show
    .ZOrder vbBringToFront
End With
End Sub

Private Sub cmdServers_Click()
With frmServers
    .Show
    .ZOrder vbBringToFront
End With
End Sub

Private Sub LittleCam_Click(Index As Integer)
On Error Resume Next
Dim fcam As Form
For Each fcam In Forms
If fcam.Tag = Index Then fcam.ZOrder vbBringToFront
Next
End Sub

Private Sub MDIForm_Load()

With Me
  .Height = 8220
  .Width = 10200
End With

CAMNUM = 0
GlobalUser = ""
End Sub

Private Sub MDIForm_Resize()
On Error Resume Next

imgICO.Left = Me.Width - 2900
strBytes.Left = Me.Width - 1900
KB.Left = Me.Width - 700

If Me.Width < 10200 Then: Me.Width = 10200
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Unload frmFavorites
Unload frmServers
Unload resources
Unload frmCam
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim strAddress() As String

If KeyAscii = 13 Then
   resources.Username.Text = ""
   StartCAM Trim(txtAddress.Text)
   txtAddress.AddItem Trim(txtAddress.Text)
   RemoveDoubles
KeyAscii = 0

End If
End Sub

'//// THIS Function Opens our CAMS & Connects to server.
Public Function StartCAM(Address As String)
On Error Resume Next
Dim strAddress() As String
Dim LittleCamera As Integer
Dim fcam As Form
Set fcam = New frmCam
For Each fcam In Forms
If fcam.Caption = Address Then GoTo Current
Next
Set fcam = New frmCam
Load LittleCam(LittleCam.UBound + 1)
strAddress = Split(Address, ":")
CAMNUM = CAMNUM + 1
With fcam
   .Username = resources.Username.Text
   .Caption = Address ' Address , Port
   .CamConnect strAddress(0), strAddress(1)
   .Show
   .Tag = CAMNUM
End With

LittleCamera = CAMNUM
LittleCamera = LittleCamera - 1

If fcam.Tag = 1 Then
Dim i As Integer
For i = 1 To LittleCam.Count - 1
           Unload LittleCam(LittleCam.UBound)
    DoEvents
Next
LittleCam(1).Visible = True
LittleCam(1).ToolTipText = "CAM1"
Else
Load LittleCam(LittleCam.UBound + 1)
With LittleCam(CAMNUM)
    .Left = LittleCam(LittleCamera).Left + 300
    .ToolTipText = "CAM" & CAMNUM
    .Visible = True
End With
End If
Current:

End Function

Public Function RemoveDoubles()
Dim i As Long
Dim J As Long
Dim Current As String
For i = txtAddress.ListCount - 1 To 0 Step -1
  Current = txtAddress.List(i)
  For J = txtAddress.ListCount - 1 To 0 Step -1
    If J <> i And txtAddress.List(J) = Current Then
      txtAddress.RemoveItem J
    End If
  Next J
Next i
End Function
