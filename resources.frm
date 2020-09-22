VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form resources 
   BorderStyle     =   0  'None
   ClientHeight    =   375
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   1260
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   375
   ScaleWidth      =   1260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   360
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox Username 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   2055
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Begin VB.Menu mnuAdd 
         Caption         =   "Add to favorites"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "Remove server"
      End
      Begin VB.Menu mnuReconnect 
         Caption         =   "Reconnect"
      End
      Begin VB.Menu SP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConnect 
         Caption         =   "Connect to server"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu Fav 
      Caption         =   "Fav"
      Begin VB.Menu mnuRemoveServer 
         Caption         =   "Remove server"
      End
      Begin VB.Menu SEP05 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close Favorites"
      End
   End
End
Attribute VB_Name = "resources"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub mnuAdd_Click()
On Error Resume Next
Dim x As String
Dim nodem As Node
Dim TempServer As String
Dim strName As String
If InStr(1, frmServers.ServerList.SelectedItem.Text, "Servers Found") Then Exit Sub
strName = frmServers.ServerList.SelectedItem.Text
x = Inet1.OpenURL("http://205.246.203.22/~wsapi/VCHDriveway.dll?screenname=" & strName)
If InStr(x, "<html><HEAD><TITLE>iFriends Live Browse</TITLE></HEAD>") Then
MsgBox "Cam was not found. Please try another server.", vbInformation, "Not found"
Exit Sub
End If

TempServer = GrabItBetween(x, frmServers.String1, frmServers.String2)

If TempServer <> "" Then
frmFavorites.AddIP strName, TempServer
Else: Exit Sub
End If
End Sub

Private Sub mnuClose_Click()
Unload frmFavorites
End Sub

Private Sub mnuReconnect_Click()
frmServers.Label3_Click
End Sub

Public Sub mnuRemove_Click()
On Error Resume Next
If frmServers.ServerList.SelectedItem = InStr(1, "Servers Found") Then: Exit Sub
frmServers.ServerList.Nodes.Remove (frmServers.ServerList.SelectedItem.Index)
End Sub

Private Sub mnuRemoveServer_Click()
On Error Resume Next
frmFavorites.ListView1.ListItems.Remove (frmFavorites.ListView1.SelectedItem.Index)
SaveLW frmFavorites.ListView1, App.Path & "\Favorites.ifr"
End Sub
