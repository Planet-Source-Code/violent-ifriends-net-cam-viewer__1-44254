VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmFavorites 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Favorites"
   ClientHeight    =   4425
   ClientLeft      =   1095
   ClientTop       =   1920
   ClientWidth     =   3960
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmFavorites.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4425
   ScaleWidth      =   3960
   Tag             =   "xccxcbv"
   Begin MSComctlLib.ListView ListView1 
      Height          =   4430
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3960
      _ExtentX        =   6985
      _ExtentY        =   7805
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   0
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Username"
         Object.Width           =   3351
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Server IP Information"
         Object.Width           =   3528
      EndProperty
   End
End
Attribute VB_Name = "frmFavorites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShowScrollBar Lib "user32" _
(ByVal hwnd As Long, ByVal wBar As Long, _
ByVal bShow As Long) As Long
Dim Fav As Node

Public Function AddIP(Name As String, Server As String)
On Error Resume Next
Dim strList_I As String

strList_I = Name & "'" & Server
If strList_I = "" Then Exit Function
If InStr(strList_I, "'") Then
ListView1.ListItems.Add , , Split(strList_I, "'")(0)
ListView1.ListItems.Item(ListView1.ListItems.Count).ListSubItems.Add , , Split(strList_I, "'")(1)
ListView1.ListItems.Item(ListView1.ListItems.Count).ListSubItems.Add , , Split(strList_I, "'")(2)

Else
Exit Function
End If
RemoveLVDuplicate ListView1
SaveLW ListView1, App.Path & "\Favorites.ifr"
End Function
Private Sub RemoveLVDuplicate(LV As ListView)
Dim fLITem As ListItem
For y = LV.ListItems.Count To 1 Step -1
    Set fLITem = LV.ListItems(y)
    For X = LV.ListItems.Count To fLITem.Index + 1 Step -1
        If fLITem.Text = LV.ListItems(X).Text Then
            LV.ListItems.Remove X
        End If
    Next
Next
End Sub
Private Sub Form_Load()
LoadLW ListView1, App.Path & "\Favorites.ifr"
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
Me.Hide
End Sub

Private Sub ListView1_Click()
On Error Resume Next
resources.Username.Text = ListView1.SelectedItem.Text
End Sub

Private Sub ListView1_DblClick()
On Error Resume Next
frmMain.StartCAM ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(1).Text
With frmMain.txtAddress
    .AddItem ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(1).Text
    .Text = ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(1).Text
End With
frmMain.RemoveDoubles
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 2 Then
resources.PopupMenu resources.Fav
End If
End Sub
