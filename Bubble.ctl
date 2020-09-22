VERSION 5.00
Begin VB.UserControl ChatBubble 
   BackColor       =   &H00C0C0C0&
   BackStyle       =   0  'Transparent
   ClientHeight    =   1920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3570
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   128
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   238
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Height          =   165
      Left            =   840
      TabIndex        =   13
      Top             =   1560
      Width           =   2565
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   165
      Left            =   840
      TabIndex        =   12
      Top             =   1560
      Width           =   45
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "CookieID:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label AccessIDLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   165
      Left            =   840
      TabIndex        =   10
      Top             =   1080
      Width           =   45
   End
   Begin VB.Label SESSIDlabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   165
      Left            =   840
      TabIndex        =   9
      Top             =   840
      Width           =   45
   End
   Begin VB.Label PortLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   165
      Left            =   1920
      TabIndex        =   8
      Top             =   1320
      Width           =   45
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "chat.ifriends.net:"
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label UserLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   165
      Left            =   840
      TabIndex        =   6
      Top             =   600
      Width           =   165
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "ServerIP:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "AccessID:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Connect to chat now."
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Recieved current Chat information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SESSID :"
      Height          =   165
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "USER:"
      ForeColor       =   &H00000000&
      Height          =   165
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   600
      Width           =   420
   End
   Begin VB.Shape Bubble 
      BorderColor     =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   1905
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   3570
   End
End
Attribute VB_Name = "ChatBubble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Default Property Values:
Const m_def_BackColor = 0
Const m_def_ForeColor = 0
Const m_def_Enabled = 0
Const m_def_BackStyle = 0
Const m_def_BorderStyle = 0
'Property Variables:
Dim m_BackColor As Long
Dim m_ForeColor As Long
Dim m_Enabled As Boolean
Dim m_Font As Font
Dim m_BackStyle As Integer
Dim m_BorderStyle As Integer
'Event Declarations:
Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal color As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function BitBlt Lib "gdi32" ( _
ByVal hDestDC As Long, _
ByVal y As Long, _
ByVal nWidth As Long, _
ByVal nWidth As Long, _
ByVal nHeight As Long, _
ByVal hSrcDC As Long, _
ByVal xSrc As Long, _
ByVal ySrc As Long, _
ByVal dwRop As Long) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get BackColor() As Long
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As Long)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get ForeColor() As Long
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As Long)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=5
'
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=14
Public Function Add(ByVal User As String, ByVal SESSID As String, ByVal AccessID As String, ByVal PORT As String, ByVal cookieID As String)
UserLabel.Caption = User
SESSIDlabel.Caption = SESSID
AccessIDLabel.Caption = AccessID
PortLabel.Caption = PORT
Label10.Caption = cookieID
End Function


'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    m_Enabled = m_def_Enabled
    Set m_Font = Ambient.Font
    m_BackStyle = m_def_BackStyle
    m_BorderStyle = m_def_BorderStyle
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
End Sub

Private Function Strip(ByVal s As String, Length As Integer) As String
    Dim i As Integer
    Strip = ""
    If InStr(s, " ") <> 0 Then
    Do
        If Len(s) > Length Then
            i = Length
            Do While Mid$(s, i, 1) <> " "
                i = i - 1
                If i = 0 Then Exit Do
            Loop
        Else
            i = Length
        End If
        
        Strip = Strip & " " & Left$(s, i) & " " & vbNewLine
        s = Mid$(s, i + 1)
    Loop Until i = 0 Or Len(s) = 0
    
    Else
    Strip = " " & s & " " & vbNewLine
End If
End Function
Private Sub RePaintPoint()
Dim hWndDesk As Long, hDCDesk As Long
imgPoint.Cls
imgPoint.Visible = False
DoEvents
hWndDesk = GetDesktopWindow()
hDCDesk = GetDC(hWndDesk)
imgPoint.Visible = True
Dim MainL As RECT
GetWindowRect imgPoint.hwnd, MainL
Call BitBlt(pctDesktop.hdc, 0, 0, Screen.Width, Screen.Height, hDCDesk, MainL.Left, MainL.Top, &HCC0020)
For x = 0 To imgPoint.Width
 For y = 0 To imgPoint.Height
  If GetPixel(imgPoint.hdc, x, y) = vbBlue Then
    Call BitBlt(imgPoint.hdc, x, y, 1, 1, pctDesktop.hdc, x, y, &HCC0020)
    End If
   Next
 Next
Call ReleaseDC(hWndDesk, hDCDesk)
End Sub


