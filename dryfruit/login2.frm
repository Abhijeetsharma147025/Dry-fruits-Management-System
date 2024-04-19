VERSION 5.00
Begin VB.Form login 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Login"
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20250
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   22.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "show password "
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10560
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   7
      Top             =   7920
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8880
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   10560
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   7200
      Width           =   4815
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10680
      TabIndex        =   0
      Top             =   5280
      Width           =   4695
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "forgot password?"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13200
      TabIndex        =   6
      Top             =   7920
      Width           =   2250
   End
   Begin VB.Image Image3 
      Height          =   3240
      Left            =   3360
      Picture         =   "login2.frx":0000
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   5055
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WELCOME!"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   945
      Left            =   10680
      TabIndex        =   4
      Top             =   3240
      Width           =   4215
   End
   Begin VB.Image Image2 
      Height          =   1710
      Left            =   11880
      Picture         =   "login2.frx":ADDF
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1830
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD:"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   23.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   10560
      TabIndex        =   3
      Top             =   6360
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "USER ID:"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   23.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   10560
      TabIndex        =   2
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   11055
      Left            =   0
      Picture         =   "login2.frx":13014
      Stretch         =   -1  'True
      Top             =   0
      Width           =   17775
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
If (Text2.PasswordChar = "*") Then
Text2.PasswordChar = ""
Else
Text2.PasswordChar = "*"
End If
End Sub

Private Sub Command1_Click()
Set R = New ADODB.Recordset
SQL = "SELECT * FROM LOGIN"
Set R = C.Execute(SQL)
If (Text1.Text = R.Fields(0) And Text2.Text = R.Fields(1)) Then
mdi.Show
Unload Me
Else
MsgBox "PLEASE ENTER CORRECT USER ID AND PASSWORD"
End If
End Sub

Private Sub Form_Load()
CONN
Command1.Enabled = False
End Sub

Private Sub Label4_Click()
forget.Show
Unload Me
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text2.SetFocus
End Sub

Private Sub Text1_Change()
Command1.Enabled = Len(Text1.Text) > 0

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command1.SetFocus
End Sub



