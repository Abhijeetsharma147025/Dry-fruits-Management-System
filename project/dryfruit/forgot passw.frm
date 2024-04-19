VERSION 5.00
Begin VB.Form resetpasswrd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reset Password"
   ClientHeight    =   9210
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7830
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9210
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   6360
      Width           =   3975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8415
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   6735
      Begin VB.CommandButton update 
         BackColor       =   &H00FFFF00&
         Caption         =   "UPDATE"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   7440
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   4080
         Width           =   3975
      End
      Begin VB.Image Image1 
         Height          =   1575
         Left            =   2040
         Picture         =   "forgot passw.frx":0000
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reset Password"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1200
         TabIndex        =   5
         Top             =   240
         Width           =   4785
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm New Password:"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   3
         Top             =   5400
         Width           =   3975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter New Password:"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1320
         TabIndex        =   1
         Top             =   3480
         Width           =   3735
      End
   End
   Begin VB.Image Image2 
      DataMember      =   "&H00808080&"
      Height          =   11655
      Left            =   -3600
      Picture         =   "forgot passw.frx":4AF2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20325
   End
End
Attribute VB_Name = "resetpasswrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub update_Click()
Set R = New ADODB.Recordset
SQL = "SELECT * FROM LOGIN"
Set R = C.Execute(SQL)
If (Text1.Text = Text2.Text) Then
SQL = "update login set password='" + Text1.Text + "' where userid='admin'"
Set R = C.Execute(SQL)
MsgBox "pass update"
Unload Me
login.Show
Else
MsgBox "pass not matched"
Text1.Text = " "
Text2.Text = " "
End If
End Sub

Private Sub Form_Load()
CONN
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text2.SetFocus
End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then update.SetFocus
End Sub

