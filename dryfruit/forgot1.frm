VERSION 5.00
Begin VB.Form forget 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "forget password?"
   ClientHeight    =   10755
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10755
   ScaleWidth      =   9495
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
      Left            =   1920
      TabIndex        =   7
      Top             =   6600
      Width           =   5535
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   9135
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   7455
      Begin VB.CommandButton Command2 
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   3960
         TabIndex        =   6
         Top             =   7440
         Width           =   1575
      End
      Begin VB.CommandButton resetpass 
         Caption         =   "Reset password"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1320
         TabIndex        =   5
         Top             =   7440
         Width           =   1695
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
         Left            =   960
         TabIndex        =   3
         Top             =   3960
         Width           =   5535
      End
      Begin VB.Image Image2 
         Height          =   1815
         Left            =   2640
         Picture         =   "forgot1.frx":0000
         Stretch         =   -1  'True
         Top             =   960
         Width           =   1680
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Registered Phone number"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   840
         TabIndex        =   4
         Top             =   5160
         Width           =   4950
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "     Enter User Id"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         TabIndex        =   2
         Top             =   3360
         Width           =   4095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "FORGOT PASSWORD??"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   23.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   720
         TabIndex        =   1
         Top             =   120
         Width           =   5535
      End
   End
   Begin VB.Image Image1 
      Height          =   11655
      Left            =   0
      Picture         =   "forgot1.frx":E2A6
      Stretch         =   -1  'True
      Top             =   -360
      Width           =   17640
   End
End
Attribute VB_Name = "forget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub resetpass_Click()
Set R = New ADODB.Recordset
SQL = "SELECT * FROM LOGIN"
Set R = C.Execute(SQL)
If (Text1.Text = R.Fields(0) And Text2.Text = R.Fields(2)) Then
resetpasswrd.Show
Unload Me
Else
MsgBox "PLEASE ENTER CORRECT USER ID AND PASSWORD"
End If
End Sub

Private Sub Form_Load()
CONN
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text2.SetFocus
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then resetpass.SetFocus
End Sub

