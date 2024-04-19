VERSION 5.00
Begin VB.Form Replace_user 
   Caption         =   "Form1"
   ClientHeight    =   8280
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   7425
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00404000&
      BorderStyle     =   0  'None
      Height          =   9135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      Begin VB.TextBox Text3 
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
         MaxLength       =   10
         TabIndex        =   9
         Top             =   5640
         Width           =   5535
      End
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
         Left            =   960
         TabIndex        =   7
         Top             =   3840
         Width           =   5535
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
         Top             =   2160
         Width           =   5535
      End
      Begin VB.CommandButton command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Replace User"
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
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   6840
         Width           =   2415
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cancel"
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
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   6840
         Width           =   2055
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Phone No."
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   960
         TabIndex        =   8
         Top             =   5040
         Width           =   3060
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Replace Existing User"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   23.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   615
         Left            =   1200
         TabIndex        =   6
         Top             =   120
         Width           =   5535
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
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   480
         TabIndex        =   5
         Top             =   1560
         Width           =   4095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Password"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   960
         TabIndex        =   4
         Top             =   3240
         Width           =   2925
      End
   End
End
Attribute VB_Name = "Replace_user"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Set R = New ADODB.Recordset
SQL = "update login set userid='" + Text1.Text + "',password='" + Text2.Text + "',phone='" + Text3.Text + "'"
Set R = C.Execute(SQL)
MsgBox "User Replaced!!"
login.Show
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
CONN
MsgBox "Connected!"
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text2.SetFocus
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text3.SetFocus
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command1.SetFocus
End Sub
