VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form splash 
   Caption         =   "splash screen"
   ClientHeight    =   9225
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19035
   BeginProperty Font 
      Name            =   "Harlow Solid Italic"
      Size            =   36
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   14400
      Top             =   5280
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   8640
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Height          =   7815
      Left            =   3600
      Top             =   2160
      Width           =   12615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   17640
      TabIndex        =   4
      Top             =   10200
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   735
      Left            =   10200
      TabIndex        =   3
      Top             =   7800
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading...."
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   735
      Left            =   7920
      TabIndex        =   2
      Top             =   7800
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Image Image2 
      Height          =   1890
      Left            =   7680
      Picture         =   "splash.frx":0000
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   4440
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Dry Fruits Distributor Management System"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1935
      Left            =   6600
      TabIndex        =   0
      Top             =   2640
      Width           =   7215
   End
   Begin VB.Image Image1 
      Height          =   10935
      Left            =   0
      Picture         =   "splash.frx":256E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20295
   End
End
Attribute VB_Name = "splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image2_Click()
Timer1.Enabled = True
Image2.Visible = False

End Sub

Private Sub Timer1_Timer()
ProgressBar1.Visible = True
ProgressBar1.Value = ProgressBar1.Value + 10
Label2.Visible = True
Label3.Visible = True

Label3.Caption = ProgressBar1.Value & "%"
If (ProgressBar1.Value = ProgressBar1.Max) Then
Unload Me
login.Show
Timer1.Enabled = False
End If
End Sub
