VERSION 5.00
Begin VB.MDIForm s 
   BackColor       =   &H8000000C&
   Caption         =   "s"
   ClientHeight    =   3015
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   4560
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   0
      Picture         =   "MDIForm1.frx":0000
      ScaleHeight     =   1275
      ScaleWidth      =   20310
      TabIndex        =   0
      Top             =   0
      Width           =   20370
      Begin VB.Timer Timer1 
         Left            =   19320
         Top             =   840
      End
      Begin VB.CommandButton Supplier 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Supplier"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Index           =   2
         Left            =   1320
         Picture         =   "MDIForm1.frx":13423
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   1335
      End
      Begin VB.CommandButton Purchase 
         BackColor       =   &H8000000E&
         Caption         =   "Purchase"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   0
         Picture         =   "MDIForm1.frx":142B2
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H8000000E&
         Caption         =   "Customer"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   5280
         Picture         =   "MDIForm1.frx":14B66
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton but 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Stock In"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Index           =   5
         Left            =   3960
         Picture         =   "MDIForm1.frx":153AC
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   1335
      End
      Begin VB.CommandButton Sales 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sales"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Index           =   4
         Left            =   6600
         Picture         =   "MDIForm1.frx":1624E
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   1335
      End
      Begin VB.CommandButton report 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reports"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Index           =   6
         Left            =   7920
         Picture         =   "MDIForm1.frx":16FFA
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   1335
      End
      Begin VB.CommandButton but1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Index           =   0
         Left            =   2640
         Picture         =   "MDIForm1.frx":18133
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "             DRY FRUITS DISTRIBUTOR MANAGEMENT SYSTEM"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10920
         TabIndex        =   8
         Top             =   240
         Width           =   9255
      End
   End
   Begin VB.Menu Masters 
      Caption         =   "Masters"
   End
   Begin VB.Menu Sale_Transaction 
      Caption         =   "Sale_Transaction"
   End
   Begin VB.Menu Purchase_Transaction 
      Caption         =   "Purchase_Transaction"
   End
End
Attribute VB_Name = "s"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub but_Click(Index As Integer)
Form1.Show
End Sub

Private Sub Command1_Click()

Form2.Show
End Sub

