VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form order 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Order"
   ClientHeight    =   8340
   ClientLeft      =   1245
   ClientTop       =   -585
   ClientWidth     =   20460
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   20460
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text4 
      BackColor       =   &H8000000B&
      Height          =   435
      Left            =   11760
      MaxLength       =   6
      TabIndex        =   91
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C000&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   17880
      Style           =   1  'Graphical
      TabIndex        =   82
      Top             =   7200
      Width           =   2055
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   9735
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   20415
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   2370
         Left            =   17640
         TabIndex        =   99
         Top             =   480
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   222625794
         CurrentDate     =   44440
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   240
         TabIndex        =   97
         Top             =   2640
         Width           =   2175
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C000&
         Caption         =   "Place Order"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   17880
         Style           =   1  'Graphical
         TabIndex        =   95
         Top             =   6720
         Width           =   2055
      End
      Begin VB.TextBox Text6 
         Height          =   495
         Left            =   2520
         TabIndex        =   94
         Top             =   8520
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   405
         Left            =   15960
         TabIndex        =   90
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   405
         Left            =   14640
         Locked          =   -1  'True
         TabIndex        =   89
         Top             =   2520
         Width           =   975
      End
      Begin VB.ListBox unit 
         BackColor       =   &H00FFFFC0&
         Height          =   2205
         ItemData        =   "purchase.frx":0000
         Left            =   8640
         List            =   "purchase.frx":0002
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000B&
         Height          =   435
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   2520
         Width           =   1215
      End
      Begin VB.ComboBox Combo2 
         BackColor       =   &H8000000B&
         DataField       =   "MED_NM"
         Height          =   315
         Left            =   2520
         TabIndex        =   83
         Top             =   2640
         Width           =   2655
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C000&
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   18120
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C000&
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   18120
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   3240
         Width           =   1215
      End
      Begin VB.ComboBox TXT41 
         BackColor       =   &H8000000B&
         DataField       =   "MED_TYPE"
         Height          =   315
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   2640
         Width           =   2415
      End
      Begin VB.TextBox TXT1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox TXT2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   14520
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   360
         Width           =   2655
      End
      Begin VB.TextBox TXT5 
         BackColor       =   &H8000000B&
         Height          =   435
         Left            =   13200
         MaxLength       =   4
         TabIndex        =   24
         Top             =   2520
         Width           =   1335
      End
      Begin VB.ListBox sr 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   2175
         ItemData        =   "purchase.frx":0004
         Left            =   240
         List            =   "purchase.frx":0006
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   3720
         Width           =   615
      End
      Begin VB.ListBox id 
         BackColor       =   &H00FFFFC0&
         Height          =   2205
         ItemData        =   "purchase.frx":0008
         Left            =   840
         List            =   "purchase.frx":000A
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   3720
         Width           =   1215
      End
      Begin VB.ListBox prc 
         BackColor       =   &H00FFFFC0&
         Height          =   2205
         ItemData        =   "purchase.frx":000C
         Left            =   11040
         List            =   "purchase.frx":000E
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   3720
         Width           =   1455
      End
      Begin VB.ListBox qty 
         BackColor       =   &H00FFFFC0&
         Height          =   2205
         ItemData        =   "purchase.frx":0010
         Left            =   9840
         List            =   "purchase.frx":0012
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   3720
         Width           =   1215
      End
      Begin VB.ListBox prate 
         BackColor       =   &H00FFFFC0&
         Height          =   2205
         ItemData        =   "purchase.frx":0014
         Left            =   7440
         List            =   "purchase.frx":0016
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   3720
         Width           =   1215
      End
      Begin VB.ListBox typ 
         BackColor       =   &H00FFFFC0&
         Height          =   2205
         ItemData        =   "purchase.frx":0018
         Left            =   5640
         List            =   "purchase.frx":001A
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   3720
         Width           =   1815
      End
      Begin VB.ListBox nm 
         BackColor       =   &H00FFFFC0&
         Height          =   2205
         ItemData        =   "purchase.frx":001C
         Left            =   2040
         List            =   "purchase.frx":001E
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   3720
         Width           =   3615
      End
      Begin VB.ListBox sgst 
         BackColor       =   &H00FFFFC0&
         Height          =   2205
         ItemData        =   "purchase.frx":0020
         Left            =   13920
         List            =   "purchase.frx":0022
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   3720
         Width           =   615
      End
      Begin VB.ListBox cmt 
         BackColor       =   &H00FFFFC0&
         Height          =   2205
         ItemData        =   "purchase.frx":0024
         Left            =   13080
         List            =   "purchase.frx":0026
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   3720
         Width           =   855
      End
      Begin VB.ListBox cgst 
         BackColor       =   &H00FFFFC0&
         Height          =   2205
         ItemData        =   "purchase.frx":0028
         Left            =   12480
         List            =   "purchase.frx":002A
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   3720
         Width           =   615
      End
      Begin VB.ListBox smt 
         BackColor       =   &H00FFFFC0&
         Height          =   2205
         ItemData        =   "purchase.frx":002C
         Left            =   14520
         List            =   "purchase.frx":002E
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   3720
         Width           =   855
      End
      Begin VB.ListBox net 
         BackColor       =   &H00FFFFC0&
         Height          =   2205
         ItemData        =   "purchase.frx":0030
         Left            =   15360
         List            =   "purchase.frx":0032
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   3720
         Width           =   1335
      End
      Begin VB.TextBox TXT7 
         BackColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   6360
         Width           =   1095
      End
      Begin VB.TextBox TXT8 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   435
         Left            =   8280
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   6480
         Width           =   1815
      End
      Begin VB.TextBox TXT11 
         DataField       =   "DUE_AMT"
         ForeColor       =   &H80000007&
         Height          =   435
         Left            =   12600
         MaxLength       =   8
         TabIndex        =   9
         Top             =   6960
         Width           =   1815
      End
      Begin VB.TextBox TXT9 
         BackColor       =   &H8000000B&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   435
         Left            =   12600
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   6480
         Width           =   1575
      End
      Begin VB.TextBox final 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   855
         Left            =   17880
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   4920
         Width           =   1935
      End
      Begin VB.TextBox TXT71 
         BackColor       =   &H8000000B&
         Height          =   435
         Left            =   17280
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2520
         Width           =   1815
      End
      Begin VB.TextBox TXT14 
         BackColor       =   &H8000000B&
         Height          =   435
         Left            =   2520
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   7320
         Width           =   2895
      End
      Begin VB.ComboBox TXT4 
         BackColor       =   &H8000000B&
         DataField       =   "MED_NM"
         Height          =   315
         Left            =   8160
         TabIndex        =   4
         Top             =   2640
         Width           =   2055
      End
      Begin VB.TextBox Text7 
         DataField       =   "DUE_AMT"
         ForeColor       =   &H80000007&
         Height          =   435
         Left            =   12600
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   7440
         Width           =   1815
      End
      Begin VB.OptionButton TXT132 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0ECB7&
         Caption         =   "Cheque"
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   2520
         TabIndex        =   2
         Top             =   6840
         Width           =   1335
      End
      Begin VB.OptionButton TXT131 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0ECB7&
         Caption         =   "CASH"
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   3960
         TabIndex        =   1
         Top             =   6840
         Width           =   1335
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   240
         TabIndex        =   28
         Top             =   960
         Width           =   19815
         Begin VB.TextBox TXT32 
            DataField       =   "S_MOB"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   435
            Left            =   8400
            Locked          =   -1  'True
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   480
            Width           =   2055
         End
         Begin VB.TextBox TXT31 
            DataField       =   "S_NM"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   435
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   480
            Width           =   3015
         End
         Begin VB.ComboBox TXT3 
            BackColor       =   &H8000000B&
            DataField       =   "S_ID"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   2640
            TabIndex        =   31
            Top             =   120
            Width           =   2295
         End
         Begin VB.TextBox TXT34 
            DataField       =   "S_LOC"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   435
            Left            =   14280
            Locked          =   -1  'True
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   480
            Width           =   5055
         End
         Begin VB.TextBox TXT33 
            DataField       =   "COMP_NM"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   435
            Left            =   14280
            Locked          =   -1  'True
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   0
            Width           =   2655
         End
         Begin VB.Label Label5 
            BackColor       =   &H8000000E&
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   495
            Index           =   0
            Left            =   1680
            TabIndex        =   39
            Top             =   0
            Width           =   135
         End
         Begin VB.Label Label114 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "Supplier ID  :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   0
            TabIndex        =   40
            Top             =   120
            Width           =   1590
         End
         Begin VB.Label Label51 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "Mobile No.  :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   6705
            TabIndex        =   38
            Top             =   600
            Width           =   1515
         End
         Begin VB.Label Label52 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "Supplier Name :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   0
            TabIndex        =   37
            Top             =   600
            Width           =   1920
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Supplier Address   :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   11745
            TabIndex        =   36
            Top             =   600
            Width           =   2385
         End
         Begin VB.Label Label92 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Company Name  :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   11760
            TabIndex        =   35
            Top             =   120
            Width           =   2415
         End
         Begin VB.Label Label5 
            BackColor       =   &H8000000E&
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   495
            Index           =   1
            Left            =   13920
            TabIndex        =   34
            Top             =   360
            Width           =   135
         End
      End
      Begin VB.Label Label24 
         BackColor       =   &H8000000E&
         Caption         =   "Cheque No.           :"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   300
         TabIndex        =   46
         Top             =   7440
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Incomplete"
         Height          =   255
         Left            =   16800
         TabIndex        =   98
         Top             =   5520
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000E&
         Caption         =   "Cash        :"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   360
         TabIndex        =   96
         Top             =   7440
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "purchase rate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11760
         TabIndex        =   93
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "unit"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   8640
         TabIndex        =   88
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Unit:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10200
         TabIndex        =   86
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "Product Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   84
         Top             =   2280
         Width           =   2655
      End
      Begin VB.Label Label48 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   13320
         TabIndex        =   44
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label43 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   11040
         TabIndex        =   79
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label41 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Quantity"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9840
         TabIndex        =   78
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label40 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Purchase Rate"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7440
         TabIndex        =   77
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label22 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Serial No."
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   76
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label Label102 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Price:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   14400
         TabIndex        =   75
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label101 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "GST Rate:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   15840
         TabIndex        =   74
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label100 
         BackColor       =   &H8000000E&
         Caption         =   "%"
         Height          =   375
         Left            =   17040
         TabIndex        =   73
         Top             =   2640
         Width           =   255
      End
      Begin VB.Label Label42 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13080
         TabIndex        =   72
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label46 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "Product Id:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   71
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label45 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "Product Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   70
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Order Date  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   12000
         TabIndex        =   69
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000E&
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Index           =   2
         Left            =   7200
         TabIndex        =   68
         Top             =   1440
         Width           =   135
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   1560
         TabIndex        =   67
         Top             =   360
         Width           =   165
      End
      Begin VB.Label Label53 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Order no :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   66
         Top             =   480
         Width           =   1245
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000080&
         X1              =   1800
         X2              =   20160
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "MAIN  INFORMATION"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   240
         TabIndex        =   65
         Top             =   120
         Width           =   1920
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "ADD   PRODUCT"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   255
         TabIndex        =   64
         Top             =   1920
         Width           =   1425
      End
      Begin VB.Label Label78 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Product Type"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5640
         TabIndex        =   63
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label77 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Product Name"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2040
         TabIndex        =   62
         Top             =   3120
         Width           =   3615
      End
      Begin VB.Label Label76 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Product Id"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   840
         TabIndex        =   61
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label75 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SGST"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13920
         TabIndex        =   60
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label74 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   13080
         TabIndex        =   59
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label Label70 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   14520
         TabIndex        =   58
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label Label69 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CGST"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12480
         TabIndex        =   57
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label68 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   12480
         TabIndex        =   56
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label Label66 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   13920
         TabIndex        =   55
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TOTAL   PRICE"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   15360
         TabIndex        =   54
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Line Line2 
         X1              =   1680
         X2              =   20160
         Y1              =   6120
         Y2              =   6120
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Payment By          :"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   300
         TabIndex        =   53
         Top             =   6900
         Width           =   1500
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "No. of product      :"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001468F3&
         Height          =   255
         Left            =   300
         TabIndex        =   52
         Top             =   6480
         Width           =   1530
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   7620
         TabIndex        =   51
         Top             =   6480
         Width           =   480
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "With Tax Amount :"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   10860
         TabIndex        =   50
         Top             =   6480
         Width           =   1665
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Advance Payment  :  "
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   10695
         TabIndex        =   49
         Top             =   6960
         Width           =   1830
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Final Amount"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   315
         Left            =   18000
         TabIndex        =   48
         Top             =   4560
         Width           =   1485
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Price"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   17040
         TabIndex        =   47
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000E&
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Index           =   9
         Left            =   1320
         TabIndex        =   45
         Top             =   480
         Width           =   135
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Product Weight"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8475
         TabIndex        =   43
         Top             =   2280
         Width           =   1365
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "CALCULATION"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   255
         TabIndex        =   42
         Top             =   6000
         Width           =   1275
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Remaining Amount :  "
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   10620
         TabIndex        =   41
         Top             =   7440
         Width           =   1905
      End
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11520
      TabIndex        =   92
      Top             =   2160
      Width           =   1095
   End
End
Attribute VB_Name = "order"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Combo1_Click()
Set R = New ADODB.Recordset
SQL = "select *from product where P_id='" + Combo1.Text + "'"
Set R = C.Execute(SQL)
Combo2.Text = R.Fields(1)
'Combo2.Text = R.Fields(1)
TXT41.Text = R.Fields(2)
txt4.Text = R.Fields(4)
Text1.Text = R.Fields("p_unit")
Text3.Text = R.Fields(5)
SQL = "select *from supplierprd where P_id='" + Combo1.Text + "'"
Set R = C.Execute(SQL)
Text4.Text = R.Fields(2)
End Sub

Private Sub Command1_Click()
id.AddItem Combo1.Text
nm.AddItem Combo2.Text
typ.AddItem TXT41.Text
prate.AddItem Text4.Text
UNIT.AddItem Text1.Text
qty.AddItem txt5.Text
prc.AddItem Text2.Text
cgst.AddItem Val(Text3.Text) / 2
sgst.AddItem Val(Text3.Text) / 2
cmt.AddItem (Val(Text2.Text) * Val(Text3.Text) / 100) / 2
smt.AddItem (Val(Text2.Text) * Val(Text3.Text) / 100) / 2
net.AddItem TXT71.Text
If sr.ListCount = 0 Then
sr.AddItem 1
Else
sr.AddItem (sr.ListCount + 1)
End If
Dim l As Long
Dim lSum As Long
For l = 0 To net.ListCount - 1
    lSum = lSum + CLng(net.List(l))
Next
final.Text = lSum
txt7.Text = sr.ListCount
TXT9.Text = final.Text


Combo1.Text = " "
Combo2.Text = " "
TXT41.Text = " "
Text4.Text = " "
Text1.Text = " "
txt5.Text = " "
Text2.Text = " "
txt4.Text = " "
Text3.Text = " "
TXT71.Text = " "
End Sub

Private Sub Command2_Click()
If sr.ListCount <> 0 Then Frame4.Enabled = False
txt11.Text = ""
Text7.Text = ""
a = InputBox("Enter the Serial No. you want to remove:", "for delete")
If a = blank Then
MsgBox "Please enter serial no."
Else
TXT9.Text = Val(TXT9.Text) - net.List(a - 1)
final.Text = TXT9.Text
id.RemoveItem (a - 1)
nm.RemoveItem (a - 1)
UNIT.RemoveItem (a - 1)
typ.RemoveItem (a - 1)
qty.RemoveItem (a - 1)
prate.RemoveItem (a - 1)
prc.RemoveItem (a - 1)
If cgst.List(a - 1) <> "" Then cgst.RemoveItem (a - 1)
If cmt.List(a - 1) <> "" Then cmt.RemoveItem (a - 1)
If sgst.List(a - 1) <> "" Then sgst.RemoveItem (a - 1)
If smt.List(a - 1) <> "" Then smt.RemoveItem (a - 1)
net.RemoveItem (a - 1)
sr.Clear
txt7.Text = id.ListCount
For i = 1 To id.ListCount
    sr.AddItem i
Next i
If sr.ListCount <> 0 Then Frame4.Enabled = False
If sr.ListCount = 0 Then Command2.Enabled = False

Dim l As Long
Dim lSum As Long
For l = 0 To prc.ListCount - 1
    lSum = lSum + CLng(prc.List(l))
Next
TXT8.Text = lSum
End If
End Sub

Private Sub Command3_Click()
If txt1.Text = blank Or txt3.Text = blank Or txt2.Text = blank Or txt7.Text = blank Or txt14.Text = blank Or txt11.Text = blank Then
MsgBox "Please fill all the details first!!"
Else
answer = MsgBox("Do you want to place order ?", vbExclamation + vbYesNo, "add confirm")
If answer = vbYes Then
Set R = New ADODB.Recordset
SQL = "insert into purordetail values('" + txt1.Text + "','" + txt2.Text + "','" + txt3.Text + "'," + txt7.Text + ",'" + Text6.Text + "','" + txt14.Text + "'," + TXT8.Text + "," + TXT9.Text + "," + txt11.Text + "," + Text7.Text + ",'" + Label9.Caption + "')"
Set R = C.Execute(SQL)
MsgBox "order placed!!"
Else
MsgBox "Data not saved"
End If

Dim i As Long
For i = 0 To sr.ListCount - 1
SQL = "insert into p_det values('" + txt1.Text + "'," + sr.List(i) + ",'" + id.List(i) + "','" + nm.List(i) + "','" + typ.List(i) + "'," + prate.List(i) + ",'" + UNIT.List(i) + "'," + qty.List(i) + "," + prc.List(i) + "," + cgst.List(i) + "," + cmt.List(i) + "," + sgst.List(i) + "," + smt.List(i) + "," + net.List(i) + ")"
Set R = C.Execute(SQL)
Next i

MsgBox "Data saved"
Unload Me
order.Show
order.Top = 0
order.Left = 0
End If
End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub final_Change()
Dim l As Long
Dim lSum As Long
For l = 0 To prc.ListCount - 1
    lSum = lSum + CLng(prc.List(l))
Next
TXT8.Text = lSum
End Sub

Private Sub Form_Load()
CONN
Dim a As String

Command1.Enabled = True
Command3.Enabled = False
Set R = New ADODB.Recordset
SQL = "select max(to_number(SUBSTR( PUR_ORDERNO,5,LENGTH(PUR_ORDERNO))))from purordetail"
Set R = C.Execute(SQL)
If IsNull(R.Fields(0)) Then
txt1.Text = "OD" & "00" & 1
Else
txt1.Text = "OD" & "00" & R.Fields(0) + 1
a = txt1.Text
End If
If (a = "OD0010") Then
Set R = New ADODB.Recordset
SQL = "select max(to_number(SUBSTR( PUR_ORDERNO,4,LENGTH(PUR_ORDERNO))))from purordetail"
Set R = C.Execute(SQL)
txt1.Text = "OD" & "0" & R.Fields(0) + 1
End If
Text6.Visible = False
MonthView1.Visible = False
Set R = New ADODB.Recordset
SQL = "select *from supplier"
Set R = C.Execute(SQL)
While R.EOF = False
txt3.AddItem R.Fields(0)
R.MoveNext
Wend
'auto_combo
MonthView1.Refresh
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
txt2.Text = Format(MonthView1, "dd-mmm-yyyy")
MonthView1.Visible = False
End Sub


Private Sub Text3_Change()
'TXT71.Text = Val(Text2.Text) + (Val(Text2.Text) * Val(Text3.Text) / 100)
End Sub



Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txt5.SetFocus
End Sub


Private Sub Text7_Change()
Command3.Enabled = True
End Sub

Private Sub Txt5_KeyPress(KeyAscii As Integer)
0 If KeyAscii = 13 Then Text3.SetFocus
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command1.SetFocus
End Sub
Private Sub TXT11_LostFocus()
Text7.Text = Val(TXT9.Text) - Val(txt11.Text)
End Sub
Private Sub Txt11_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text7.SetFocus
End Sub
Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command3.SetFocus
End Sub
Private Sub TXT131_Click()
Text6.Text = TXT131.Caption
txt14.Text = "Cash"
Label24.Visible = False
Label7.Visible = True
End Sub

Private Sub TXT132_Click()
Text6.Text = TXT132.Caption
Label7.Visible = False
Label24.Visible = True
End Sub

Private Sub TXT2_Click()
MonthView1.Visible = True
End Sub

Private Sub TXT3_Click()
Set R = New ADODB.Recordset
SQL = "select *from supplier where SUP_ID='" + txt3.Text + "'"
Set R = C.Execute(SQL)
TXT31.Text = R.Fields(1)
TXT32.Text = R.Fields(2)
TXT33.Text = R.Fields(7)
TXT34.Text = R.Fields(3)
Combo1.Clear
Set R = New ADODB.Recordset
SQL = "select *from supplierprd where sup_id='" + txt3.Text + "'"
Set R = C.Execute(SQL)
While R.EOF = False
Combo1.AddItem R.Fields(1)
R.MoveNext
Wend
End Sub


Public Function auto_combo()
Set R = New ADODB.Recordset
SQL = "select *from product"
Set R = C.Execute(SQL)
While R.EOF = False
Combo1.AddItem R.Fields(0)
R.MoveNext
Wend
End Function

Private Sub Txt5_Change()
Text2.Text = Val(Text4.Text) * Val(txt5.Text)
TXT71.Text = Val(Text2.Text) + (Val(Text2.Text) * Val(Text3.Text) / 100)
End Sub

Private Sub TXT71_Change()
'Command1.Enabled = True
End Sub




