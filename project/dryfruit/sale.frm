VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Sell 
   Caption         =   "Sell"
   ClientHeight    =   8850
   ClientLeft      =   270
   ClientTop       =   5220
   ClientWidth     =   20250
   FillStyle       =   2  'Horizontal Line
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "sale.frx":0000
   ScaleHeight     =   8850
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.PictureBox SSTab1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   14805
      Index           =   0
      Left            =   0
      ScaleHeight     =   14745
      ScaleWidth      =   21825
      TabIndex        =   0
      Top             =   0
      Width           =   21885
      Begin VB.Frame Frame9 
         BorderStyle     =   0  'None
         Caption         =   " "
         Height          =   8175
         Index           =   0
         Left            =   -69000
         TabIndex        =   73
         Top             =   360
         Width           =   13215
         Begin VB.PictureBox Adodc3 
            BackColor       =   &H000000FF&
            Height          =   1000
            Index           =   0
            Left            =   0
            ScaleHeight     =   945
            ScaleWidth      =   945
            TabIndex        =   77
            Top             =   0
            Width           =   1000
         End
         Begin VB.TextBox Text10 
            DataField       =   "SORD_NO"
            Height          =   495
            Index           =   0
            Left            =   8280
            TabIndex        =   76
            Top             =   3960
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.PictureBox DATA 
            BackColor       =   &H000000FF&
            Height          =   1000
            Index           =   0
            Left            =   0
            ScaleHeight     =   945
            ScaleWidth      =   945
            TabIndex        =   75
            Top             =   0
            Width           =   1000
         End
         Begin VB.PictureBox MED 
            BackColor       =   &H000000FF&
            Height          =   1000
            Index           =   0
            Left            =   0
            ScaleHeight     =   945
            ScaleWidth      =   945
            TabIndex        =   74
            Top             =   0
            Width           =   1000
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         Height          =   14055
         Index           =   0
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   20175
         Begin VB.TextBox camnt 
            Height          =   285
            Left            =   10080
            TabIndex        =   118
            Text            =   "Text5"
            Top             =   8280
            Width           =   855
         End
         Begin MSComCtl2.MonthView MonthView1 
            Height          =   2370
            Left            =   15600
            TabIndex        =   116
            Top             =   -240
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   4180
            _Version        =   393216
            ForeColor       =   -2147483630
            BackColor       =   -2147483633
            Appearance      =   1
            StartOfWeek     =   133103618
            CurrentDate     =   44440
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   5520
            TabIndex        =   115
            Top             =   7200
            Width           =   975
         End
         Begin VB.TextBox Text3 
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   15480
            TabIndex        =   114
            Top             =   7680
            Width           =   975
         End
         Begin VB.TextBox Txt20 
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   12600
            TabIndex        =   112
            Top             =   7080
            Width           =   1575
         End
         Begin VB.TextBox txt18 
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8280
            TabIndex        =   111
            Top             =   7440
            Width           =   1815
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2520
            TabIndex        =   109
            Top             =   360
            Width           =   2175
         End
         Begin VB.ListBox MRP 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2100
            ItemData        =   "sale.frx":0342
            Left            =   7200
            List            =   "sale.frx":0344
            TabIndex        =   107
            Top             =   3960
            Width           =   975
         End
         Begin VB.ListBox id 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2100
            ItemData        =   "sale.frx":0346
            Left            =   840
            List            =   "sale.frx":0348
            TabIndex        =   106
            Top             =   3960
            Width           =   1215
         End
         Begin VB.ComboBox Combo6 
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   2520
            TabIndex        =   101
            Top             =   960
            Width           =   2175
         End
         Begin VB.TextBox TX2 
            BackColor       =   &H00FFFFFF&
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   7560
            TabIndex        =   99
            TabStop         =   0   'False
            Top             =   240
            Width           =   2175
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00FFFF00&
            Caption         =   "Add New Customer"
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
            Left            =   18240
            Style           =   1  'Graphical
            TabIndex        =   97
            Top             =   960
            Width           =   1695
         End
         Begin VB.CommandButton Command7 
            BackColor       =   &H00FFFF00&
            Caption         =   "Cancel"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   17280
            Style           =   1  'Graphical
            TabIndex        =   96
            Top             =   7440
            Width           =   2775
         End
         Begin VB.CommandButton Command6 
            BackColor       =   &H00FFFF00&
            Caption         =   "Generate Invoice"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   17280
            Style           =   1  'Graphical
            TabIndex        =   95
            Top             =   6840
            Width           =   2775
         End
         Begin VB.CommandButton Command5 
            BackColor       =   &H00FFFF00&
            Caption         =   "Delete"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   18480
            Style           =   1  'Graphical
            TabIndex        =   94
            Top             =   4080
            Width           =   1215
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00FFFF00&
            Caption         =   "Add"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   18480
            Style           =   1  'Graphical
            TabIndex        =   93
            Top             =   3480
            Width           =   1215
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H8000000E&
            BorderStyle     =   0  'None
            Height          =   1095
            Index           =   0
            Left            =   120
            TabIndex        =   35
            Top             =   1200
            Width           =   20175
            Begin RichTextLib.RichTextBox tx4 
               Height          =   855
               Left            =   7560
               TabIndex        =   117
               Top             =   120
               Width           =   3975
               _ExtentX        =   7011
               _ExtentY        =   1508
               _Version        =   393217
               TextRTF         =   $"sale.frx":034A
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Palatino Linotype"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.TextBox Text2 
               Height          =   375
               Left            =   14400
               TabIndex        =   110
               Top             =   0
               Width           =   975
            End
            Begin VB.ComboBox Combo4 
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Left            =   2400
               TabIndex        =   104
               Top             =   480
               Width           =   2175
            End
            Begin VB.OptionButton Option3 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Transgender"
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
               Left            =   15240
               TabIndex        =   102
               Top             =   480
               Width           =   1695
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Male"
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
               Left            =   13320
               TabIndex        =   37
               Top             =   480
               Width           =   855
            End
            Begin VB.OptionButton Option2 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Female"
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
               Left            =   14160
               TabIndex        =   36
               Top             =   480
               Width           =   1095
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Customer Name    :"
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   2
               Left            =   120
               TabIndex        =   42
               Top             =   480
               Width           =   2025
            End
            Begin VB.Label Label6 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H8000000E&
               Caption         =   "Gender :"
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   12240
               TabIndex        =   41
               Top             =   600
               Width           =   945
            End
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H8000000E&
               Caption         =   "Customer Address    :"
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   5160
               TabIndex        =   40
               Top             =   360
               Width           =   2265
            End
            Begin VB.Label Label30 
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
               Left            =   1800
               TabIndex        =   39
               Top             =   240
               Width           =   135
            End
            Begin VB.Label Label10 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H8000000E&
               Caption         =   "Add Product:"
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004000&
               Height          =   330
               Index           =   1
               Left            =   0
               TabIndex        =   38
               Top             =   840
               Width           =   1515
            End
            Begin VB.Line Line1 
               Index           =   0
               X1              =   1560
               X2              =   19920
               Y1              =   1080
               Y2              =   1080
            End
         End
         Begin VB.ListBox UNIT 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2100
            ItemData        =   "sale.frx":03D0
            Left            =   5640
            List            =   "sale.frx":03D2
            TabIndex        =   34
            Top             =   3960
            Width           =   1575
         End
         Begin VB.ComboBox Combo2 
            BackColor       =   &H00FFFFFF&
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3165
            TabIndex        =   33
            Top             =   2760
            Width           =   2295
         End
         Begin VB.ComboBox Combo3 
            BackColor       =   &H00FFFFFF&
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   6075
            TabIndex        =   32
            Top             =   2760
            Width           =   1695
         End
         Begin VB.TextBox Tx9 
            BackColor       =   &H00FFFFFF&
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   9960
            Locked          =   -1  'True
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   2760
            Width           =   975
         End
         Begin VB.TextBox Tx10 
            BackColor       =   &H00E0E0E0&
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000009&
            Height          =   435
            Left            =   11355
            Locked          =   -1  'True
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   2760
            Width           =   855
         End
         Begin VB.TextBox TXT16 
            BackColor       =   &H8000000B&
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   7560
            Width           =   2895
         End
         Begin VB.TextBox TX14 
            BackColor       =   &H8000000B&
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   16920
            Locked          =   -1  'True
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   2760
            Width           =   1455
         End
         Begin VB.TextBox final 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            CausesValidation=   0   'False
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
            Left            =   18360
            Locked          =   -1  'True
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   5400
            Width           =   1575
         End
         Begin VB.TextBox TXT21 
            BackColor       =   &H8000000B&
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   12600
            TabIndex        =   26
            Top             =   7560
            Width           =   1575
         End
         Begin VB.TextBox TXT19 
            BackColor       =   &H8000000B&
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   12600
            Locked          =   -1  'True
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   6600
            Width           =   1575
         End
         Begin VB.TextBox TXT17 
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   8280
            Locked          =   -1  'True
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   6720
            Width           =   1815
         End
         Begin VB.TextBox TXT15 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   6600
            Width           =   1095
         End
         Begin VB.OptionButton TXT131 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Caption         =   "Cash"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   435
            Left            =   2520
            TabIndex        =   22
            Top             =   7080
            Width           =   1335
         End
         Begin VB.OptionButton TXT132 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Caption         =   "Cheque"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   435
            Left            =   3960
            TabIndex        =   21
            Top             =   7080
            Width           =   1335
         End
         Begin VB.ListBox net 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2100
            ItemData        =   "sale.frx":03D4
            Left            =   14520
            List            =   "sale.frx":03D6
            TabIndex        =   20
            Top             =   3960
            Width           =   1455
         End
         Begin VB.ListBox smt 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2100
            ItemData        =   "sale.frx":03D8
            Left            =   13560
            List            =   "sale.frx":03DA
            TabIndex        =   19
            Top             =   3960
            Width           =   975
         End
         Begin VB.ListBox cgst 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2100
            ItemData        =   "sale.frx":03DC
            Left            =   11040
            List            =   "sale.frx":03DE
            TabIndex        =   18
            Top             =   3960
            Width           =   735
         End
         Begin VB.ListBox cmt 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2100
            ItemData        =   "sale.frx":03E0
            Left            =   11760
            List            =   "sale.frx":03E2
            TabIndex        =   17
            Top             =   3960
            Width           =   1095
         End
         Begin VB.ListBox sgst 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2100
            ItemData        =   "sale.frx":03E4
            Left            =   12840
            List            =   "sale.frx":03E6
            TabIndex        =   16
            Top             =   3960
            Width           =   735
         End
         Begin VB.ListBox PN 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2100
            ItemData        =   "sale.frx":03E8
            Left            =   2040
            List            =   "sale.frx":03EA
            TabIndex        =   15
            Top             =   3960
            Width           =   3615
         End
         Begin VB.ListBox qty 
            BackColor       =   &H00FFFFC0&
            Columns         =   3
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2100
            ItemData        =   "sale.frx":03EC
            Left            =   8160
            List            =   "sale.frx":03EE
            TabIndex        =   14
            Top             =   3960
            Width           =   1215
         End
         Begin VB.ListBox prc 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2100
            ItemData        =   "sale.frx":03F0
            Left            =   9360
            List            =   "sale.frx":03F2
            TabIndex        =   13
            Top             =   3960
            Width           =   1695
         End
         Begin VB.ListBox sr 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2100
            IntegralHeight  =   0   'False
            ItemData        =   "sale.frx":03F4
            Left            =   240
            List            =   "sale.frx":03F6
            TabIndex        =   12
            Top             =   3960
            Width           =   615
         End
         Begin VB.TextBox TX13 
            BackColor       =   &H8000000B&
            CausesValidation=   0   'False
            DataField       =   "GST"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   15600
            Locked          =   -1  'True
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   2760
            Width           =   615
         End
         Begin VB.TextBox TXT8 
            BackColor       =   &H00FFFFFF&
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   8340
            Locked          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   2760
            Width           =   1335
         End
         Begin VB.TextBox TX11 
            BackColor       =   &H00DEF4D9&
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   12600
            MaxLength       =   3
            TabIndex        =   9
            Top             =   2760
            Width           =   960
         End
         Begin VB.TextBox TX12 
            BackColor       =   &H8000000B&
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   13845
            Locked          =   -1  'True
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   2760
            Width           =   1335
         End
         Begin VB.TextBox TXT6 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000E&
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   435
            Left            =   14040
            Locked          =   -1  'True
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   120
            Width           =   2175
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H00DEF4D9&
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            TabIndex        =   6
            Top             =   2760
            Width           =   2175
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "New Dues:"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   14280
            TabIndex        =   113
            Top             =   7680
            Width           =   1335
         End
         Begin VB.Label Label28 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "MRP"
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
            Index           =   1
            Left            =   7200
            TabIndex        =   108
            Top             =   3360
            Width           =   975
         End
         Begin VB.Label Label450 
            BackColor       =   &H8000000E&
            Caption         =   "Cash Payment:"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   600
            TabIndex        =   105
            Top             =   7680
            Width           =   2295
         End
         Begin VB.Label Label23 
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
            Height          =   255
            Index           =   2
            Left            =   1560
            TabIndex        =   103
            Top             =   720
            Width           =   135
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Customer Id          :"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   255
            TabIndex        =   100
            Top             =   960
            Width           =   1995
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Mobile No.            :"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   5280
            TabIndex        =   98
            Top             =   360
            Width           =   1950
         End
         Begin VB.Label Label23 
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
            Height          =   375
            Index           =   1
            Left            =   6840
            TabIndex        =   92
            Top             =   120
            Width           =   135
         End
         Begin VB.Label Label23 
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
            Height          =   255
            Index           =   0
            Left            =   1560
            TabIndex        =   91
            Top             =   240
            Width           =   135
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   " Order No               :"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   240
            TabIndex        =   90
            Top             =   360
            Width           =   2040
         End
         Begin VB.Label Label40 
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
            Index           =   1
            Left            =   14520
            TabIndex        =   89
            Top             =   3360
            Width           =   1455
         End
         Begin VB.Label Label36 
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
            Index           =   2
            Left            =   13560
            TabIndex        =   88
            Top             =   3720
            Width           =   975
         End
         Begin VB.Label Label35 
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
            Index           =   2
            Left            =   12840
            TabIndex        =   87
            Top             =   3720
            Width           =   735
         End
         Begin VB.Label Label34 
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
            Left            =   12840
            TabIndex        =   86
            Top             =   3360
            Width           =   1695
         End
         Begin VB.Label Label33 
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
            Index           =   7
            Left            =   11760
            TabIndex        =   85
            Top             =   3720
            Width           =   1095
         End
         Begin VB.Label Label32 
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
            Index           =   6
            Left            =   11040
            TabIndex        =   84
            Top             =   3720
            Width           =   735
         End
         Begin VB.Label Label30 
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
            Index           =   1
            Left            =   9360
            TabIndex        =   83
            Top             =   3360
            Width           =   1695
         End
         Begin VB.Label Label29 
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
            Index           =   1
            Left            =   8160
            TabIndex        =   82
            Top             =   3360
            Width           =   1215
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "  Qty   Available"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   11280
            TabIndex        =   72
            Top             =   2280
            Width           =   1065
         End
         Begin VB.Label Label451 
            BackColor       =   &H8000000E&
            Caption         =   "Cheque No.    :"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   600
            TabIndex        =   71
            Top             =   7680
            Width           =   2295
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Total Price"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   16800
            TabIndex        =   70
            Top             =   2400
            Width           =   1455
         End
         Begin VB.Label Label41 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Final Amount:"
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
            Height          =   795
            Index           =   0
            Left            =   18480
            TabIndex        =   69
            Top             =   4680
            Width           =   1305
         End
         Begin VB.Label Label48 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Net Amount:  "
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   10920
            TabIndex        =   68
            Top             =   7200
            Width           =   1485
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "With Tax Amount:"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   1
            Left            =   10560
            TabIndex        =   67
            Top             =   6720
            Width           =   1965
         End
         Begin VB.Label Label49 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Amount Paid:  "
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   10800
            TabIndex        =   66
            Top             =   7680
            Width           =   1575
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Prev Dues:"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   1
            Left            =   6960
            TabIndex        =   65
            Top             =   7320
            Width           =   1140
         End
         Begin VB.Label Label113 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Total:"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   0
            Left            =   7440
            TabIndex        =   64
            Top             =   6840
            Width           =   600
         End
         Begin VB.Label Label43 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "No. of product:"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   600
            TabIndex        =   63
            Top             =   6720
            Width           =   1590
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "Payment By    :"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   0
            Left            =   600
            TabIndex        =   62
            Top             =   7200
            Width           =   1560
         End
         Begin VB.Label Label42 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "Calculation:"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   330
            Index           =   8
            Left            =   240
            TabIndex        =   61
            Top             =   6240
            Width           =   1365
         End
         Begin VB.Line Line2 
            Index           =   0
            X1              =   1560
            X2              =   20040
            Y1              =   6480
            Y2              =   6480
         End
         Begin VB.Label Label31 
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
            Left            =   11040
            TabIndex        =   60
            Top             =   3360
            Width           =   1815
         End
         Begin VB.Label Label24 
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
            Index           =   1
            Left            =   840
            TabIndex        =   59
            Top             =   3360
            Width           =   1215
         End
         Begin VB.Label Label25 
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
            Index           =   1
            Left            =   2040
            TabIndex        =   58
            Top             =   3360
            Width           =   3615
         End
         Begin VB.Label Label23 
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
            Index           =   5
            Left            =   240
            TabIndex        =   57
            Top             =   3360
            Width           =   615
         End
         Begin VB.Label percent 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            Height          =   375
            Left            =   16320
            TabIndex        =   56
            Top             =   2880
            Width           =   255
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "GST Rate:"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   1
            Left            =   15435
            TabIndex        =   55
            Top             =   2400
            Width           =   945
         End
         Begin VB.Label Lable11 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "Product Id:"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   390
            TabIndex        =   54
            Top             =   2400
            Width           =   1185
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "Unit:"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   6600
            TabIndex        =   53
            Top             =   2400
            Width           =   555
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "MRP  :"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   8400
            TabIndex        =   52
            Top             =   2400
            Width           =   555
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Price:"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   14085
            TabIndex        =   51
            Top             =   2400
            Width           =   525
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "Main Information:"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   330
            Index           =   0
            Left            =   210
            TabIndex        =   50
            Top             =   0
            Width           =   2085
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Order Date    :"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   12360
            TabIndex        =   49
            Top             =   240
            Width           =   1515
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "Product Name:"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   3450
            TabIndex        =   48
            Top             =   2400
            Width           =   1575
         End
         Begin VB.Label Labe17 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Rack No."
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   10020
            TabIndex        =   47
            Top             =   2400
            Width           =   855
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Quantity:"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   1
            Left            =   12600
            TabIndex        =   46
            Top             =   2400
            Width           =   885
         End
         Begin VB.Label Label27 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Unit"
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
            TabIndex        =   45
            Top             =   3360
            Width           =   1575
         End
         Begin VB.Label Label29 
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
            Index           =   0
            Left            =   13800
            TabIndex        =   44
            Top             =   120
            Width           =   135
         End
         Begin VB.Label Label51 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   14280
            TabIndex        =   43
            Top             =   7680
            Width           =   45
         End
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Show All"
         Height          =   495
         Index           =   0
         Left            =   -72360
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   4440
         Width           =   1695
      End
      Begin VB.TextBox Text9 
         DataField       =   "SORD_NO"
         Height          =   495
         Index           =   0
         Left            =   -73320
         TabIndex        =   3
         Top             =   2640
         Width           =   2535
      End
      Begin VB.CommandButton Command4 
         Height          =   495
         Index           =   0
         Left            =   -74280
         Picture         =   "sale.frx":03F8
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   4440
         Width           =   1695
      End
      Begin VB.CommandButton CLOSE 
         BackColor       =   &H00FBE2BD&
         Caption         =   "CLOSE"
         Height          =   315
         Index           =   1
         Left            =   -61440
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   6615
      End
      Begin VB.Image Image3 
         Height          =   8025
         Index           =   0
         Left            =   -75000
         Picture         =   "sale.frx":0B0E
         Stretch         =   -1  'True
         Top             =   360
         Width           =   5655
      End
      Begin VB.Label Label45 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "FROM:"
         ForeColor       =   &H00400000&
         Height          =   315
         Index           =   0
         Left            =   -74280
         TabIndex        =   81
         Top             =   3360
         Width           =   780
      End
      Begin VB.Label Label46 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "SEARCH  : "
         ForeColor       =   &H000000FF&
         Height          =   315
         Index           =   0
         Left            =   -74640
         TabIndex        =   80
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label47 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "TO:"
         ForeColor       =   &H00400000&
         Height          =   315
         Index           =   0
         Left            =   -73920
         TabIndex        =   79
         Top             =   3840
         Width           =   405
      End
      Begin VB.Label Label107 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "(mm/dd/yyyy)"
         BeginProperty Font 
            Name            =   "Constantia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Index           =   0
         Left            =   -70800
         TabIndex        =   78
         Top             =   3360
         Width           =   1320
      End
   End
End
Attribute VB_Name = "Sell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
Set R = New ADODB.Recordset
SQL = "select *from product where p_id='" + Combo1.Text + "'"
Set R = C.Execute(SQL)
Combo2.Text = R.Fields(1)
Combo3.Text = R.Fields(7)
TXT8.Text = R.Fields(6)
TX13.Text = R.Fields(5)
Set R = New ADODB.Recordset
SQL = "select *from stock where p_id='" + Combo1.Text + "'"
Set R = C.Execute(SQL)
tx9.Text = R.Fields(0)
tx10.Text = R.Fields(2)
End Sub

Private Sub Combo4_Click()
Set R = New ADODB.Recordset
SQL = "select *from customer where c_nm='" + Combo4.Text + "'"
Set R = C.Execute(SQL)
Combo6.Text = R.Fields(0)
tx2.Text = R.Fields(2)
tx4.Text = R.Fields(3)
Text2.Text = R.Fields(4)
txt18.Text = R.Fields(6)
End Sub

Private Sub Combo6_Click()
Combo6.Refresh
Set R = New ADODB.Recordset
SQL = "select *from customer where c_id='" + Combo6.Text + "'"
Set R = C.Execute(SQL)
Combo4.Text = R.Fields(1)
tx2.Text = R.Fields(2)
tx4.Text = R.Fields(3)
Text2.Text = R.Fields(4)
txt18.Text = R.Fields(6)
End Sub

Private Sub Command1_Click()
customer.Show
End Sub

Public Function auto_c_id()
Set R = New ADODB.Recordset
SQL = "select *from customer order by c_id"
Set R = C.Execute(SQL)
While R.EOF = False
Combo6.AddItem R.Fields(0)
R.MoveNext
Wend
End Function

Public Function auto_c_nm()
Set R = New ADODB.Recordset
SQL = "select *from customer order by c_nm"
Set R = C.Execute(SQL)
While R.EOF = False
Combo4.AddItem R.Fields(1)
R.MoveNext
Wend
End Function
Public Function auto_p_id()
Set R = New ADODB.Recordset
SQL = "select *from stock order by p_id"
Set R = C.Execute(SQL)
While R.EOF = False
Combo1.AddItem R.Fields(1)
R.MoveNext
Wend
End Function

Private Sub Command2_Click()
id.AddItem Combo1.Text
PN.AddItem Combo2.Text
UNIT.AddItem Combo3.Text
MRP.AddItem TXT8.Text
qty.AddItem TX11.Text
prc.AddItem TX12.Text
net.AddItem TX14.Text
cgst.AddItem Val(TX13.Text) / 2
sgst.AddItem Val(TX13.Text) / 2
cmt.AddItem ((Val(TX13.Text) / 2) * Val(TX12.Text)) / 100
smt.AddItem ((Val(TX13.Text) / 2) * Val(TX12.Text)) / 100
Combo1.Text = " "
Combo2.Text = " "
Combo3.Text = " "
TXT8.Text = " "
TX11.Text = " "
TX12.Text = " "
TX14.Text = " "
TX13.Text = " "
tx9.Text = " "
tx10.Text = " "
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

Dim f As Long
Dim fSum As Long
For f = 0 To cmt.ListCount - 1
    fSum = fSum + CLng(cmt.List(f))
Next
camnt.Text = fSum

TXT15.Text = sr.ListCount
TXT19.Text = final.Text
'TXT20.Text = Val(TXT19.Text)
End Sub
Private Sub Command5_Click()
'If sr.ListCount = 0 Then
'MsgBox "please add product first"
a = InputBox("Enter the Serial No. you want to remove:", "for delete")
If a = blank Then
MsgBox "Please enter serial no."
Else
TXT19.Text = Val(TXT19.Text) - net.List(a - 1)
final.Text = Val(TXT19.Text)
Txt20.Text = Val(final.Text)
id.RemoveItem (a - 1)
PN.RemoveItem (a - 1)
UNIT.RemoveItem (a - 1)
MRP.RemoveItem (a - 1)
qty.RemoveItem (a - 1)
prc.RemoveItem (a - 1)
net.RemoveItem (a - 1)
cgst.RemoveItem (a - 1)
sgst.RemoveItem (a - 1)
cmt.RemoveItem (a - 1)
smt.RemoveItem (a - 1)
sr.Clear
For i = 1 To id.ListCount
    sr.AddItem i
Next i
Dim l As Long
Dim lSum As Long
For l = 0 To prc.ListCount - 1
    lSum = lSum + CLng(prc.List(l))
Next
TXT17.Text = lSum
TXT15.Text = sr.ListCount
Txt20.Text = Val(txt18.Text) + Val(Txt20.Text)
End If

End Sub

Private Sub Command6_Click()
If Text1.Text = blank Or Combo6.Text = blank Or TXT6.Text = blank Or TXT15.Text = blank Or Text4.Text = blank Or TXT16.Text = blank Or TXT17.Text = blank Or TXT21.Text = blank Then
MsgBox "Please fill the details first !!"
Else
Set R = New ADODB.Recordset
SQL = "insert into sell_details values('" + Text1.Text + "','" + TXT6.Text + "','" + Combo6.Text + "'," + TXT15.Text + ",'" + Text4.Text + "','" + TXT16.Text + "'," + TXT17.Text + "," + TXT19.Text + "," + Text3.Text + "," + TXT21.Text + ")"

Set R = C.Execute(SQL)

Dim i As Long
For i = 0 To sr.ListCount - 1
SQL = "insert into sold_pdet values('" + Text1.Text + "','" + id.List(i) + "'," + qty.List(i) + "," + net.List(i) + ")"
Set R = C.Execute(SQL)
Next

Dim k As Long
For k = 0 To sr.ListCount - 1
SQL = "update stock set avl_qty= avl_qty-" + qty.List(k) + " where p_id='" + id.List(k) + "'"
Set R = C.Execute(SQL)
Next k


SQL = "update customer set dues=" + Text3.Text + " where c_id='" + Combo6.Text + "'"
Set R = C.Execute(SQL)
MsgBox "Sell Completed!!"
'Unload Me
'Sell.Show
sale_invoice.Show
sale_invoice.Top = 0
sale_invoice.Left = 0
sale_invoice.Text2.Text = Sell.Text1.Text
sale_invoice.Text3.Text = Sell.Combo6.Text
sale_invoice.txt3.Text = Sell.Combo4.Text
sale_invoice.txt4.Text = Sell.Text2.Text
sale_invoice.Text6.Text = Sell.tx4.Text
sale_invoice.Text8.Text = Sell.tx2.Text
Dim m As Long
For m = 0 To sr.ListCount - 1
sale_invoice.List1.AddItem Sell.PN.List(m)
sale_invoice.List2.AddItem Sell.MRP.List(m)
sale_invoice.List3.AddItem Sell.qty.List(m)
sale_invoice.List4.AddItem Sell.prc.List(m)
sale_invoice.List5.AddItem Sell.cmt.List(m)
sale_invoice.List6.AddItem Sell.smt.List(m)
sale_invoice.List7.AddItem Sell.net.List(m)
Next
sale_invoice.Txt88.Text = Sell.TXT19.Text
sale_invoice.Text13.Text = Sell.Txt20.Text
sale_invoice.Text15.Text = Sell.TXT21.Text

sale_invoice.paidby.Text = Sell.Text4.Text
sale_invoice.Text10.Text = Sell.TXT17.Text
sale_invoice.Text4.Text = Sell.txt18.Text
sale_invoice.Text5.Text = Sell.Text3.Text
sale_invoice.cgmt.Text = Sell.camnt.Text
sale_invoice.sgmt.Text = Sell.camnt.Text
End If
End Sub

Private Sub Command7_Click()
Unload Me
End Sub




Private Sub final_Change()
Dim l As Long
Dim lSum As Long
For l = 0 To prc.ListCount - 1
    lSum = lSum + CLng(prc.List(l))
Next
TXT17.Text = lSum
End Sub

Private Sub Form_Load()
CONN
auto_c_id
auto_c_nm
auto_p_id

Dim a As String
Set R = New ADODB.Recordset
SQL = "select max(to_number(SUBSTR(s_ono,5,LENGTH(s_ono))))from sell_details"
Set R = C.Execute(SQL)
If IsNull(R.Fields(0)) Then
Text1.Text = "SO" & "00" & 1
Else
Text1.Text = "SO" & "00" & R.Fields(0) + 1
a = Text1.Text
End If
If (a = "SO0010") Then
Set R = New ADODB.Recordset
SQL = "select max(to_number(SUBSTR(s_ono,4,LENGTH(s_ono))))from sell_details"
Set R = C.Execute(SQL)
Text1.Text = "SO" & "0" & R.Fields(0) + 1
End If
Text2.Visible = False
MonthView1.Visible = False
camnt.Visible = False
Text4.Visible = False
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
Dim d1, d2 As String
d1 = MonthView1.Value
d2 = Date
MonthView1.Visible = False
If d1 > d2 Then
MsgBox "Don't enter future date"
TXT6.Text = ""
Else
TXT6.Text = Format(MonthView1, "dd-mmm-yyyy")
End If
End Sub
Private Sub Text2_Change()
If Text2.Text = "Male" Then
Option1.Value = True
Option2.Value = False
Option3.Value = False
ElseIf Text2.Text = "Female" Then
Option2.Value = True
Option1.Value = False
Option3.Value = False
ElseIf Text2.Text = "Transgender" Then
Option3.Value = True
Option1.Value = False
Option2.Value = False
End If
End Sub

Private Sub TX11_Change()
If Val(TX11.Text) > Val(tx10.Text) Then
MsgBox "not enough quantity"
TX11.Text = ""

End If
   TX12.Text = Val(TXT8.Text) * Val(TX11.Text)
   TX14.Text = Val(TX12.Text) + (Val(TX12.Text) * (Val(TX13.Text) / 100))
End Sub

Private Sub TXT131_Click()
Label451.Visible = False
Label450.Visible = True
TXT16.Text = "CASH"
Text4.Text = "Cash"
End Sub

Private Sub TXT132_Click()
Label450.Visible = False
Label451.Visible = True
TXT16.Text = ""
Text4.Text = "Cheque"
End Sub

Private Sub TXT19_Change()
Txt20.Text = Val(txt18.Text) + Val(TXT19.Text)
End Sub

Private Sub TXT21_Change()
Text3.Text = Val(Txt20.Text) - Val(TXT21.Text)

End Sub

Private Sub TXT6_Click()
MonthView1.Visible = True
End Sub
