VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Report 
   Caption         =   "Report"
   ClientHeight    =   7965
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11235
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11055
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   10695
      Left            =   -120
      TabIndex        =   0
      Top             =   0
      Width           =   20175
      Begin VB.Frame Frame20 
         BackColor       =   &H00FFFFFF&
         Height          =   7695
         Left            =   240
         TabIndex        =   1
         Top             =   0
         Width           =   11055
         Begin VB.Frame Frame101 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   6615
            Left            =   3000
            TabIndex        =   61
            Top             =   1080
            Width           =   8055
            Begin VB.ComboBox Combo10 
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Left            =   2880
               TabIndex        =   87
               Text            =   "Select"
               Top             =   1680
               Width           =   2655
            End
            Begin VB.CommandButton Command22 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Generate Report In Detail"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   2400
               Style           =   1  'Graphical
               TabIndex        =   73
               Top             =   4200
               Width           =   3975
            End
            Begin VB.CommandButton Command17 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Generate Report"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   2880
               MaskColor       =   &H00FFFFC0&
               Style           =   1  'Graphical
               TabIndex        =   72
               Top             =   3600
               Width           =   3015
            End
            Begin VB.ComboBox sale1 
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Left            =   2880
               TabIndex        =   63
               Text            =   "Select"
               Top             =   960
               Width           =   2655
            End
            Begin MSComCtl2.DTPicker DT1 
               Height          =   495
               Left            =   2400
               TabIndex        =   62
               Top             =   2640
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   873
               _Version        =   393216
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Palatino Linotype"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   236650497
               UpDown          =   -1  'True
               CurrentDate     =   44402
            End
            Begin MSComCtl2.DTPicker DT2 
               Height          =   495
               Left            =   5160
               TabIndex        =   64
               Top             =   2640
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   873
               _Version        =   393216
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Palatino Linotype"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   236650497
               UpDown          =   -1  'True
               CurrentDate     =   44402
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H80000004&
               Caption         =   "Search By:"
               BeginProperty Font 
                  Name            =   "Book Antiqua"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   360
               TabIndex        =   70
               Top             =   600
               Width           =   7395
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H80000004&
               Caption         =   "Select:"
               BeginProperty Font 
                  Name            =   "Book Antiqua"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   360
               TabIndex        =   69
               Top             =   1320
               Width           =   7425
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H80000004&
               Caption         =   "Select Day:"
               BeginProperty Font 
                  Name            =   "Book Antiqua"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   360
               TabIndex        =   68
               Top             =   2160
               Width           =   7365
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000E&
               Caption         =   "To :"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   4680
               TabIndex        =   67
               Top             =   2760
               Width           =   405
            End
            Begin VB.Label Label10 
               BackColor       =   &H8000000E&
               Caption         =   "From  :"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   1440
               TabIndex        =   66
               Top             =   2640
               Width           =   2055
            End
            Begin VB.Label Label38 
               Alignment       =   2  'Center
               BackColor       =   &H00C0C000&
               Caption         =   "Sales Report"
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   510
               Left            =   -120
               TabIndex        =   65
               Top             =   0
               Width           =   8205
            End
         End
         Begin VB.Frame Frame9 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   6615
            Left            =   3000
            TabIndex        =   50
            Top             =   1080
            Visible         =   0   'False
            Width           =   8055
            Begin VB.Frame Frame10 
               BackColor       =   &H8000000E&
               BorderStyle     =   0  'None
               Height          =   4695
               Left            =   120
               TabIndex        =   51
               Top             =   600
               Width           =   7680
               Begin VB.ComboBox Combo9 
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   390
                  Left            =   2640
                  TabIndex        =   81
                  Text            =   "Select"
                  Top             =   360
                  Width           =   2535
               End
               Begin VB.CommandButton Command201 
                  BackColor       =   &H00FFFFC0&
                  Caption         =   "Generate Report In Detail"
                  BeginProperty Font 
                     Name            =   "Times New Roman"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   495
                  Left            =   1800
                  Style           =   1  'Graphical
                  TabIndex        =   75
                  Top             =   3720
                  Width           =   3975
               End
               Begin VB.CommandButton Command101 
                  BackColor       =   &H00FFFFC0&
                  Caption         =   "Generate Report"
                  BeginProperty Font 
                     Name            =   "Times New Roman"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   495
                  Left            =   2280
                  MaskColor       =   &H00FFFFC0&
                  Style           =   1  'Graphical
                  TabIndex        =   74
                  Top             =   3120
                  Width           =   3015
               End
               Begin VB.ComboBox Combo17 
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   390
                  Left            =   2640
                  TabIndex        =   52
                  Text            =   "Select"
                  Top             =   1080
                  Width           =   2655
               End
               Begin MSComCtl2.DTPicker DATE2 
                  Height          =   495
                  Left            =   5040
                  TabIndex        =   53
                  Top             =   2040
                  Width           =   1935
                  _ExtentX        =   3413
                  _ExtentY        =   873
                  _Version        =   393216
                  Enabled         =   0   'False
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Palatino Linotype"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Format          =   236584961
                  UpDown          =   -1  'True
                  CurrentDate     =   44378
               End
               Begin MSComCtl2.DTPicker DATE1 
                  DataSource      =   "DataEnvironment2"
                  Height          =   495
                  Left            =   1440
                  TabIndex        =   54
                  Top             =   2040
                  Width           =   1935
                  _ExtentX        =   3413
                  _ExtentY        =   873
                  _Version        =   393216
                  Enabled         =   0   'False
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Palatino Linotype"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Format          =   236584961
                  UpDown          =   -1  'True
                  CurrentDate     =   44378
               End
               Begin VB.Label Label40 
                  BackColor       =   &H8000000E&
                  Caption         =   "From  :"
                  BeginProperty Font 
                     Name            =   "Cambria"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   495
                  Index           =   0
                  Left            =   360
                  TabIndex        =   59
                  Top             =   2040
                  Width           =   2055
               End
               Begin VB.Label Label41 
                  AutoSize        =   -1  'True
                  BackColor       =   &H8000000E&
                  Caption         =   "To :"
                  BeginProperty Font 
                     Name            =   "Cambria"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   0
                  Left            =   4440
                  TabIndex        =   58
                  Top             =   2040
                  Width           =   405
               End
               Begin VB.Label Label43 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackColor       =   &H8000000B&
                  Caption         =   "Select Day:"
                  BeginProperty Font 
                     Name            =   "Book Antiqua"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   315
                  Index           =   0
                  Left            =   120
                  TabIndex        =   57
                  Top             =   1560
                  Width           =   7485
               End
               Begin VB.Label Label51 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackColor       =   &H8000000B&
                  Caption         =   "Search By:"
                  BeginProperty Font 
                     Name            =   "Book Antiqua"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   315
                  Left            =   120
                  TabIndex        =   56
                  Top             =   0
                  Width           =   7395
               End
               Begin VB.Label Label52 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackColor       =   &H8000000B&
                  Caption         =   "Select:"
                  BeginProperty Font 
                     Name            =   "Book Antiqua"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   315
                  Left            =   120
                  TabIndex        =   55
                  Top             =   720
                  Width           =   7425
               End
            End
            Begin VB.Label Label39 
               Alignment       =   2  'Center
               BackColor       =   &H00C0C000&
               Caption         =   "Purchase Order Report"
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   495
               Left            =   0
               TabIndex        =   60
               Top             =   0
               Width           =   8055
            End
         End
         Begin VB.Frame Frame3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   6615
            Left            =   3000
            TabIndex        =   42
            Top             =   1080
            Visible         =   0   'False
            Width           =   8055
            Begin VB.CommandButton Command16 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Generate Report"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   2520
               MaskColor       =   &H00FFFFC0&
               Style           =   1  'Graphical
               TabIndex        =   76
               Top             =   2880
               Width           =   3015
            End
            Begin VB.ComboBox Combo1 
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Left            =   2760
               TabIndex        =   45
               Text            =   "Select"
               Top             =   960
               Width           =   2655
            End
            Begin VB.ComboBox Combo2 
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Left            =   2760
               TabIndex        =   44
               Text            =   "Select"
               Top             =   1800
               Width           =   2655
            End
            Begin VB.CommandButton Command12 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Generate All Product Report"
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   1320
               Style           =   1  'Graphical
               TabIndex        =   43
               Top             =   3480
               Width           =   5295
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               Caption         =   "Select:"
               BeginProperty Font 
                  Name            =   "Book Antiqua"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   360
               TabIndex        =   49
               Top             =   1440
               Width           =   7425
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               Caption         =   "Search By:"
               BeginProperty Font 
                  Name            =   "Book Antiqua"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   360
               TabIndex        =   48
               Top             =   600
               Width           =   7395
            End
            Begin VB.Label Label30 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BeginProperty Font 
                  Name            =   "Book Antiqua"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   360
               TabIndex        =   47
               Top             =   2160
               Width           =   7260
            End
            Begin VB.Label Label44 
               Alignment       =   2  'Center
               BackColor       =   &H00C0C000&
               Caption         =   "Product  Report"
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   495
               Left            =   0
               TabIndex        =   46
               Top             =   0
               Width           =   8055
            End
         End
         Begin VB.Frame Frame2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   6615
            Left            =   3000
            TabIndex        =   35
            Top             =   1080
            Visible         =   0   'False
            Width           =   8055
            Begin VB.CommandButton Command152 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Generate Report"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   2640
               MaskColor       =   &H00E0E0E0&
               Style           =   1  'Graphical
               TabIndex        =   77
               Top             =   2280
               Width           =   3015
            End
            Begin VB.ComboBox STK2 
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Left            =   2880
               TabIndex        =   38
               Text            =   "Select"
               Top             =   1800
               Width           =   2655
            End
            Begin VB.ComboBox STK1 
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Left            =   2880
               TabIndex        =   37
               Text            =   "Select"
               Top             =   960
               Width           =   2655
            End
            Begin VB.CommandButton Command14 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Generate All Stock Report"
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   1200
               Style           =   1  'Graphical
               TabIndex        =   36
               Top             =   2880
               Width           =   5415
            End
            Begin VB.Label Label7 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               Caption         =   "Search By:"
               BeginProperty Font 
                  Name            =   "Book Antiqua"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   240
               TabIndex        =   41
               Top             =   600
               Width           =   7635
            End
            Begin VB.Label Label8 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               Caption         =   "Select:"
               BeginProperty Font 
                  Name            =   "Book Antiqua"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   240
               TabIndex        =   40
               Top             =   1440
               Width           =   7665
            End
            Begin VB.Label Label55 
               Alignment       =   2  'Center
               BackColor       =   &H00C0C000&
               Caption         =   "Stock Report"
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   495
               Left            =   0
               TabIndex        =   39
               Top             =   0
               Width           =   8055
            End
         End
         Begin VB.Frame Frame5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   6615
            Left            =   3000
            TabIndex        =   27
            Top             =   1080
            Visible         =   0   'False
            Width           =   8055
            Begin VB.CommandButton Command153 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Generate Report"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   2760
               MaskColor       =   &H00E0E0E0&
               Style           =   1  'Graphical
               TabIndex        =   78
               Top             =   3960
               Width           =   3015
            End
            Begin VB.ComboBox Combo6 
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Left            =   2760
               TabIndex        =   30
               Text            =   "Select"
               Top             =   960
               Width           =   2655
            End
            Begin VB.ComboBox Combo7 
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Left            =   2760
               TabIndex        =   29
               Text            =   "Select"
               Top             =   1800
               Width           =   2655
            End
            Begin VB.CommandButton Command24 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Generate All PurchaseStatus Report"
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   1440
               Style           =   1  'Graphical
               TabIndex        =   28
               Top             =   4560
               Width           =   5655
            End
            Begin MSComCtl2.DTPicker Date4 
               Height          =   495
               Left            =   4800
               TabIndex        =   82
               Top             =   3240
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   873
               _Version        =   393216
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Palatino Linotype"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   236847105
               UpDown          =   -1  'True
               CurrentDate     =   44378
            End
            Begin MSComCtl2.DTPicker DATE3 
               Height          =   495
               Left            =   1320
               TabIndex        =   83
               Top             =   3240
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   873
               _Version        =   393216
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Palatino Linotype"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   236847105
               UpDown          =   -1  'True
               CurrentDate     =   44378
            End
            Begin VB.Label Label43 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               Caption         =   "Select Day:"
               BeginProperty Font 
                  Name            =   "Book Antiqua"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   1
               Left            =   120
               TabIndex        =   86
               Top             =   2760
               Width           =   7485
            End
            Begin VB.Label Label41 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000E&
               Caption         =   "To :"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   1
               Left            =   3720
               TabIndex        =   85
               Top             =   3360
               Width           =   405
            End
            Begin VB.Label Label40 
               BackColor       =   &H8000000E&
               Caption         =   "From  :"
               BeginProperty Font 
                  Name            =   "Cambria"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Index           =   1
               Left            =   240
               TabIndex        =   84
               Top             =   3360
               Width           =   2055
            End
            Begin VB.Label Label13 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               Caption         =   "Select:"
               BeginProperty Font 
                  Name            =   "Book Antiqua"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   240
               TabIndex        =   34
               Top             =   1440
               Width           =   7425
            End
            Begin VB.Label Label14 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               Caption         =   "Search By:"
               BeginProperty Font 
                  Name            =   "Book Antiqua"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   240
               TabIndex        =   33
               Top             =   600
               Width           =   7515
            End
            Begin VB.Label Label27 
               Alignment       =   2  'Center
               BackColor       =   &H8000000B&
               BeginProperty Font 
                  Name            =   "Book Antiqua"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   240
               TabIndex        =   32
               Top             =   2160
               Width           =   7395
            End
            Begin VB.Label Label53 
               Alignment       =   2  'Center
               BackColor       =   &H00C0C000&
               Caption         =   "Purchase Status Report"
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   495
               Left            =   0
               TabIndex        =   31
               Top             =   0
               Width           =   8055
            End
         End
         Begin VB.Frame Frame4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   6615
            Left            =   3000
            TabIndex        =   18
            Top             =   1080
            Visible         =   0   'False
            Width           =   8055
            Begin VB.CommandButton Command155 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Generate Report"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   2280
               MaskColor       =   &H00E0E0E0&
               Style           =   1  'Graphical
               TabIndex        =   79
               Top             =   2400
               Width           =   3015
            End
            Begin VB.CommandButton Command48 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Customers With Dues"
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   1320
               Style           =   1  'Graphical
               TabIndex        =   22
               Top             =   3600
               Width           =   5415
            End
            Begin VB.CommandButton Command21 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Generate All Customer Report"
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   1320
               Style           =   1  'Graphical
               TabIndex        =   21
               Top             =   3000
               Width           =   5415
            End
            Begin VB.ComboBox Combo4 
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Left            =   2760
               TabIndex        =   20
               Text            =   "Select"
               Top             =   1680
               Width           =   2655
            End
            Begin VB.ComboBox Combo5 
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Left            =   2760
               TabIndex        =   19
               Text            =   "Select"
               Top             =   960
               Width           =   2655
            End
            Begin VB.Label Label11 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               Caption         =   "Search By:"
               BeginProperty Font 
                  Name            =   "Book Antiqua"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   120
               TabIndex        =   26
               Top             =   600
               Width           =   7635
            End
            Begin VB.Label Label12 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               Caption         =   "Select:"
               BeginProperty Font 
                  Name            =   "Book Antiqua"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   120
               TabIndex        =   25
               Top             =   1320
               Width           =   7665
            End
            Begin VB.Label Label24 
               Alignment       =   2  'Center
               BackColor       =   &H8000000B&
               BeginProperty Font 
                  Name            =   "Book Antiqua"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   120
               TabIndex        =   24
               Top             =   2040
               Width           =   7635
            End
            Begin VB.Label Label54 
               Alignment       =   2  'Center
               BackColor       =   &H00C0C000&
               Caption         =   "Customer Report"
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   495
               Left            =   0
               TabIndex        =   23
               Top             =   0
               Width           =   8055
            End
         End
         Begin VB.Frame Frame7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   6615
            Left            =   3000
            TabIndex        =   10
            Top             =   1080
            Visible         =   0   'False
            Width           =   8055
            Begin VB.CommandButton Command15 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Generate Report"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   2640
               MaskColor       =   &H00E0E0E0&
               Style           =   1  'Graphical
               TabIndex        =   80
               Top             =   2760
               Width           =   3015
            End
            Begin VB.CommandButton Command33 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Generate All Supplier Report"
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   1560
               Style           =   1  'Graphical
               TabIndex        =   13
               Top             =   3360
               Width           =   5295
            End
            Begin VB.ComboBox Combo11 
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Left            =   2880
               TabIndex        =   12
               Text            =   "Select"
               Top             =   1800
               Width           =   2655
            End
            Begin VB.ComboBox Combo12 
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Left            =   2880
               TabIndex        =   11
               Text            =   "Select"
               Top             =   960
               Width           =   2655
            End
            Begin VB.Label Label35 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BeginProperty Font 
                  Name            =   "Book Antiqua"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   240
               TabIndex        =   17
               Top             =   2280
               Width           =   7500
            End
            Begin VB.Label Label36 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               Caption         =   "Search By:"
               BeginProperty Font 
                  Name            =   "Book Antiqua"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   240
               TabIndex        =   16
               Top             =   600
               Width           =   7515
            End
            Begin VB.Label Label37 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               Caption         =   "Select:"
               BeginProperty Font 
                  Name            =   "Book Antiqua"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   240
               TabIndex        =   15
               Top             =   1440
               Width           =   7545
            End
            Begin VB.Label Label33 
               Alignment       =   2  'Center
               BackColor       =   &H00C0C000&
               Caption         =   "Supplier Report"
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   495
               Left            =   0
               TabIndex        =   14
               Top             =   0
               Width           =   8055
            End
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H8000000E&
            Caption         =   "CUSTOMER REPORT"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   1920
            Width           =   3015
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H8000000E&
            Caption         =   "SUPPLIER REPORT"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   3600
            Width           =   3015
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H8000000E&
            Caption         =   "STOCK REPORT"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   2760
            Width           =   3015
         End
         Begin VB.CommandButton Command5 
            BackColor       =   &H8000000E&
            Caption         =   "SALES REPORT"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   6120
            Width           =   3015
         End
         Begin VB.CommandButton Command29 
            BackColor       =   &H8000000E&
            Caption         =   "PURCHASE  REPORT"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   4440
            Width           =   3015
         End
         Begin VB.CommandButton Command30 
            BackColor       =   &H8000000E&
            Caption         =   "PURCHASE STATUS REPORT"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   5280
            Width           =   3015
         End
         Begin VB.CommandButton COMMAND44 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Caption         =   "CLOSE"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   6960
            Width           =   3015
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "PRODUCT REPORT"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   1080
            Width           =   3015
         End
         Begin VB.Image Image1 
            Height          =   1800
            Left            =   5160
            Picture         =   "Report.frx":0000
            Stretch         =   -1  'True
            Top             =   3120
            Width           =   2880
         End
         Begin VB.Label Label23 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "REPORTS"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   27.75
               Charset         =   186
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   975
            Left            =   0
            TabIndex        =   71
            Top             =   120
            Width           =   11055
         End
      End
   End
End
Attribute VB_Name = "Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
If Combo1.Text = "ID" Then
Combo2.Clear
auto_p_id
ElseIf Combo1.Text = "NAME" Then
Combo2.Clear
auto_P_nm
ElseIf Combo1.Text = "COMPANY" Then
Combo2.Clear
auto_P_comp
ElseIf Combo1.Text = "TYPE" Then
Combo2.Clear
auto_P_type
End If


End Sub

Private Sub Combo10_CLICK()
Command17.Enabled = True
Command17.SetFocus
Command22.Enabled = True
End Sub

Private Sub Combo12_CLICK()
If Combo12.Text = "ID" Then
Combo11.Clear
auto_sup_id
ElseIf Combo12.Text = "NAME" Then
Combo11.Clear
auto_sup_nm
ElseIf Combo12.Text = "MOBILE NO." Then
Combo11.Clear
auto_sup_mob
End If
End Sub





Private Sub Combo5_Click()
If Combo5.Text = "ID" Then
Combo4.Clear
auto_cust_id

ElseIf Combo5.Text = "MOBILE" Then
Combo4.Clear
auto_cust_mob

ElseIf Combo5.Text = "NAME" Then
Combo4.Clear
auto_cust_nm



End If

End Sub

Private Sub Combo6_Click()
If Combo6.Text = "Order No" Then
 Combo7.Clear
 auto_orderno
ElseIf Combo6.Text = "Invoice No." Then
 Combo7.Clear
 auto_invno
ElseIf Combo6.Text = "Between Dates" Then
 Combo7.Clear
 DATE3.Enabled = True
 Date4.Enabled = True
ElseIf Combo6.Text = "Invoice Date" Then
 Combo7.Clear
 auto_invdate
ElseIf Combo6.Text = "Month" Then
Combo7.Clear
Set R = New ADODB.Recordset
SQL = "select distinct upper(to_char(invdate,'MON')) from ordetails"
Set R = C.Execute(SQL)
While R.EOF = False
Combo7.AddItem R.Fields(0)
R.MoveNext
Wend
End If

End Sub

Private Sub Combo9_Click()
If Combo9.Text = "Order No" Then
 Combo17.Clear
 auto_pur_orderno
ElseIf Combo9.Text = "Supplier Id" Then
 Combo17.Clear
 auto_pur_sup_id
ElseIf Combo9.Text = "Between Dates" Then
 Combo17.Clear
 DATE1.Enabled = True
 
 DATE2.Enabled = True
ElseIf Combo9.Text = "Date" Then
 Combo17.Clear
 auto_pur_date
ElseIf Combo9.Text = "Month" Then
Combo17.Clear
Set R = New ADODB.Recordset
SQL = "select distinct upper(to_char(pur_orderdate,'MON')) from purordetail"
Set R = C.Execute(SQL)
While R.EOF = False
Combo17.AddItem R.Fields(0)
R.MoveNext
Wend
End If
End Sub

Private Sub Command1_Click()
frame3.Visible = True
Frame101.Visible = False
Frame7.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Frame2.Visible = False
'Frame11.Visible = False
Combo1.Clear
Frame9.Visible = False
Combo1.AddItem "ID"
Combo1.AddItem "NAME"
Combo1.AddItem "COMPANY"
Combo1.AddItem "TYPE"
End Sub

Private Sub Command101_Click()

If Combo9.Text = "Between Dates" Then
    DataEnvironment2.Connection1.Open
    DataEnvironment2.Command14 DATE1.Value, DATE2.Value
    pur_btw_date.Show
    pur_btw_date.Refresh
    DataEnvironment2.Connection1.Close


ElseIf Combo9.Text = blank Or Combo17.Text = blank Then
MsgBox "Please select the parameters first.!!"
Else
If Combo9.Text = "Order No" Then

    DataEnvironment2.Connection1.Open
    DataEnvironment2.Command10 Combo17.Text
    pur_orderno.Show
    pur_orderno.Refresh
    DataEnvironment2.Connection1.Close

ElseIf Combo9.Text = "Supplier Id" Then
    DataEnvironment2.Connection1.Open
    DataEnvironment2.Command11 Combo17.Text
    pur_sup_id.Show
    pur_sup_id.Refresh
    DataEnvironment2.Connection1.Close

ElseIf Combo9.Text = "Date" Then
    DataEnvironment2.Connection1.Open
    DataEnvironment2.Command12 Combo17.Text
    Pur_date.Show
    Pur_date.Refresh
    DataEnvironment2.Connection1.Close
    
ElseIf Combo9.Text = "Month" Then
    DataEnvironment2.Connection1.Open
    DataEnvironment2.Command13 Combo17.Text
    pur_month.Show
    pur_month.Refresh
    DataEnvironment2.Connection1.Close
    
End If
End If

End Sub

Private Sub Command12_Click()
allPrd_Report.Show
End Sub

Private Sub Command14_Click()
allstock_Report.Show
End Sub

Private Sub Command15_Click()
If Combo12.Text = blank Or Combo11.Text = blank Then
MsgBox "Please select the parameters first.!!"
Else
If Combo12.Text = "ID" Then
    DataEnvironment2.Connection1.Open
    DataEnvironment2.Command3 Combo11.Text
    Suppl_report.Show
    Suppl_report.Refresh
    DataEnvironment2.Connection1.Close

ElseIf Combo12.Text = "NAME" Then
    DataEnvironment2.Connection1.Open
    DataEnvironment2.Command1 Combo11.Text
    suppl_report_nm.Show
    suppl_report_nm.Refresh
    DataEnvironment2.Connection1.Close
    
ElseIf Combo12.Text = "MOBILE NO." Then
    DataEnvironment2.Connection1.Open
    DataEnvironment2.Command2 Combo11.Text
    suppl_Report_mob.Show
    suppl_Report_mob.Refresh
    DataEnvironment2.Connection1.Close


End If
End If
End Sub

Private Sub Command152_Click()
If STK1.Text = blank Or STK2.Text = blank Then
MsgBox "Please select the parameters first.!!"
Else
If STK1.Text = "ID" Then
DataEnvironment2.Connection1.Open
   DataEnvironment2.Command8 STK2.Text
    Stock_pid.Show
    Stock_pid.Refresh
    DataEnvironment2.Connection1.Close
    
ElseIf STK1.Text = "RACK NO" Then
DataEnvironment2.Connection1.Open
    DataEnvironment2.Command9 STK2.Text
       stock_Report_rackno.Show
       stock_Report_rackno.Refresh
    DataEnvironment2.Connection1.Close
    End If
    End If


End Sub

Private Sub Command153_Click()

If Combo6.Text = "Between Dates" Then

    DataEnvironment2.Connection1.Open
    DataEnvironment2.Command22 DATE3.Value, Date4.Value
    purst_btwdate.Show
    purst_btwdate.Refresh
    DataEnvironment2.Connection1.Close

ElseIf Combo6.Text = blank Or Combo7.Text = blank Then
MsgBox "Please select the parameters first.!!"
Else
If Combo6.Text = "Order No" Then

    DataEnvironment2.Connection1.Open
    DataEnvironment2.Command20 Combo7.Text
    purst_ordno.Show
    purst_ordno.Refresh
    DataEnvironment2.Connection1.Close
    
ElseIf Combo6.Text = "Invoice No." Then

    DataEnvironment2.Connection1.Open
    DataEnvironment2.Command18 Combo7.Text
    purst_invno.Show
    purst_invno.Refresh
    DataEnvironment2.Connection1.Close

ElseIf Combo6.Text = "Invoice Date" Then

    DataEnvironment2.Connection1.Open
    DataEnvironment2.Command19 Combo7.Text
    purst_invdate.Show
    purst_invdate.Refresh
    DataEnvironment2.Connection1.Close
    
ElseIf Combo6.Text = "Month" Then

    DataEnvironment2.Connection1.Open
    DataEnvironment2.Command21 Combo7.Text
    purst_month.Show
    purst_month.Refresh
    DataEnvironment2.Connection1.Close
    

End If
End If
End Sub

Private Sub Command155_Click()
If Combo5.Text = blank Or Combo4.Text = blank Then
MsgBox "Please select the parameters first.!!"
Else
If Combo5.Text = "ID" Then
    DataEnvironment2.Connection1.Open
    DataEnvironment2.Command15 Combo4.Text
    cust_Report_id.Show
    cust_Report_id.Refresh
    DataEnvironment2.Connection1.Close
    
ElseIf Combo5.Text = "MOBILE" Then
    DataEnvironment2.Connection1.Open
    DataEnvironment2.Command17 Combo4.Text
    cust_report_mob.Show
    cust_report_mob.Refresh
    DataEnvironment2.Connection1.Close
    
ElseIf Combo5.Text = "NAME" Then
    DataEnvironment2.Connection1.Open
    DataEnvironment2.Command16 Combo4.Text
    cust_report_nm.Show
    cust_report_nm.Refresh
    DataEnvironment2.Connection1.Close
End If
End If
End Sub

Private Sub Command16_Click()
If Combo1.Text = blank Or Combo2.Text = blank Then
MsgBox "Please select the parameters first.!!"
Else
If Combo1.Text = "ID" Then
    DataEnvironment2.Connection1.Open
    DataEnvironment2.Command4 Combo2.Text
    prd_report_id.Show
    prd_report_id.Refresh
    DataEnvironment2.Connection1.Close

ElseIf Combo1.Text = "NAME" Then
    DataEnvironment2.Connection1.Open
    DataEnvironment2.Command5 Combo2.Text
    prd_report_nm.Show
    prd_report_nm.Refresh
    DataEnvironment2.Connection1.Close
ElseIf Combo1.Text = "COMPANY" Then
    DataEnvironment2.Connection1.Open
    DataEnvironment2.Command6 Combo2.Text
    prd_report_comp.Show
    prd_report_comp.Refresh
    DataEnvironment2.Connection1.Close

ElseIf Combo1.Text = "TYPE" Then
    DataEnvironment2.Connection1.Open
    DataEnvironment2.Command7 Combo2.Text
    prd_report_type.Show
    prd_report_type.Refresh
    DataEnvironment2.Connection1.Close
End If
End If
End Sub

Private Sub Command17_Click()
If sale1.Text = "Between Dates" Then
    DataEnvironment2.Connection1.Open
    DataEnvironment2.Command27 DT1.Value, DT2.Value
    selReport_btw_date.Show
    selReport_btw_date.Refresh
    DataEnvironment2.Connection1.Close

ElseIf sale1.Text = blank Or Combo10.Text = blank Then
MsgBox "Please select the parameters first.!!"
Else
If sale1.Text = "Order Id" Then
    DataEnvironment2.Connection1.Open
    DataEnvironment2.Command23 Combo10.Text
    sel_Report_ordno.Show
    sel_Report_ordno.Refresh
    DataEnvironment2.Connection1.Close
    
ElseIf sale1.Text = "Customer Id" Then
    DataEnvironment2.Connection1.Open
    DataEnvironment2.Command24 Combo10.Text
    selReport_cust_id.Show
    selReport_cust_id.Refresh
    DataEnvironment2.Connection1.Close
    
ElseIf sale1.Text = "Date" Then
    DataEnvironment2.Connection1.Open
    DataEnvironment2.Command25 Combo10.Text
    selReport_date.Show
    selReport_date.Refresh
    DataEnvironment2.Connection1.Close
      
 ElseIf sale1.Text = "Month" Then
    DataEnvironment2.Connection1.Open
    DataEnvironment2.Command26 Combo10.Text
    selReport_month.Show
    selReport_month.Refresh
    DataEnvironment2.Connection1.Close
   End If
   End If
  
End Sub

Private Sub Command2_Click()
Frame4.Visible = True
Frame101.Visible = False
Frame7.Visible = False

Frame5.Visible = False
Frame2.Visible = False
frame3.Visible = False
Frame9.Visible = False
Combo5.Clear
Combo5.AddItem "ID"
Combo5.AddItem "MOBILE"
Combo5.AddItem "NAME"

End Sub

Private Sub Command20_Click()
DataReport2.Show
End Sub

Private Sub Command201_Click()
Allpurord_Report.Show
End Sub

Private Sub Command21_Click()
allcust_Report.Show
End Sub

Private Sub Command22_Click()
Allsel_report.Show

End Sub

Private Sub Command24_Click()
All_ordetail_report.Show
End Sub

Private Sub Command29_Click()
Frame9.Visible = True
Frame101.Visible = False
Frame7.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Frame2.Visible = False
'Frame11.Visible = False
frame3.Visible = False

Combo9.Clear
Combo9.AddItem "Order No"
Combo9.AddItem "Supplier Id"
Combo9.AddItem "Date"
Combo9.AddItem "Month"
Combo9.AddItem "Between Dates"


End Sub

Private Sub Command3_Click()
Frame7.Visible = True
Frame101.Visible = False

Frame4.Visible = False
Frame5.Visible = False
Frame2.Visible = False
'Frame11.Visible = False
frame3.Visible = False
Frame9.Visible = False
Combo12.Clear
Combo12.AddItem "ID"
Combo12.AddItem "NAME"
Combo12.AddItem "MOBILE NO."

End Sub

Private Sub Command30_Click()
Frame5.Visible = True
Frame101.Visible = False
Frame7.Visible = False
Frame4.Visible = False

Frame2.Visible = False
'Frame11.Visible = False
frame3.Visible = False
Frame9.Visible = False
Combo6.Clear
Combo6.AddItem "Order No"
Combo6.AddItem "Invoice No."
Combo6.AddItem "Invoice Date"
Combo6.AddItem "Month"
Combo6.AddItem "Between Dates"
End Sub

Private Sub Command33_Click()
AllSuppl_report.Show

End Sub

Private Sub Command4_Click()
Frame2.Visible = True
Frame101.Visible = False
Frame7.Visible = False
Frame4.Visible = False
Frame5.Visible = False

'Frame11.Visible = False
frame3.Visible = False
Frame9.Visible = False
STK1.Clear

STK1.AddItem "ID"
STK1.AddItem "RACK NO"

End Sub

Private Sub COMMAND44_Click()
Unload Me
End Sub

Private Sub Command45_Click()
'Frame11.Visible = True
Frame101.Visible = False
Frame7.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Frame2.Visible = False

frame3.Visible = False
Frame9.Visible = False
End Sub


Private Sub Command48_Click()
cust_report_dues.Show
End Sub

Private Sub Command5_Click()
Frame101.Visible = True

Frame7.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Frame2.Visible = False
'Frame11.Visible = False
frame3.Visible = False
Frame9.Visible = False
sale1.Clear
sale1.AddItem "Order Id"
sale1.AddItem "Customer Id"
sale1.AddItem "Date"
sale1.AddItem "Month"
sale1.AddItem "Between Dates"


End Sub

Private Sub Command6_Click()
Frame6.Visible = True
Frame101.Visible = False
Frame7.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Frame2.Visible = False
'Frame11.Visible = False

End Sub

Private Sub Form_Load()
CONN
Frame101.Visible = False
Frame7.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Frame2.Visible = False
'Frame11.Visible = False
frame3.Visible = False
Frame9.Visible = False
If (DataEnvironment2.Connection1.State) Then
DataEnvironment2.Connection1.Close
End If


End Sub



Public Function auto_sup_id()
Set R = New ADODB.Recordset
SQL = "select *from supplier"
Set R = C.Execute(SQL)
While R.EOF = False
Combo11.AddItem R.Fields(0)
R.MoveNext
Wend
End Function
Public Function auto_sup_nm()
Set R = New ADODB.Recordset
SQL = "select *from supplier"
Set R = C.Execute(SQL)
While R.EOF = False
Combo11.AddItem R.Fields(1)
R.MoveNext
Wend
End Function


Public Function auto_sup_mob()
Set R = New ADODB.Recordset
SQL = "select *from supplier"
Set R = C.Execute(SQL)
While R.EOF = False
Combo11.AddItem R.Fields(2)
R.MoveNext
Wend


End Function


Public Function auto_p_id()
Set R = New ADODB.Recordset
SQL = "select *from product"
Set R = C.Execute(SQL)
While R.EOF = False
Combo2.AddItem R.Fields(0)
R.MoveNext
Wend
End Function
Public Function auto_P_nm()
Set R = New ADODB.Recordset
SQL = "select *from product"
Set R = C.Execute(SQL)
While R.EOF = False
Combo2.AddItem R.Fields(1)
R.MoveNext
Wend
End Function
Public Function auto_P_comp()
Set R = New ADODB.Recordset
SQL = "select p_comp from product"
Set R = C.Execute(SQL)
While R.EOF = False
Combo2.AddItem R.Fields(0)
R.MoveNext
Wend
End Function
Public Function auto_P_type()
Set R = New ADODB.Recordset
SQL = "select p_type from product"
Set R = C.Execute(SQL)
While R.EOF = False
Combo2.AddItem R.Fields(0)
R.MoveNext
Wend
End Function
Public Function auto_pur_orderno()
Set R = New ADODB.Recordset
SQL = "select *from purordetail order by pur_orderno"
Set R = C.Execute(SQL)
While R.EOF = False
Combo17.AddItem R.Fields(0)
R.MoveNext
Wend

End Function
Public Function auto_orderno()
Set R = New ADODB.Recordset
SQL = "select *from ordetails order by orderno"
Set R = C.Execute(SQL)
While R.EOF = False
Combo7.AddItem R.Fields(2)
R.MoveNext
Wend
End Function
Public Function auto_invno()
Set R = New ADODB.Recordset
SQL = "select *from ordetails order by invoiceno"
Set R = C.Execute(SQL)
While R.EOF = False
Combo7.AddItem R.Fields(0)
R.MoveNext
Wend
End Function
Public Function auto_invdate()
Set R = New ADODB.Recordset
SQL = "select *from ordetails order by invdate"
Set R = C.Execute(SQL)
While R.EOF = False
Combo7.AddItem R.Fields(1)
R.MoveNext
Wend
End Function
Public Function auto_pur_date()
Set R = New ADODB.Recordset
SQL = "select *from purordetail order by pur_orderdate"
Set R = C.Execute(SQL)
While R.EOF = False
Combo17.AddItem R.Fields(1)
R.MoveNext
Wend
End Function
Public Function auto_pur_sup_id()
Set R = New ADODB.Recordset
SQL = "select distinct sup_id from purordetail order by sup_id"
Set R = C.Execute(SQL)
While R.EOF = False
Combo17.AddItem R.Fields(0)
R.MoveNext
Wend
End Function
Public Function auto_cust_id()
Set R = New ADODB.Recordset
SQL = "select *from customer"
Set R = C.Execute(SQL)
While R.EOF = False
Combo4.AddItem R.Fields(0)
R.MoveNext
Wend
End Function
Public Function auto_cust_nm()
Set R = New ADODB.Recordset
SQL = "select *from customer"
Set R = C.Execute(SQL)
While R.EOF = False
Combo4.AddItem R.Fields(1)
R.MoveNext
Wend
End Function
Public Function auto_cust_mob()
Set R = New ADODB.Recordset
SQL = "select *from customer"
Set R = C.Execute(SQL)
While R.EOF = False
Combo4.AddItem R.Fields(2)
R.MoveNext
Wend
End Function


Private Sub sale1_Click()
If sale1.Text = "Order Id" Then
 Combo10.Clear
 auto_sell_orderid
ElseIf sale1.Text = "Customer Id" Then
 Combo10.Clear
 auto_sell_cust_id
ElseIf sale1.Text = "Between Dates" Then
 Combo10.Clear
 DT1.Enabled = True
 Format (DT1.Value = "dd-mmm-yyyy")
 DT2.Enabled = True
ElseIf sale1.Text = "Date" Then
 Combo10.Clear
 auto_sell_date
ElseIf sale1.Text = "Month" Then
Combo10.Clear
Set R = New ADODB.Recordset
SQL = "select distinct  upper (to_char(s_date, 'MON') ) from sell_details"
Set R = C.Execute(SQL)
While R.EOF = False
Combo10.AddItem R.Fields(0)
R.MoveNext
Wend
End If
End Sub




Private Sub STK1_click()
If STK1.Text = "ID" Then
STK2.Clear
auto_stock_id

ElseIf STK1.Text = "RACK NO" Then
STK2.Clear
auto_stock_rno
End If
End Sub
Public Function auto_stock_id()
Set R = New ADODB.Recordset
SQL = "select *from stock"
Set R = C.Execute(SQL)
While R.EOF = False
STK2.AddItem R.Fields(1)
R.MoveNext
Wend
End Function
Public Function auto_stock_rno()
Set R = New ADODB.Recordset
SQL = "select *from stock"
Set R = C.Execute(SQL)
While R.EOF = False
STK2.AddItem R.Fields(0)
R.MoveNext
Wend
End Function

Public Function auto_sell_orderid()

Set R = New ADODB.Recordset
SQL = "select *from sell_details"
Set R = C.Execute(SQL)
While R.EOF = False
Combo10.AddItem R.Fields(0)
R.MoveNext
Wend
End Function

Public Function auto_sell_cust_id()
Set R = New ADODB.Recordset
SQL = "select *from sell_details"
Set R = C.Execute(SQL)
While R.EOF = False
Combo10.AddItem R.Fields(2)
R.MoveNext
Wend
End Function
Public Function auto_cust_dues()
Set R = New ADODB.Recordset
SQL = "select *from customer"
Set R = C.Execute(SQL)
While R.EOF = False
Combo4.AddItem R.Fields(6)
R.MoveNext
Wend
End Function

Public Function auto_sell_date()
Set R = New ADODB.Recordset
SQL = "select *from sell_details"
Set R = C.Execute(SQL)
While R.EOF = False
Combo10.AddItem R.Fields(1)
R.MoveNext
Wend
End Function

