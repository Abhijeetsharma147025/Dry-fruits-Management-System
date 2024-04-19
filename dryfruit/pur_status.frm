VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form purchase 
   Caption         =   "Purchase"
   ClientHeight    =   9450
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9450
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame6 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Caption         =   "-"
      ForeColor       =   &H00008000&
      Height          =   9375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20415
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   12480
         Style           =   1  'Graphical
         TabIndex        =   138
         Top             =   4680
         Width           =   855
      End
      Begin MSComCtl2.MonthView MonthView2 
         Height          =   2370
         Left            =   16560
         TabIndex        =   137
         Top             =   360
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   113377282
         CurrentDate     =   44442
      End
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   2370
         Left            =   12960
         TabIndex        =   136
         Top             =   600
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   113377282
         CurrentDate     =   44442
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   14640
         TabIndex        =   135
         Top             =   9000
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox Text5 
         Height          =   615
         Left            =   5160
         TabIndex        =   134
         Top             =   8400
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   16560
         TabIndex        =   133
         TabStop         =   0   'False
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   16560
         TabIndex        =   131
         TabStop         =   0   'False
         Top             =   0
         Width           =   1815
      End
      Begin VB.TextBox Text4 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   16320
         TabIndex        =   128
         Top             =   6480
         Width           =   1335
      End
      Begin VB.TextBox gst 
         Height          =   615
         Left            =   14760
         TabIndex        =   127
         Top             =   4560
         Width           =   735
      End
      Begin VB.ListBox List12 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1560
         ItemData        =   "pur_status.frx":0000
         Left            =   10800
         List            =   "pur_status.frx":0002
         TabIndex        =   126
         TabStop         =   0   'False
         Top             =   6000
         Width           =   615
      End
      Begin VB.ListBox List11 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1560
         ItemData        =   "pur_status.frx":0004
         Left            =   12840
         List            =   "pur_status.frx":0006
         TabIndex        =   125
         TabStop         =   0   'False
         Top             =   6000
         Width           =   855
      End
      Begin VB.ListBox List10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1560
         ItemData        =   "pur_status.frx":0008
         Left            =   11400
         List            =   "pur_status.frx":000A
         TabIndex        =   124
         TabStop         =   0   'False
         Top             =   6000
         Width           =   855
      End
      Begin VB.ListBox List9 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1560
         ItemData        =   "pur_status.frx":000C
         Left            =   13680
         List            =   "pur_status.frx":000E
         TabIndex        =   123
         TabStop         =   0   'False
         Top             =   6000
         Width           =   1335
      End
      Begin VB.ListBox List8 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1560
         ItemData        =   "pur_status.frx":0010
         Left            =   720
         List            =   "pur_status.frx":0012
         TabIndex        =   122
         TabStop         =   0   'False
         Top             =   6000
         Width           =   1215
      End
      Begin VB.ListBox List7 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1560
         ItemData        =   "pur_status.frx":0014
         Left            =   9600
         List            =   "pur_status.frx":0016
         TabIndex        =   121
         TabStop         =   0   'False
         Top             =   6000
         Width           =   1215
      End
      Begin VB.ListBox List6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1560
         ItemData        =   "pur_status.frx":0018
         Left            =   8400
         List            =   "pur_status.frx":001A
         TabIndex        =   120
         TabStop         =   0   'False
         Top             =   6000
         Width           =   1215
      End
      Begin VB.ListBox List5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1560
         ItemData        =   "pur_status.frx":001C
         Left            =   7200
         List            =   "pur_status.frx":001E
         TabIndex        =   119
         TabStop         =   0   'False
         Top             =   6000
         Width           =   1215
      End
      Begin VB.ListBox List4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1560
         ItemData        =   "pur_status.frx":0020
         Left            =   5280
         List            =   "pur_status.frx":0022
         TabIndex        =   118
         TabStop         =   0   'False
         Top             =   6000
         Width           =   1935
      End
      Begin VB.ListBox List3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1560
         ItemData        =   "pur_status.frx":0024
         Left            =   12240
         List            =   "pur_status.frx":0026
         TabIndex        =   117
         TabStop         =   0   'False
         Top             =   6000
         Width           =   615
      End
      Begin VB.ListBox List2 
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
         Height          =   1590
         ItemData        =   "pur_status.frx":0028
         Left            =   1920
         List            =   "pur_status.frx":002A
         TabIndex        =   116
         Top             =   6000
         Width           =   3375
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
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
         Height          =   1560
         ItemData        =   "pur_status.frx":002C
         Left            =   120
         List            =   "pur_status.frx":002E
         TabIndex        =   115
         TabStop         =   0   'False
         Top             =   6000
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
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
         Left            =   120
         TabIndex        =   114
         Top             =   4920
         Width           =   1695
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   113
         Top             =   4920
         Width           =   3135
      End
      Begin VB.TextBox txt19 
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
         Left            =   1800
         TabIndex        =   84
         Top             =   7800
         Width           =   1215
      End
      Begin VB.TextBox txt8 
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   13080
         Locked          =   -1  'True
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1815
      End
      Begin VB.ListBox pnm 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000007&
         Height          =   1305
         ItemData        =   "pur_status.frx":0030
         Left            =   1920
         List            =   "pur_status.frx":0032
         TabIndex        =   82
         Top             =   2760
         Width           =   3375
      End
      Begin VB.CommandButton ok 
         BackColor       =   &H00FFFFC0&
         Caption         =   "ADD"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   11400
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   4680
         Width           =   855
      End
      Begin VB.TextBox txt14 
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
         Left            =   8400
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   4920
         Width           =   1095
      End
      Begin VB.TextBox txt6 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   13080
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   120
         Width           =   1815
      End
      Begin VB.TextBox txt2 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox txt1 
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   600
         Width           =   2175
      End
      Begin VB.ComboBox odno 
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
         Height          =   375
         Left            =   2280
         TabIndex        =   32
         Top             =   120
         Width           =   2415
      End
      Begin VB.TextBox txt4 
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txt5 
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1080
         Width           =   2175
      End
      Begin VB.ListBox sgst 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1305
         ItemData        =   "pur_status.frx":0034
         Left            =   12240
         List            =   "pur_status.frx":0036
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   2760
         Width           =   615
      End
      Begin VB.ListBox ptyp 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1305
         ItemData        =   "pur_status.frx":0038
         Left            =   5280
         List            =   "pur_status.frx":003A
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1935
      End
      Begin VB.ListBox prate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1305
         ItemData        =   "pur_status.frx":003C
         Left            =   7200
         List            =   "pur_status.frx":003E
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1215
      End
      Begin VB.ListBox qty1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1305
         ItemData        =   "pur_status.frx":0040
         Left            =   8400
         List            =   "pur_status.frx":0042
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1215
      End
      Begin VB.ListBox prc 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1305
         ItemData        =   "pur_status.frx":0044
         Left            =   9600
         List            =   "pur_status.frx":0046
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1215
      End
      Begin VB.ListBox pid 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1305
         ItemData        =   "pur_status.frx":0048
         Left            =   720
         List            =   "pur_status.frx":004A
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1215
      End
      Begin VB.ListBox sr 
         Appearance      =   0  'Flat
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
         Height          =   1305
         ItemData        =   "pur_status.frx":004C
         Left            =   120
         List            =   "pur_status.frx":004E
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2760
         Width           =   615
      End
      Begin VB.ListBox totprc 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1305
         ItemData        =   "pur_status.frx":0050
         Left            =   13680
         List            =   "pur_status.frx":0052
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1455
      End
      Begin VB.ListBox camt 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1305
         ItemData        =   "pur_status.frx":0054
         Left            =   11400
         List            =   "pur_status.frx":0056
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   2760
         Width           =   855
      End
      Begin VB.ListBox samt 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1305
         ItemData        =   "pur_status.frx":0058
         Left            =   12840
         List            =   "pur_status.frx":005A
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2760
         Width           =   855
      End
      Begin VB.ListBox cgst 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1305
         ItemData        =   "pur_status.frx":005C
         Left            =   10800
         List            =   "pur_status.frx":005E
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2760
         Width           =   615
      End
      Begin VB.TextBox txt11 
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4920
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   4920
         Width           =   2415
      End
      Begin VB.OptionButton ST01 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Completely Received"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11160
         TabIndex        =   17
         Top             =   7800
         Width           =   2775
      End
      Begin VB.OptionButton ST02 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Not Received"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   13920
         TabIndex        =   16
         Top             =   7800
         Width           =   2775
      End
      Begin VB.OptionButton ST03 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Partially Received"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   16680
         TabIndex        =   15
         Top             =   7800
         Width           =   2655
      End
      Begin VB.TextBox txt15 
         BackColor       =   &H8000000B&
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
         Left            =   9480
         MaxLength       =   4
         TabIndex        =   14
         Top             =   4920
         Width           =   1335
      End
      Begin VB.TextBox txt22 
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
         Left            =   16320
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   8280
         Width           =   1695
      End
      Begin VB.TextBox txt23 
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   16320
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   " "
         Top             =   8640
         Width           =   1695
      End
      Begin VB.TextBox txt21 
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
         Left            =   11160
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   8400
         Width           =   1695
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
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   855
         Left            =   16200
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox txt13 
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
         Left            =   7320
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   4920
         Width           =   1050
      End
      Begin VB.TextBox txt7 
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   13080
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton save 
         BackColor       =   &H00FFFF00&
         Caption         =   "Save"
         DisabledPicture =   "pur_status.frx":0060
         DownPicture     =   "pur_status.frx":0B5E
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   18360
         MaskColor       =   &H00C0FFC0&
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   8280
         Width           =   1455
      End
      Begin VB.CommandButton CANCEL 
         BackColor       =   &H00FFFF00&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   18360
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   8760
         Width           =   1455
      End
      Begin VB.TextBox txt3 
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   120
         Width           =   2175
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   1800
         TabIndex        =   2
         Top             =   8160
         Width           =   2655
         Begin VB.OptionButton ST121 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Caption         =   "Cheque"
            ForeColor       =   &H80000008&
            Height          =   435
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   1335
         End
         Begin VB.OptionButton ST122 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Caption         =   "CASH"
            ForeColor       =   &H80000008&
            Height          =   435
            Left            =   1320
            TabIndex        =   3
            Top             =   0
            Width           =   1335
         End
      End
      Begin VB.TextBox txt20 
         BackColor       =   &H8000000B&
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   8640
         Width           =   2775
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "Final Amount :"
         ForeColor       =   &H000080FF&
         Height          =   675
         Left            =   16320
         TabIndex        =   139
         Top             =   6000
         Width           =   1035
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Invoice Date :"
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
         Left            =   15120
         TabIndex        =   132
         Top             =   600
         Width           =   1440
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Invoice No. :"
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
         Left            =   15120
         TabIndex        =   130
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "PRODUCT RECIEVED"
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
         Left            =   0
         TabIndex        =   129
         Top             =   4200
         Width           =   1860
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000080&
         X1              =   240
         X2              =   20160
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Line Line2 
         X1              =   4080
         X2              =   4200
         Y1              =   5760
         Y2              =   5760
      End
      Begin VB.Label Label62 
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
         Left            =   12240
         TabIndex        =   112
         Top             =   5400
         Width           =   1455
      End
      Begin VB.Label Label60 
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
         Left            =   11400
         TabIndex        =   111
         Top             =   5760
         Width           =   855
      End
      Begin VB.Label Label56 
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
         Left            =   12840
         TabIndex        =   110
         Top             =   5760
         Width           =   855
      End
      Begin VB.Label Label55 
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
         Left            =   12240
         TabIndex        =   109
         Top             =   5760
         Width           =   615
      End
      Begin VB.Label Label53 
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
         Left            =   13680
         TabIndex        =   108
         Top             =   5400
         Width           =   1335
      End
      Begin VB.Label Label52 
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
         Left            =   9600
         TabIndex        =   107
         Top             =   5400
         Width           =   1215
      End
      Begin VB.Label Label51 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Quantity Recieved"
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
         Left            =   8400
         TabIndex        =   106
         Top             =   5400
         Width           =   1215
      End
      Begin VB.Label Label50 
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
         Left            =   7200
         TabIndex        =   105
         Top             =   5400
         Width           =   1215
      End
      Begin VB.Label Label49 
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
         Left            =   5280
         TabIndex        =   104
         Top             =   5400
         Width           =   1935
      End
      Begin VB.Label Label48 
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
         Left            =   1920
         TabIndex        =   103
         Top             =   5400
         Width           =   3375
      End
      Begin VB.Label Label47 
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
         Left            =   720
         TabIndex        =   102
         Top             =   5400
         Width           =   1215
      End
      Begin VB.Label Label46 
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
         Left            =   10800
         TabIndex        =   101
         Top             =   5400
         Width           =   1455
      End
      Begin VB.Label Label45 
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
         Left            =   10800
         TabIndex        =   100
         Top             =   5760
         Width           =   615
      End
      Begin VB.Label Label44 
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
         Left            =   120
         TabIndex        =   99
         Top             =   5400
         Width           =   615
      End
      Begin VB.Label Label89 
         BackColor       =   &H8000000E&
         Caption         =   "Cheque No.     :"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   375
         Left            =   120
         TabIndex        =   80
         Top             =   8640
         Width           =   2295
      End
      Begin VB.Label Label72 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Supply Date             :"
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
         Left            =   10800
         TabIndex        =   79
         Top             =   240
         Width           =   2145
      End
      Begin VB.Label Label71 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Order no               :"
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
         Left            =   240
         TabIndex        =   78
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label112 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Supplier ID          :"
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
         Left            =   240
         TabIndex        =   77
         Top             =   720
         Width           =   1905
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Supplier Address    :"
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
         Left            =   10800
         TabIndex        =   76
         Top             =   1200
         Width           =   2145
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Supplier Name    :"
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
         Left            =   240
         TabIndex        =   75
         Top             =   1200
         Width           =   1905
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Mobile No.             :"
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
         Left            =   6000
         TabIndex        =   74
         Top             =   720
         Width           =   2010
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Company Name     :"
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
         Left            =   5925
         TabIndex        =   73
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label73 
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
         Left            =   12600
         TabIndex        =   72
         Top             =   0
         Width           =   135
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
         Index           =   3
         Left            =   1200
         TabIndex        =   71
         Top             =   0
         Width           =   135
      End
      Begin VB.Label Label11 
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
         Left            =   10800
         TabIndex        =   70
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label32 
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
         Left            =   10800
         TabIndex        =   69
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label33 
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
         Left            =   720
         TabIndex        =   68
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label34 
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
         Left            =   1920
         TabIndex        =   67
         Top             =   2160
         Width           =   3375
      End
      Begin VB.Label Label35 
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
         Left            =   5280
         TabIndex        =   66
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label36 
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
         Left            =   120
         TabIndex        =   65
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label38 
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
         Left            =   7200
         TabIndex        =   64
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label39 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Quantity Ordered"
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
         Left            =   8400
         TabIndex        =   63
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label54 
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
         Left            =   9600
         TabIndex        =   62
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label58 
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
         Left            =   13680
         TabIndex        =   61
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label65 
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
         Left            =   12240
         TabIndex        =   60
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label103 
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
         Left            =   12840
         TabIndex        =   59
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label104 
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
         Left            =   11400
         TabIndex        =   58
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label105 
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
         Left            =   12240
         TabIndex        =   57
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label59 
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
         Left            =   2475
         TabIndex        =   56
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Label Label61 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Product Type"
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
         Left            =   5265
         TabIndex        =   55
         Top             =   4560
         Width           =   1425
      End
      Begin VB.Label Label81 
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
         Left            =   225
         TabIndex        =   54
         Top             =   4560
         Width           =   1185
      End
      Begin VB.Label Label82 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "STATUS          :"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9075
         TabIndex        =   53
         Top             =   7920
         Width           =   1215
      End
      Begin VB.Label Label57 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "Quantity Received :"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   9600
         TabIndex        =   52
         Top             =   4320
         Width           =   1065
      End
      Begin VB.Label Label91 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Net Amount      : "
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   14040
         TabIndex        =   51
         Top             =   8640
         Width           =   1365
      End
      Begin VB.Label Label94 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "No. of product :"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001468F3&
         Height          =   255
         Left            =   60
         TabIndex        =   50
         Top             =   7920
         Width           =   1290
      End
      Begin VB.Label Label95 
         BackColor       =   &H8000000E&
         Caption         =   "Payment By     :"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   495
         Left            =   120
         TabIndex        =   49
         Top             =   8280
         Width           =   1695
      End
      Begin VB.Label Label96 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Amt paid in advance : "
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   13605
         TabIndex        =   48
         Top             =   8400
         Width           =   1845
      End
      Begin VB.Label Label97 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "With Tax Amount  :"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   8880
         TabIndex        =   47
         Top             =   8400
         Width           =   1620
      End
      Begin VB.Label Label98 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "Final Amount :"
         ForeColor       =   &H000080FF&
         Height          =   675
         Left            =   16320
         TabIndex        =   46
         Top             =   2400
         Width           =   1035
      End
      Begin VB.Label Label86 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "Purchase rate:"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   7320
         TabIndex        =   45
         Top             =   4320
         Width           =   945
      End
      Begin VB.Label Label87 
         BackColor       =   &H8000000E&
         Caption         =   "Challan No.             :"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   375
         Left            =   10800
         TabIndex        =   44
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label93 
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
         Left            =   120
         TabIndex        =   43
         Top             =   7560
         Width           =   1275
      End
      Begin VB.Line Line3 
         X1              =   1545
         X2              =   20025
         Y1              =   7680
         Y2              =   7680
      End
      Begin VB.Label Label108 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Order Date            :"
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
         Left            =   6000
         TabIndex        =   42
         Top             =   240
         Width           =   1965
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
         Index           =   7
         Left            =   12600
         TabIndex        =   41
         Top             =   480
         Width           =   135
      End
      Begin VB.Label Label110 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "PRODUCT ORDERED"
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
         Left            =   120
         TabIndex        =   40
         Top             =   1800
         Width           =   1845
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00000080&
         X1              =   1200
         X2              =   20040
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label Label5 
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
         Height          =   300
         Index           =   4
         Left            =   10680
         TabIndex        =   39
         Top             =   4320
         Width           =   165
      End
      Begin VB.Label Label90 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "Quantity Ordered:"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   8400
         TabIndex        =   38
         Top             =   4320
         Width           =   1065
      End
      Begin VB.Label Label88 
         BackColor       =   &H8000000E&
         Caption         =   "DD No.             :"
         ForeColor       =   &H8000000C&
         Height          =   375
         Left            =   120
         TabIndex        =   37
         Top             =   8760
         Visible         =   0   'False
         Width           =   2295
      End
   End
   Begin VB.Label Label43 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
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
      Left            =   14040
      TabIndex        =   98
      Top             =   13080
      Width           =   1455
   End
   Begin VB.Label Label42 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
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
      Left            =   13200
      TabIndex        =   97
      Top             =   13440
      Width           =   855
   End
   Begin VB.Label Label41 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
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
      Left            =   14640
      TabIndex        =   96
      Top             =   13440
      Width           =   855
   End
   Begin VB.Label Label40 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
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
      Left            =   14040
      TabIndex        =   95
      Top             =   13440
      Width           =   615
   End
   Begin VB.Label Label37 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
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
      Left            =   15480
      TabIndex        =   94
      Top             =   13080
      Width           =   1335
   End
   Begin VB.Label Label31 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
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
      Left            =   11400
      TabIndex        =   93
      Top             =   13080
      Width           =   1215
   End
   Begin VB.Label Label30 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Quantity Ordered"
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
      Left            =   10200
      TabIndex        =   92
      Top             =   13080
      Width           =   1215
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
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
      Left            =   9000
      TabIndex        =   91
      Top             =   13080
      Width           =   1215
   End
   Begin VB.Label Label28 
      BackColor       =   &H008080FF&
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
      Left            =   1920
      TabIndex        =   90
      Top             =   13080
      Width           =   615
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
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
      Left            =   7080
      TabIndex        =   89
      Top             =   13080
      Width           =   1935
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
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
      Left            =   3720
      TabIndex        =   88
      Top             =   13080
      Width           =   3375
   End
   Begin VB.Label Label25 
      BackColor       =   &H008080FF&
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
      Left            =   2520
      TabIndex        =   87
      Top             =   13080
      Width           =   1215
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
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
      Left            =   12600
      TabIndex        =   86
      Top             =   13080
      Width           =   1455
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
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
      Left            =   12600
      TabIndex        =   85
      Top             =   13440
      Width           =   615
   End
End
Attribute VB_Name = "purchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
Set R = New ADODB.Recordset
SQL = "select *from product where P_id='" + Combo1.Text + "'"
Set R = C.Execute(SQL)
Text1.Text = R.Fields(1)
txt11.Text = R.Fields(2)
'txt13.Text = R.Fields(6)
gst.Text = R.Fields(5)
SQL = "select *from supplierprd where P_id='" + Combo1.Text + "'"
Set R = C.Execute(SQL)
txt13.Text = R.Fields(2)

End Sub

Private Sub Command1_Click()
a = InputBox("Enter the Serial No. you want to remove:", "for delete")
If a = blank Then
MsgBox "Please enter serial no."
Else
List8.RemoveItem (a - 1)
List2.RemoveItem (a - 1)
List4.RemoveItem (a - 1)
List5.RemoveItem (a - 1)
List6.RemoveItem (a - 1)
List7.RemoveItem (a - 1)
List12.RemoveItem (a - 1)
List10.RemoveItem (a - 1)
List3.RemoveItem (a - 1)
List11.RemoveItem (a - 1)
List9.RemoveItem (a - 1)
List1.Clear
For i = 1 To List8.ListCount
    List1.AddItem i
Next i
Dim l As Long
Dim lSum As Long
For l = 0 To List9.ListCount - 1
    lSum = lSum + CLng(List9.List(l))
Next
Text4.Text = lSum

End If

End Sub

Private Sub Form_Load()
CONN
Set R = New ADODB.Recordset
SQL = "select *from purordetail where postatus='Incomplete' or postatus='Partially Received' "
Set R = C.Execute(SQL)
While R.EOF = False
odno.AddItem R.Fields(0)
R.MoveNext
Wend
MonthView1.Visible = False
MonthView2.Visible = False
gst.Visible = False
MonthView1.Refresh
MonthView2.Refresh



End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
TXT6.Text = Format(MonthView1, "dd-mmm-yyyy")
MonthView1.Visible = False
If TXT6.Text < txt3.Text Then
MsgBox "invalid date"
TXT6.Text = " "
End If
End Sub

Private Sub MonthView2_DateClick(ByVal DateClicked As Date)
Text3.Text = Format(MonthView2, "dd-mmm-yyyy")
MonthView2.Visible = False
If Text3.Text < txt3.Text Or Text3.Text > TXT6.Text Then
MsgBox "invalid date "
Text3.Text = " "
End If
End Sub


Private Sub odno_Click()
Set R = New ADODB.Recordset
SQL = " select *from purordetail where pur_orderno='" + odno.Text + "'"
Set R = C.Execute(SQL)
txt3.Text = R.Fields(1)
txt1.Text = R.Fields("sup_id")
Txt20.Text = R.Fields(5)
txt22.Text = R.Fields(8)
final.Text = R.Fields(7)
Set R = New ADODB.Recordset
SQL = "select *from supplier where sup_id='" + txt1.Text + "'"
Set R = C.Execute(SQL)
txt2.Text = R.Fields(1)
txt4.Text = R.Fields(2)
txt5.Text = R.Fields(7)
TXT8.Text = R.Fields(3)
Set R = New ADODB.Recordset
SQL = "select *from p_det where pur_orderno='" + odno.Text + "'"
Set R = C.Execute(SQL)
sr.Clear
pid.Clear
pnm.Clear
ptyp.Clear
prate.Clear
qty1.Clear
prc.Clear
cgst.Clear
camt.Clear
sgst.Clear
samt.Clear
totprc.Clear
Combo1.Clear
While R.EOF = False
sr.AddItem R.Fields(1)
pid.AddItem R.Fields(2)
pnm.AddItem R.Fields(3)
ptyp.AddItem R.Fields(4)
prate.AddItem R.Fields(5)
qty1.AddItem R.Fields(7)
prc.AddItem R.Fields(8)
cgst.AddItem R.Fields(9)
camt.AddItem R.Fields(10)
sgst.AddItem R.Fields(11)
samt.AddItem R.Fields(12)
totprc.AddItem R.Fields(13)
Combo1.AddItem R.Fields(2)
R.MoveNext
Wend

End Sub

Private Sub ok_Click()
List8.AddItem Combo1.Text
List2.AddItem Text1.Text
List4.AddItem txt11.Text
List5.AddItem txt13.Text
List6.AddItem TXT15.Text
List7.AddItem Val(txt13.Text) * Val(TXT15.Text)
List12.AddItem Val(gst.Text) / 2
List3.AddItem Val(gst.Text) / 2
List10.AddItem (Val(Val(TXT15.Text) * Val(txt13.Text)) * Val(gst.Text) / 100) / 2
List11.AddItem (Val(Val(TXT15.Text) * Val(txt13.Text)) * Val(gst.Text) / 100) / 2
List9.AddItem (Val(Val(TXT15.Text) * Val(txt13.Text)) * Val(gst.Text) / 100) + (Val(txt13.Text) * Val(TXT15.Text))




Dim l As Long
Dim lSum As Long
For l = 0 To List9.ListCount - 1
    lSum = lSum + CLng(List9.List(l))
Next
Text4.Text = lSum

If List1.ListCount = 0 Then
List1.AddItem 1
Else
List1.AddItem (List1.ListCount + 1)
End If
TXT19.Text = List1.ListCount
TXT21.Text = Text4.Text
txt23.Text = Val(TXT21.Text) - Val(txt22.Text)
Text1.Text = " "
txt11.Text = " "
txt13.Text = " "
txt14.Text = " "
TXT15.Text = " "
Combo1.Text = " "
End Sub



Private Sub save_Click()
If odno.Text = blank Or TXT6.Text = blank Or txt7.Text = blank Or Text2.Text = blank Or Text3.Text = blank Or TXT19.Text = blank Or Txt20.Text = blank Or TXT21.Text = blank Or Text6.Text = blank Then
MsgBox "Please fill the details first !!"
Else
Set R = New ADODB.Recordset
SQL = "insert into ordetails values ('" + Text2.Text + "','" + Text3.Text + "','" + odno.Text + "','" + txt7.Text + "'," + TXT19.Text + ",'" + Text5.Text + "','" + Txt20.Text + "'," + TXT21.Text + "," + txt22.Text + "," + txt23.Text + ")"
Set R = C.Execute(SQL)
MsgBox "Record saved!"

Dim i As Long
For i = 0 To List1.ListCount - 1
SQL = "insert into recvd_p_det values('" + odno.Text + "'," + List1.List(i) + ",'" + List8.List(i) + "','" + List2.List(i) + "','" + List4.List(i) + "'," + List5.List(i) + "," + List6.List(i) + "," + List7.List(i) + "," + List12.List(i) + "," + List10.List(i) + "," + List3.List(i) + "," + List11.List(i) + "," + List9.List(i) + ")"
Set R = C.Execute(SQL)
Next i
MsgBox "data saved"

SQL = "update purordetail set postatus = '" + Text6.Text + "' where pur_orderno='" + odno.Text + "' "
Set R = C.Execute(SQL)

Dim k As Long
For k = 0 To List1.ListCount - 1
SQL = "update stock set avl_qty= avl_qty+" + List6.List(k) + " where p_id='" + List8.List(k) + "'"
Set R = C.Execute(SQL)
Next k
MsgBox "stock updated.....!"
End If
Unload Me
purchase.Show

End Sub

Private Sub ST01_Click()
Text6.Text = ST01.Caption
End Sub

Private Sub ST02_Click()
Text6.Text = ST02.Caption
End Sub

Private Sub ST03_Click()
Text6.Text = ST03.Caption
End Sub

Private Sub ST121_Click()
Text5.Text = ST121.Caption

End Sub

Private Sub ST122_Click()
Text5.Text = ST122.Caption
Txt20.Text = ST122.Caption
End Sub

Private Sub Text3_Click()
MonthView2.Visible = True
End Sub

Private Sub txt14_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then TXT15.SetFocus
End Sub

Private Sub txt15_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then ok.SetFocus
End Sub

Private Sub TXT6_Click()
MonthView1.Visible = True
End Sub
