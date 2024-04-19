VERSION 5.00
Begin VB.Form sale_return 
   Caption         =   "Form2"
   ClientHeight    =   9540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15825
   LinkTopic       =   "Form2"
   ScaleHeight     =   9540
   ScaleWidth      =   15825
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   9885
      Left            =   120
      ScaleHeight     =   9825
      ScaleWidth      =   20025
      TabIndex        =   0
      Top             =   -600
      Width           =   20085
      Begin VB.Frame Frame4 
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1.02135e5
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   22335
         Begin VB.OptionButton O3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "Defect"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   18600
            TabIndex        =   48
            Top             =   4080
            Width           =   1380
         End
         Begin VB.TextBox Txt7 
            BackColor       =   &H00FFFF00&
            DataField       =   "STK_ID"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   13920
            Locked          =   -1  'True
            TabIndex        =   47
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1080
            Width           =   4815
         End
         Begin VB.TextBox Txt6 
            BackColor       =   &H00FFFF00&
            DataField       =   "STK_ID"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   13920
            Locked          =   -1  'True
            TabIndex        =   46
            TabStop         =   0   'False
            Text            =   " "
            Top             =   600
            Width           =   2655
         End
         Begin VB.TextBox Txt4 
            BackColor       =   &H00FFFF00&
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   8400
            Locked          =   -1  'True
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   960
            Width           =   2175
         End
         Begin VB.TextBox Txt3 
            BackColor       =   &H00FFFF00&
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   8400
            Locked          =   -1  'True
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   480
            Width           =   2175
         End
         Begin VB.TextBox Txt1 
            BackColor       =   &H00FFFF00&
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   1
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   480
            Width           =   2175
         End
         Begin VB.TextBox Txt5 
            BackColor       =   &H00FFFF80&
            DataField       =   "STK_ID"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   13920
            Locked          =   -1  'True
            TabIndex        =   42
            TabStop         =   0   'False
            Text            =   " "
            Top             =   120
            Width           =   2415
         End
         Begin VB.TextBox TXT20 
            DataField       =   "M_ID"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   9720
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   7680
            Width           =   1335
         End
         Begin VB.ListBox net 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1950
            ItemData        =   "sale return.frx":0000
            Left            =   15240
            List            =   "sale return.frx":0002
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   5640
            Width           =   1335
         End
         Begin VB.ListBox smt 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1950
            ItemData        =   "sale return.frx":0004
            Left            =   14280
            List            =   "sale return.frx":0006
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   5640
            Width           =   975
         End
         Begin VB.ListBox cgst 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1950
            ItemData        =   "sale return.frx":0008
            Left            =   11760
            List            =   "sale return.frx":000A
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   5640
            Width           =   615
         End
         Begin VB.ListBox cmt 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1950
            ItemData        =   "sale return.frx":000C
            Left            =   12360
            List            =   "sale return.frx":000E
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   5640
            Width           =   1215
         End
         Begin VB.ListBox sgst 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1950
            ItemData        =   "sale return.frx":0010
            Left            =   13560
            List            =   "sale return.frx":0012
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   5640
            Width           =   735
         End
         Begin VB.ListBox nm 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1950
            ItemData        =   "sale return.frx":0014
            Left            =   1920
            List            =   "sale return.frx":0016
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   5640
            Width           =   3135
         End
         Begin VB.ListBox bat 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1950
            ItemData        =   "sale return.frx":0018
            Left            =   5040
            List            =   "sale return.frx":001A
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   5640
            Width           =   1575
         End
         Begin VB.ListBox id 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1950
            ItemData        =   "sale return.frx":001C
            Left            =   720
            List            =   "sale return.frx":001E
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   5640
            Width           =   1215
         End
         Begin VB.ListBox sr 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1950
            ItemData        =   "sale return.frx":0020
            Left            =   120
            List            =   "sale return.frx":0022
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   5640
            Width           =   615
         End
         Begin VB.ComboBox TXT2 
            BackColor       =   &H00FFFF00&
            CausesValidation=   0   'False
            DataField       =   "MED_NM"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2400
            TabIndex        =   31
            Top             =   960
            Width           =   3495
         End
         Begin VB.ListBox unt 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1950
            ItemData        =   "sale return.frx":0024
            Left            =   6600
            List            =   "sale return.frx":0026
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   5640
            Width           =   1215
         End
         Begin VB.TextBox Text22 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000E&
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
            Height          =   495
            Left            =   17880
            TabIndex        =   29
            TabStop         =   0   'False
            Text            =   "0.00"
            Top             =   6120
            Width           =   1695
         End
         Begin VB.TextBox Text21 
            BackColor       =   &H8000000B&
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   16560
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   7680
            Width           =   1455
         End
         Begin VB.TextBox Text19 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   1920
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   7680
            Width           =   1575
         End
         Begin VB.CommandButton cancel 
            BackColor       =   &H80000016&
            Caption         =   "Cancel"
            Height          =   495
            Left            =   18120
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   6960
            Width           =   1575
         End
         Begin VB.CommandButton save 
            BackColor       =   &H8000000B&
            Caption         =   "Save"
            DisabledPicture =   "sale return.frx":0028
            DownPicture     =   "sale return.frx":0B26
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
            Left            =   18120
            MaskColor       =   &H00C0FFC0&
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   7560
            Width           =   1575
         End
         Begin VB.TextBox TXT8 
            DataField       =   "M_ID"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   24
            TabStop         =   0   'False
            Text            =   " "
            Top             =   4080
            Width           =   1575
         End
         Begin VB.TextBox TXT9 
            DataField       =   "MED_NM"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   23
            TabStop         =   0   'False
            Text            =   " "
            Top             =   4080
            Width           =   2895
         End
         Begin VB.TextBox TXT10 
            DataField       =   "BATCH"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   4560
            Locked          =   -1  'True
            TabIndex        =   22
            TabStop         =   0   'False
            Text            =   " "
            Top             =   4080
            Width           =   1695
         End
         Begin VB.TextBox TXT11 
            DataField       =   "UOM"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   6240
            Locked          =   -1  'True
            TabIndex        =   21
            TabStop         =   0   'False
            Text            =   " "
            Top             =   4080
            Width           =   1335
         End
         Begin VB.TextBox TXT18 
            BackColor       =   &H8000000B&
            DataField       =   "STK_ID"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   14640
            Locked          =   -1  'True
            TabIndex        =   20
            TabStop         =   0   'False
            Text            =   " "
            Top             =   4080
            Width           =   1335
         End
         Begin VB.TextBox TXT17 
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   13800
            Locked          =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            Text            =   " "
            Top             =   4080
            Width           =   735
         End
         Begin VB.TextBox TXT16 
            BackColor       =   &H8000000B&
            DataField       =   "STK_ID"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   12600
            Locked          =   -1  'True
            TabIndex        =   18
            TabStop         =   0   'False
            Text            =   " "
            Top             =   4080
            Width           =   1215
         End
         Begin VB.TextBox TXT13 
            DataField       =   "SO_QTY"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   8640
            Locked          =   -1  'True
            TabIndex        =   17
            TabStop         =   0   'False
            Text            =   " "
            Top             =   4080
            Width           =   1215
         End
         Begin VB.TextBox TXT12 
            DataField       =   "SP"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   7560
            Locked          =   -1  'True
            TabIndex        =   16
            TabStop         =   0   'False
            Text            =   " "
            Top             =   4080
            Width           =   1095
         End
         Begin VB.OptionButton O2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "Expired"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   17280
            TabIndex        =   15
            Top             =   4080
            Width           =   1140
         End
         Begin VB.OptionButton O1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "Excess"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   16080
            TabIndex        =   14
            Top             =   4080
            Width           =   1020
         End
         Begin VB.ListBox qty 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1950
            ItemData        =   "sale return.frx":1D41
            Left            =   9000
            List            =   "sale return.frx":1D43
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   5640
            Width           =   1455
         End
         Begin VB.ListBox sp 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1950
            ItemData        =   "sale return.frx":1D45
            Left            =   7800
            List            =   "sale return.frx":1D47
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   5640
            Width           =   1215
         End
         Begin VB.ListBox prc 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1950
            ItemData        =   "sale return.frx":1D49
            Left            =   10440
            List            =   "sale return.frx":1D4B
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   5640
            Width           =   1335
         End
         Begin VB.ListBox STAT 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1950
            ItemData        =   "sale return.frx":1D4D
            Left            =   16560
            List            =   "sale return.frx":1D4F
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   5640
            Width           =   1335
         End
         Begin VB.CommandButton add 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Add"
            CausesValidation=   0   'False
            Enabled         =   0   'False
            Height          =   495
            Left            =   18600
            Style           =   1  'Graphical
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   4440
            Width           =   1215
         End
         Begin VB.TextBox Text15 
            DataField       =   "SP"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   1
            Left            =   11640
            Locked          =   -1  'True
            TabIndex        =   8
            TabStop         =   0   'False
            Text            =   " "
            Top             =   4080
            Width           =   855
         End
         Begin VB.TextBox Text7 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000E&
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
            ForeColor       =   &H00000000&
            Height          =   495
            Left            =   17520
            TabIndex        =   7
            TabStop         =   0   'False
            Text            =   "0.00"
            Top             =   2400
            Width           =   1695
         End
         Begin VB.TextBox TXT14 
            BackColor       =   &H8000000B&
            DataField       =   "STK_ID"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   10320
            TabIndex        =   6
            Text            =   " "
            Top             =   4080
            Width           =   1215
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "Invoice date             :"
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
            Left            =   6210
            TabIndex        =   93
            Top             =   960
            Width           =   2145
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "Return Date             :"
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
            Left            =   6240
            TabIndex        =   92
            Top             =   480
            Width           =   2130
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "Invoice No.           :"
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
            Left            =   360
            TabIndex        =   91
            Top             =   960
            Width           =   1965
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "Return Id               :"
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
            Left            =   360
            TabIndex        =   90
            Top             =   480
            Width           =   1995
         End
         Begin VB.Label Label39 
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
            Left            =   15240
            TabIndex        =   89
            Top             =   5040
            Width           =   1335
         End
         Begin VB.Label Label34 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
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
            Height          =   435
            Left            =   13500
            TabIndex        =   88
            Top             =   5400
            Width           =   855
         End
         Begin VB.Label Label31 
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
            Left            =   11760
            TabIndex        =   87
            Top             =   5400
            Width           =   615
         End
         Begin VB.Label Label30 
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
            Index           =   0
            Left            =   11760
            TabIndex        =   86
            Top             =   5040
            Width           =   1815
         End
         Begin VB.Label Label35 
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
            Height          =   375
            Index           =   0
            Left            =   14280
            TabIndex        =   85
            Top             =   5400
            Width           =   975
         End
         Begin VB.Label Label32 
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
            Height          =   375
            Index           =   0
            Left            =   12360
            TabIndex        =   84
            Top             =   5400
            Width           =   1215
         End
         Begin VB.Label Label33 
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
            Height          =   495
            Left            =   13560
            TabIndex        =   83
            Top             =   5040
            Width           =   1695
         End
         Begin VB.Label Label23 
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
            TabIndex        =   82
            Top             =   5040
            Width           =   1215
         End
         Begin VB.Label Label24 
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
            Left            =   1920
            TabIndex        =   81
            Top             =   5040
            Width           =   3135
         End
         Begin VB.Label Label25 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Batch No."
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
            Left            =   5040
            TabIndex        =   80
            Top             =   5040
            Width           =   1575
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
            Left            =   360
            TabIndex        =   79
            Top             =   0
            Width           =   2085
         End
         Begin VB.Label Label20 
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
            Height          =   285
            Index           =   0
            Left            =   15960
            TabIndex        =   78
            Top             =   3600
            Width           =   285
         End
         Begin VB.Label Label8 
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
            Left            =   11640
            TabIndex        =   77
            Top             =   1080
            Width           =   2265
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "Customer Mob No. :"
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
            Left            =   11640
            TabIndex        =   76
            Top             =   120
            Width           =   2115
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "Customer Name     :"
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
            Left            =   11640
            TabIndex        =   75
            Top             =   600
            Width           =   2085
         End
         Begin VB.Label Label26 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
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
            Height          =   690
            Left            =   6555
            TabIndex        =   74
            Top             =   5040
            Width           =   1305
         End
         Begin VB.Label Label41 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            Caption         =   "Final Amount: "
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   645
            Left            =   18120
            TabIndex        =   73
            Top             =   5400
            Width           =   1260
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
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
            ForeColor       =   &H00008000&
            Height          =   315
            Left            =   14520
            TabIndex        =   72
            Top             =   7680
            Width           =   1965
         End
         Begin VB.Label Label43 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
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
            ForeColor       =   &H00008000&
            Height          =   315
            Left            =   8880
            TabIndex        =   71
            Top             =   7680
            Width           =   600
         End
         Begin VB.Label Label42 
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
            ForeColor       =   &H001468F3&
            Height          =   315
            Left            =   240
            TabIndex        =   70
            Top             =   7680
            Width           =   1590
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
            Index           =   0
            Left            =   120
            TabIndex        =   69
            Top             =   5040
            Width           =   615
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
            Height          =   405
            Index           =   9
            Left            =   1440
            TabIndex        =   68
            Top             =   360
            Width           =   165
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
            Height          =   405
            Index           =   1
            Left            =   1560
            TabIndex        =   67
            Top             =   840
            Width           =   165
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
            Height          =   405
            Index           =   2
            Left            =   7560
            TabIndex        =   66
            Top             =   360
            Width           =   165
         End
         Begin VB.Line Line1 
            X1              =   0
            X2              =   20160
            Y1              =   1560
            Y2              =   1560
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "Product Id"
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
            Left            =   375
            TabIndex        =   65
            Top             =   3720
            Width           =   1110
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "Product Name"
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
            Left            =   2040
            TabIndex        =   64
            Top             =   3720
            Width           =   1500
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "Batch No :"
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
            Left            =   4680
            TabIndex        =   63
            Top             =   3720
            Width           =   1095
         End
         Begin VB.Label Label12 
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
            Left            =   6480
            TabIndex        =   62
            Top             =   3720
            Width           =   540
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            Caption         =   "MRP"
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
            Left            =   7560
            TabIndex        =   61
            Top             =   3720
            Width           =   1035
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            Caption         =   "Qty Ordered"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   8760
            TabIndex        =   60
            Top             =   3580
            Width           =   855
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            Caption         =   "Qty Returned"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   10320
            TabIndex        =   59
            Top             =   3720
            Width           =   1215
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            Caption         =   "Price:"
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
            Left            =   12720
            TabIndex        =   58
            Top             =   3720
            Width           =   855
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            Caption         =   "Gst(%)"
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
            Left            =   13800
            TabIndex        =   57
            Top             =   3720
            Width           =   855
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            Caption         =   "Total Price"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   14760
            TabIndex        =   56
            Top             =   3720
            Width           =   1215
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "REASON"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   1
            Left            =   17400
            TabIndex        =   55
            Top             =   3720
            Width           =   885
         End
         Begin VB.Label Label27 
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
            Left            =   7800
            TabIndex        =   54
            Top             =   5040
            Width           =   1215
         End
         Begin VB.Label Label40 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Reason"
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
            Left            =   16560
            TabIndex        =   53
            Top             =   5040
            Width           =   1335
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "Final amount: "
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
            Index           =   0
            Left            =   17760
            TabIndex        =   52
            Top             =   1920
            Width           =   1545
         End
         Begin VB.Label Label16 
            BackColor       =   &H8000000E&
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
            Height          =   555
            Left            =   11880
            TabIndex        =   51
            Top             =   3585
            Width           =   705
         End
         Begin VB.Line Line4 
            X1              =   -120
            X2              =   19800
            Y1              =   4920
            Y2              =   4920
         End
         Begin VB.Label Label28 
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
            Left            =   9000
            TabIndex        =   50
            Top             =   5040
            Width           =   1455
         End
         Begin VB.Label Label29 
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
            Left            =   10440
            TabIndex        =   49
            Top             =   5040
            Width           =   1335
         End
         Begin VB.Line Line3 
            BorderStyle     =   2  'Dash
            X1              =   11160
            X2              =   11160
            Y1              =   1560
            Y2              =   0
         End
      End
      Begin VB.CommandButton Command4 
         Height          =   495
         Left            =   -74280
         Picture         =   "sale return.frx":1D51
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   4920
         Width           =   1695
      End
      Begin VB.TextBox Text9 
         DataField       =   "SORD_NO"
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
         Left            =   -72960
         TabIndex        =   3
         Top             =   3120
         Width           =   2535
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Show All"
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
         Left            =   -72240
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   4920
         Width           =   1695
      End
      Begin VB.CommandButton CLOSE 
         BackColor       =   &H00FBE2BD&
         Caption         =   "CLOSE"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   -61680
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   6615
      End
   End
End
Attribute VB_Name = "sale_return"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
