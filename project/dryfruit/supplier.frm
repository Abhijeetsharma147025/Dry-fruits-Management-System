VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form supplier 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Supplier"
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15990
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   15990
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox tx10 
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3000
      MaxLength       =   15
      TabIndex        =   36
      Text            =   " "
      Top             =   4800
      Width           =   2055
   End
   Begin VB.CommandButton Close 
      BackColor       =   &H00C0C000&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7920
      MaskColor       =   &H00404000&
      Style           =   1  'Graphical
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton update 
      BackColor       =   &H00C0C000&
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      MaskColor       =   &H00404000&
      Style           =   1  'Graphical
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton save 
      BackColor       =   &H00C0C000&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      MaskColor       =   &H00404000&
      Style           =   1  'Graphical
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   7680
      Width           =   1455
   End
   Begin VB.TextBox tx2 
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3000
      MaxLength       =   20
      TabIndex        =   0
      Text            =   " "
      Top             =   1200
      Width           =   4455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   9855
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   16095
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   12240
         Top             =   6000
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=MSDAORA.1;User ID=aniket/arpit;Persist Security Info=False"
         OLEDBString     =   "Provider=MSDAORA.1;User ID=aniket/arpit;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from supplier"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "supplier.frx":0000
         Height          =   1215
         Left            =   0
         TabIndex        =   53
         Top             =   6240
         Width           =   15855
         _ExtentX        =   27966
         _ExtentY        =   2143
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   11
         BeginProperty Column00 
            DataField       =   "SUP_ID"
            Caption         =   "ID"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "SUP_NM"
            Caption         =   "Name"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "SUP_MOB"
            Caption         =   "SUP_MOB"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "SUP_LOCATION"
            Caption         =   "Address"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "SUP_STATE"
            Caption         =   "State"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "SUP_CITY"
            Caption         =   "City"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "SUP_PINCODE"
            Caption         =   "Pincode"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "COM"
            Caption         =   "Company Name"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "SUP_EMAIL"
            Caption         =   "Email"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "SUP_GSTNO"
            Caption         =   "GST No."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "STATUS"
            Caption         =   "Supplier Status"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1019.906
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1349.858
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1620.284
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1154.835
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1365.165
            EndProperty
            BeginProperty Column10 
            EndProperty
         EndProperty
      End
      Begin VB.TextBox Text2 
         Height          =   525
         Left            =   6120
         TabIndex        =   51
         Top             =   120
         Width           =   1815
      End
      Begin VB.Frame frame3 
         BackColor       =   &H8000000E&
         Caption         =   "Product Sold By Supplier"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Left            =   11040
         TabIndex        =   38
         Top             =   1560
         Width           =   4575
         Begin VB.ListBox rate 
            BackColor       =   &H00FFFFC0&
            Height          =   2010
            ItemData        =   "supplier.frx":0015
            Left            =   2040
            List            =   "supplier.frx":0017
            TabIndex        =   47
            Top             =   1800
            Width           =   1095
         End
         Begin VB.ListBox id 
            BackColor       =   &H00FFFFC0&
            Height          =   2010
            ItemData        =   "supplier.frx":0019
            Left            =   840
            List            =   "supplier.frx":001B
            TabIndex        =   44
            Top             =   1800
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   1920
            TabIndex        =   43
            Top             =   960
            Width           =   1335
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00C0C000&
            Caption         =   "Delete"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3360
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   2400
            Width           =   975
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00C0C000&
            Caption         =   "Add"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3360
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   1680
            Width           =   975
         End
         Begin VB.ComboBox Combo2 
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   240
            TabIndex        =   40
            Top             =   960
            Width           =   1575
         End
         Begin VB.ListBox sr 
            BackColor       =   &H00FFFFC0&
            Height          =   2010
            ItemData        =   "supplier.frx":001D
            Left            =   240
            List            =   "supplier.frx":001F
            TabIndex        =   39
            Top             =   1800
            Width           =   615
         End
         Begin VB.Label Label25 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Rate:"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1920
            TabIndex        =   50
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label24 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Product Id:"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   49
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label23 
            BackColor       =   &H00FFFF00&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Rate"
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
            Left            =   2040
            TabIndex        =   48
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label Label22 
            BackColor       =   &H00FFFF00&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Product Id"
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
            Left            =   840
            TabIndex        =   46
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label18 
            BackColor       =   &H00FFFF00&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "S.no"
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
            Left            =   240
            TabIndex        =   45
            Top             =   1440
            Width           =   615
         End
      End
      Begin VB.Frame searchby 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   6120
         TabIndex        =   32
         Top             =   4560
         Width           =   3975
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   600
            TabIndex        =   33
            Text            =   "Sup_Id"
            Top             =   840
            Width           =   2415
         End
         Begin VB.Image close2 
            BorderStyle     =   1  'Fixed Single
            Height          =   510
            Left            =   3480
            Picture         =   "supplier.frx":0021
            Top             =   120
            Width           =   510
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Search by"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   600
            TabIndex        =   34
            Top             =   840
            Width           =   1935
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   1095
         Left            =   0
         TabIndex        =   26
         Top             =   7440
         Width           =   15975
         Begin VB.CommandButton delete 
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
            Height          =   615
            Left            =   6360
            MaskColor       =   &H00404000&
            Style           =   1  'Graphical
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton search 
            BackColor       =   &H00C0C000&
            Caption         =   "Search"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   4800
            MaskColor       =   &H00404000&
            Style           =   1  'Graphical
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton add 
            BackColor       =   &H00C0C000&
            Caption         =   "Add New"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            MaskColor       =   &H00404000&
            Style           =   1  'Graphical
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   240
            Width           =   1455
         End
         Begin VB.Line Line2 
            X1              =   -120
            X2              =   15960
            Y1              =   0
            Y2              =   0
         End
      End
      Begin VB.TextBox tx8 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3000
         MaxLength       =   30
         TabIndex        =   18
         Top             =   3600
         Width           =   4815
      End
      Begin VB.TextBox tx3 
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3000
         MaxLength       =   11
         TabIndex        =   17
         Text            =   " "
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox tx9 
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3000
         MaxLength       =   30
         TabIndex        =   8
         Top             =   4200
         Width           =   3615
      End
      Begin VB.TextBox tx4 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3000
         MaxLength       =   50
         TabIndex        =   7
         Top             =   2160
         Width           =   4455
      End
      Begin VB.TextBox tx5 
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   6
         Text            =   " "
         Top             =   2640
         Width           =   2055
      End
      Begin VB.TextBox tx6 
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   6360
         MaxLength       =   10
         TabIndex        =   5
         Text            =   " "
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox tx7 
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3000
         MaxLength       =   7
         TabIndex        =   4
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label27 
         Caption         =   "active"
         Height          =   255
         Left            =   8400
         TabIndex        =   52
         Top             =   3960
         Width           =   495
      End
      Begin VB.Label Label21 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11520
         TabIndex        =   37
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Email Id                      :"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   35
         Top             =   4200
         Width           =   2325
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1680
         TabIndex        =   25
         Top             =   4680
         Width           =   135
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFFF&
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
         Left            =   2280
         TabIndex        =   24
         Top             =   3600
         Width           =   135
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
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
         Left            =   2280
         TabIndex        =   23
         Top             =   2040
         Width           =   135
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1560
         TabIndex        =   21
         Top             =   1680
         Width           =   135
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
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
         Left            =   2160
         TabIndex        =   20
         Top             =   1200
         Width           =   135
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFFC0&
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
         Left            =   11280
         TabIndex        =   19
         Top             =   120
         Width           =   135
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "G.S.T No.                    :"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   16
         Top             =   4800
         Width           =   2340
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Supplier Location     :"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   15
         Top             =   2280
         Width           =   2355
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Company Name        :"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   14
         Top             =   3720
         Width           =   2355
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Mobile No.                 :"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   13
         Top             =   1800
         Width           =   2325
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Supplier Name          :"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   12
         Top             =   1320
         Width           =   2355
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "State                            :"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   11
         Top             =   2760
         Width           =   2310
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "City        :"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   5280
         TabIndex        =   10
         Top             =   2640
         Width           =   1020
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pin no .                        :"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   9
         Top             =   3240
         Width           =   2355
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   15960
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Supplier Id:"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Left            =   9840
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "  Supplier :"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   855
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   16095
      End
      Begin VB.Image Image1 
         Height          =   1620
         Left            =   7680
         OLEDragMode     =   1  'Automatic
         Picture         =   "supplier.frx":0563
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   3180
      End
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
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
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   135
   End
End
Attribute VB_Name = "supplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub add_Click()
Dim a As String
Set R = New ADODB.Recordset
SQL = "select max(to_number(SUBSTR(SUP_ID,4,LENGTH(SUP_ID))))from supplier"
Set R = C.Execute(SQL)
If IsNull(R.Fields(0)) Then
Label21.Caption = "S" & "00" & 1
Else
Label21.Caption = "S" & "00" & R.Fields(0) + 1
a = Label21.Caption
End If
If (a = "S0010") Then
Set R = New ADODB.Recordset
SQL = "select max(to_number(SUBSTR(SUP_ID,3,LENGTH(SUP_ID))))from supplier"
Set R = C.Execute(SQL)
Label21.Caption = "S" & "0" & R.Fields(0) + 1
End If
tx2.SetFocus
tx2.Text = " "
tx3.Text = " "
tx4.Text = " "
tx5.Text = " "
tx6.Text = " "
tx7.Text = " "
tx8.Text = " "
tx9.Text = " "
tx10.Text = " "
End Sub

Private Sub Close_Click()
Unload Me
End Sub

Private Sub close2_Click()
searchby.Visible = False
End Sub

Private Sub Combo1_Click()
Set R = New ADODB.Recordset
SQL = "select *from supplier where SUP_ID='" + Combo1.Text + "'"
Set R = C.Execute(SQL)
Label21.Caption = R.Fields(0)
tx2.Text = R.Fields(1)
tx3.Text = R.Fields(2)
tx4.Text = R.Fields(3)
tx5.Text = R.Fields(4)
tx6.Text = R.Fields(5)
tx7.Text = R.Fields(6)
tx8.Text = R.Fields(7)
tx9.Text = R.Fields(8)
tx10.Text = R.Fields(9)
End Sub

Private Sub Command1_Click()
If sr.ListCount = 0 Then
sr.AddItem 1
Else
sr.AddItem (sr.ListCount + 1)
End If
id.AddItem Combo2.Text
rate.AddItem Text1.Text

Combo2.Text = " "
Text1.Text = " "
End Sub

Private Sub Command2_Click()
a = InputBox("Enter the Serial No. you want to remove:", "for delete")
If a = blank Then
MsgBox "Please enter serial no."
Else
id.RemoveItem (a - 1)
rate.RemoveItem (a - 1)
sr.Clear
For i = 1 To id.ListCount
    sr.AddItem i
Next i
End If

End Sub

Private Sub delete_Click()
Set R = New ADODB.Recordset
SQL = "update supplier set status='Inactive' where sup_id='" + Combo1.Text + "'"


'SQL = "delete from supplier where sup_id='" + Combo1.Text + "'"
Set R = C.Execute(SQL)
MsgBox "record deleted"
Adodc1.Refresh
Combo1.Clear
auto_sup_id
Label21.Caption = " "
tx2.Text = " "
tx3.Text = " "
tx4.Text = " "
tx5.Text = " "
tx6.Text = " "
tx7.Text = " "
tx8.Text = " "
tx9.Text = " "
tx10.Text = " "
End Sub

Private Sub Form_Load()
CONN
Adodc1.Visible = False
searchby.Visible = False
auto_sup_id
auto_prd_id
Adodc1.Refresh
Label27.Visible = False
Text2.Visible = False
End Sub

Private Sub save_Click()

If Label21.Caption = " " Or tx2.Text = " " Or tx3.Text = " " Or tx4.Text = " " Or tx5.Text = " " Or tx6.Text = " " Or tx7.Text = " " Or tx8.Text = " " Or tx9.Text = " " Or tx10.Text = " " Then
MsgBox "Please click on ADD NEW button first and fill details!!"
Else
Set R = New ADODB.Recordset
SQL = "select *from Supplier"
Set R = C.Execute(SQL)

While R.EOF = False
    If (R.Fields(0) = Label21.Caption) Then
        Text2.Text = Label21.Caption
    End If
    R.MoveNext
Wend
   
    If (Text2.Text = Label21.Caption) Then
                  Set R = New ADODB.Recordset
          SQL = "insert into supplierprd values('" + Label21.Caption + "','" + id.List(i) + "'," + rate.List(i) + ")"
                Set R = C.Execute(SQL)
        MsgBox "record saved"

    Else
                
        Set R = New ADODB.Recordset
        SQL = "insert into supplier values('" + Label21.Caption + "','" + tx2.Text + "'," + tx3.Text + ",'" + tx4.Text + "','" + tx5.Text + "','" + tx6.Text + "'," + tx7.Text + ",'" + tx8.Text + "','" + tx9.Text + "','" + tx10.Text + "','" + Label27.Caption + "')"
        Set R = C.Execute(SQL)


        MsgBox "record saved"

        Adodc1.Refresh
        Combo1.Clear
        auto_sup_id
        Dim p As Long
        For p = 0 To sr.ListCount - 1
        SQL = "insert into supplierprd values('" + Label21.Caption + "','" + id.List(p) + "'," + rate.List(p) + ")"
        
        Set R = C.Execute(SQL)
        Next p
        MsgBox "data saved"
    End If

End If


Label21.Caption = " "
tx2.Text = " "
tx3.Text = " "
tx4.Text = " "
tx5.Text = " "
tx6.Text = " "
tx7.Text = " "
tx8.Text = " "
tx9.Text = " "
tx10.Text = " "
Combo2.Text = " "
Text1.Text = " "
sr.Clear
id.Clear
rate.Clear
End Sub

Private Sub search_Click()
searchby.Visible = True
End Sub

Private Sub Tx2_KeyPress(KeyAscii As Integer)
If KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32 Then
        tx2.Locked = False
Else
        tx2.Locked = True
End If
If KeyAscii = 13 Or KeyAscii = 9 Then tx3.SetFocus


End Sub

Private Sub tx2_LostFocus()
tx2.Text = UCase(Mid(tx2.Text, 1, 1)) & Mid(tx2.Text, 2, Len(tx2.Text))
End Sub

Private Sub Tx3_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
    tx3.Locked = False
Else
    tx3.Locked = True
End If
If KeyAscii = 13 Then tx4.SetFocus
End Sub
Private Sub Tx4_KeyPress(KeyAscii As Integer)
If Len(tx4.Text) = 50 And KeyAscii <> 8 Then
    tx4.Locked = True
ElseIf KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii = 44 Then
    tx4.Locked = False
ElseIf KeyAscii = 13 Then
    tx5.SetFocus
Else
    tx4.Locked = True
End If
End Sub

Private Sub tx4_LostFocus()
tx4.Text = UCase(Mid(tx4.Text, 1, 1)) & Mid(tx4.Text, 2, Len(tx4.Text))
End Sub

Private Sub Tx5_KeyPress(KeyAscii As Integer)
If KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32 Then
        tx5.Locked = False
Else
        tx5.Locked = True
End If
If KeyAscii = 13 Then tx6.SetFocus
End Sub

Private Sub tx5_LostFocus()
tx5.Text = UCase(Mid(tx5.Text, 1, 1)) & Mid(tx5.Text, 2, Len(tx5.Text))
End Sub

Private Sub Tx6_KeyPress(KeyAscii As Integer)
If KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32 Then
        tx6.Locked = False
Else
        tx6.Locked = True
End If
If KeyAscii = 13 Then tx7.SetFocus
End Sub

Private Sub tx6_LostFocus()
tx6.Text = UCase(Mid(tx6.Text, 1, 1)) & Mid(tx6.Text, 2, Len(tx6.Text))
End Sub

Private Sub Tx7_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
    tx7.Locked = False
Else
    tx7.Locked = True
End If
If KeyAscii = 13 Then tx8.SetFocus
End Sub
Private Sub Tx8_KeyPress(KeyAscii As Integer)
If Len(tx8.Text) = 30 And KeyAscii <> 8 Then
    tx8.Locked = True
ElseIf KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii = 44 Then
    tx8.Locked = False
ElseIf KeyAscii = 13 Then
    tx9.SetFocus
Else
    tx8.Locked = True
End If
End Sub

Private Sub tx8_LostFocus()
tx8.Text = UCase(Mid(tx8.Text, 1, 1)) & Mid(tx8.Text, 2, Len(tx8.Text))
End Sub

Private Sub tx9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then tx10.SetFocus
End Sub

Private Sub update_Click()
Set R = New ADODB.Recordset
SQL = "update supplier set  SUP_NM='" + tx2.Text + "',SUP_MOB =" + tx3.Text + ", SUP_LOCATION='" + tx4.Text + "', SUP_STATE ='" + tx5.Text + "',SUP_CITY='" + tx6.Text + "',SUP_PINCODE=" + tx7.Text + ",COM='" + tx8.Text + "',SUP_EMAIL='" + tx9.Text + "',SUP_GSTNO='" + tx10.Text + "'where SUP_ID='" + Combo1.Text + "'"
Set R = C.Execute(SQL)
MsgBox "record updated"
Adodc1.Refresh
Label21.Caption = " "
tx2.Text = " "
tx3.Text = " "
tx4.Text = " "
tx5.Text = " "
tx6.Text = " "
tx7.Text = " "
tx8.Text = " "
tx9.Text = " "
tx10.Text = " "
End Sub

Public Function auto_sup_id()
Set R = New ADODB.Recordset
SQL = "select *from supplier"
Set R = C.Execute(SQL)
While R.EOF = False
Combo1.AddItem R.Fields(0)
R.MoveNext
Wend
End Function
Public Function auto_prd_id()
Set R = New ADODB.Recordset
SQL = "select *from product"
Set R = C.Execute(SQL)
While R.EOF = False
Combo2.AddItem R.Fields(0)
R.MoveNext
Wend
End Function

