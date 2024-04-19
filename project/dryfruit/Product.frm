VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Product 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Product"
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15150
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   15150
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   8415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16935
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   495
         Left            =   4080
         Top             =   6240
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   873
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
         RecordSource    =   "select *from product"
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
         Bindings        =   "Product.frx":0000
         Height          =   1815
         Left            =   0
         TabIndex        =   40
         Top             =   5280
         Width           =   15135
         _ExtentX        =   26696
         _ExtentY        =   3201
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         Appearance      =   0
         BackColor       =   16709579
         HeadLines       =   1
         RowHeight       =   27
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "P_ID"
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
            DataField       =   "P_NM"
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
            DataField       =   "P_TYPE"
            Caption         =   "Type"
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
            DataField       =   "P_COMP"
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
         BeginProperty Column04 
            DataField       =   "P_WT"
            Caption         =   "Weight"
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
            DataField       =   "P_GST"
            Caption         =   "G.S.T"
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
            DataField       =   "P_RATE"
            Caption         =   "Rate"
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
            DataField       =   "P_UNIT"
            Caption         =   "Unit"
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
            DataField       =   "P_HSN"
            Caption         =   "HSN Code"
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
               Alignment       =   2
               WrapText        =   -1  'True
               ColumnWidth     =   1409.953
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2174.74
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1769.953
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2700.284
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1530.142
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1184.882
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1440
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1184.882
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1425.26
            EndProperty
         EndProperty
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
         Left            =   10320
         MaxLength       =   6
         TabIndex        =   39
         Top             =   1080
         Width           =   1815
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
         TabIndex        =   1
         Text            =   " "
         Top             =   1200
         Width           =   4455
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
         Left            =   3000
         MaxLength       =   6
         TabIndex        =   14
         Top             =   3120
         Width           =   1815
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
         TabIndex        =   13
         Text            =   " "
         Top             =   2640
         Width           =   2055
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
         MaxLength       =   20
         TabIndex        =   12
         Top             =   2160
         Width           =   4455
      End
      Begin VB.TextBox tx8 
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
         MaxLength       =   5
         TabIndex        =   11
         Top             =   4200
         Width           =   1695
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
         MaxLength       =   10
         TabIndex        =   10
         Text            =   " "
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox tx7 
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
         TabIndex        =   9
         Top             =   3600
         Width           =   4815
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   1215
         Left            =   0
         TabIndex        =   5
         Top             =   7080
         Width           =   15375
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
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   360
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
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   360
            Width           =   1455
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
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   360
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
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   360
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
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   360
            Width           =   1455
         End
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
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   360
            Width           =   1455
         End
         Begin VB.Line Line2 
            X1              =   -120
            X2              =   15120
            Y1              =   0
            Y2              =   0
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
         Height          =   1575
         Left            =   5160
         TabIndex        =   2
         Top             =   4560
         Width           =   3975
         Begin VB.ComboBox combo1 
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
            TabIndex        =   3
            Text            =   "Product_Id"
            Top             =   840
            Width           =   2415
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
            TabIndex        =   4
            Top             =   480
            Width           =   1935
         End
         Begin VB.Image close2 
            BorderStyle     =   1  'Fixed Single
            Height          =   510
            Left            =   3480
            Picture         =   "Product.frx":0015
            Top             =   120
            Width           =   510
         End
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   255
         Left            =   4920
         TabIndex        =   41
         Top             =   3240
         Width           =   255
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "G.S.T Rate                      :"
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
         TabIndex        =   29
         Top             =   3240
         Width           =   2565
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Package weight             :"
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
         TabIndex        =   38
         Top             =   2760
         Width           =   2580
      End
      Begin VB.Label Label22 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Product id:"
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
         Left            =   11520
         TabIndex        =   37
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label18 
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
         Left            =   12720
         TabIndex        =   36
         Top             =   0
         Width           =   135
      End
      Begin VB.Label Label12 
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
         Left            =   13200
         TabIndex        =   35
         Top             =   120
         Width           =   1695
      End
      Begin VB.Image Image1 
         Height          =   3540
         Left            =   8160
         OLEDragMode     =   1  'Automatic
         Picture         =   "Product.frx":0557
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   6420
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "  Product :"
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
         TabIndex        =   31
         Top             =   0
         Width           =   15375
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
         TabIndex        =   30
         Top             =   240
         Width           =   1695
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   15360
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Product weight"
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
         TabIndex        =   28
         Top             =   2760
         Width           =   15
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Product name"
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
         TabIndex        =   27
         Top             =   1320
         Width           =   2355
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Product Type"
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
         TabIndex        =   26
         Top             =   1800
         Width           =   1515
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Company Name            :"
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
         TabIndex        =   24
         Top             =   2280
         Width           =   2595
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "HSN Code      :"
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
         Left            =   8520
         TabIndex        =   23
         Top             =   1080
         Width           =   1605
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
         TabIndex        =   22
         Top             =   120
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
         TabIndex        =   21
         Top             =   1200
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
         TabIndex        =   20
         Top             =   1680
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
         TabIndex        =   19
         Top             =   2040
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
         Left            =   2160
         TabIndex        =   18
         Top             =   3600
         Width           =   135
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
         Left            =   9840
         TabIndex        =   17
         Top             =   960
         Width           =   135
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Unit Of Measurement  :"
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
         Top             =   4200
         Width           =   2625
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
         TabIndex        =   15
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Purchase Rate                :"
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
         TabIndex        =   25
         Top             =   3720
         Width           =   2580
      End
   End
End
Attribute VB_Name = "Product"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub add_Click()
Dim a As String
Set R = New ADODB.Recordset
SQL = "select max(to_number(SUBSTR(p_id,4,LENGTH(p_id))))from product"
Set R = C.Execute(SQL)
If IsNull(R.Fields(0)) Then
Label12.Caption = "P" & "00" & 1
Else
Label12.Caption = "P" & "00" & R.Fields(0) + 1
a = Label12.Caption
End If
If (a = "P0010") Then
Set R = New ADODB.Recordset
SQL = "select max(to_number(SUBSTR(p_id,3,LENGTH(p_id))))from product"
Set R = C.Execute(SQL)
Label12.Caption = "P" & "0" & R.Fields(0) + 1
End If
tx2.SetFocus
End Sub

Private Sub Close_Click()
Unload Me
End Sub

Private Sub Combo1_Click()
Set R = New ADODB.Recordset
SQL = "select *from product where p_id='" + Combo1.Text + "'"
Set R = C.Execute(SQL)
Label12.Caption = R.Fields(0)
tx2.Text = R.Fields(1)
tx3.Text = R.Fields(2)
tx4.Text = R.Fields(3)
tx5.Text = R.Fields(4)
tx6.Text = R.Fields(5)
tx7.Text = R.Fields(6)
tx8.Text = R.Fields(7)
tx9.Text = R.Fields(8)
End Sub

Private Sub close2_Click()
searchby.Visible = False
End Sub

Private Sub delete_Click()
Set R = New ADODB.Recordset
SQL = "delete from product where p_id='" + Combo1.Text + "'"
Set R = C.Execute(SQL)
MsgBox "record deleted"
Adodc1.Refresh
Combo1.Clear
refresh_search
Label12.Caption = " "
tx2.Text = " "
tx3.Text = " "
tx4.Text = " "
tx5.Text = " "
tx6.Text = " "
tx7.Text = " "
tx8.Text = " "
tx9.Text = " "

End Sub

Private Sub Form_Load()
CONN
searchby.Visible = False
Adodc1.Visible = False
refresh_search
End Sub

Private Sub save_Click()
If Label12.Caption = blank Or tx2.Text = " " Or tx3.Text = " " Or tx4.Text = " " Or tx5.Text = " " Or tx6.Text = " " Or tx7.Text = " " Or tx8.Text = " " Or tx9.Text = " " Then
MsgBox "please fill all the data"
Else
Set R = New ADODB.Recordset
SQL = "insert into product values('" + Label12.Caption + "','" + tx2.Text + "','" + tx3.Text + "','" + tx4.Text + "'," + tx5.Text + "," + tx6.Text + ",'" + tx7.Text + "','" + tx8.Text + "','" + tx9.Text + "')"
Set R = C.Execute(SQL)
MsgBox "record saved"
End If
Adodc1.Refresh
Combo1.Clear
refresh_search
Label12.Caption = " "
tx2.Text = " "
tx3.Text = " "
tx4.Text = " "
tx5.Text = " "
tx6.Text = " "
tx7.Text = " "
tx8.Text = " "
tx9.Text = " "
tx2.SetFocus

End Sub

Private Sub search_Click()
searchby.Visible = True
End Sub

Private Sub Tx2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then tx3.SetFocus
End Sub

Private Sub tx2_LostFocus()
tx2.Text = UCase(Mid(tx2.Text, 1, 1)) & Mid(tx2.Text, 2, Len(tx2.Text))
End Sub

Private Sub Tx3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then tx4.SetFocus
End Sub

Private Sub tx3_LostFocus()
tx3.Text = UCase(Mid(tx3.Text, 1, 1)) & Mid(tx3.Text, 2, Len(tx3.Text))
End Sub

Private Sub Tx4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then tx5.SetFocus
End Sub

Private Sub tx4_LostFocus()
tx4.Text = UCase(Mid(tx4.Text, 1, 1)) & Mid(tx4.Text, 2, Len(tx4.Text))
End Sub

Private Sub Tx5_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
    tx5.Locked = False
Else
    tx5.Locked = True
End If
If KeyAscii = 13 Then tx6.SetFocus
End Sub

Private Sub Tx6_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
    tx6.Locked = False
Else
    tx6.Locked = True
End If
If KeyAscii = 13 Then tx7.SetFocus
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
If KeyAscii = 13 Or KeyAscii = 9 Then tx9.SetFocus
End Sub

Private Sub tx8_LostFocus()
tx8.Text = UCase(Mid(tx8.Text, 1, 1)) & Mid(tx8.Text, 2, Len(tx8.Text))
End Sub

Private Sub tx9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then save.SetFocus
End Sub


Public Function refresh_search()
Set R = New ADODB.Recordset
SQL = "select *from product"
Set R = C.Execute(SQL)
While R.EOF = False
Combo1.AddItem R.Fields(0)
R.MoveNext
Wend
End Function

Private Sub update_Click()
Set R = New ADODB.Recordset
SQL = "update product set  p_nm='" + tx2.Text + "',p_type ='" + tx3.Text + "', p_comp='" + tx4.Text + "', p_wt =" + tx5.Text + ",p_gst=" + tx6.Text + ",p_rate=" + tx7.Text + ",p_unit='" + tx8.Text + "',p_hsn='" + tx9.Text + "' where p_id='" + Combo1.Text + "'"
Set R = C.Execute(SQL)
MsgBox "record updated"
Adodc1.Refresh
Label12.Caption = " "
tx2.Text = " "
tx3.Text = " "
tx4.Text = " "
tx5.Text = " "
tx6.Text = " "
tx7.Text = " "
tx8.Text = " "
tx9.Text = " "

End Sub
