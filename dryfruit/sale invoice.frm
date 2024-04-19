VERSION 5.00
Begin VB.Form sale_invoice 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Invoice"
   ClientHeight    =   10350
   ClientLeft      =   -750
   ClientTop       =   -2340
   ClientWidth     =   15630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10350
   ScaleWidth      =   15630
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List7 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1920
      ItemData        =   "sale invoice.frx":0000
      Left            =   13200
      List            =   "sale invoice.frx":0002
      TabIndex        =   6
      Top             =   4320
      Width           =   2175
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1920
      ItemData        =   "sale invoice.frx":0004
      Left            =   240
      List            =   "sale invoice.frx":0006
      TabIndex        =   5
      Top             =   4320
      Width           =   2655
   End
   Begin VB.ListBox List6 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1920
      ItemData        =   "sale invoice.frx":0008
      Left            =   11040
      List            =   "sale invoice.frx":000A
      TabIndex        =   4
      Top             =   4320
      Width           =   2175
   End
   Begin VB.ListBox List5 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1920
      ItemData        =   "sale invoice.frx":000C
      Left            =   8880
      List            =   "sale invoice.frx":000E
      TabIndex        =   3
      Top             =   4320
      Width           =   2175
   End
   Begin VB.ListBox List4 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1920
      ItemData        =   "sale invoice.frx":0010
      Left            =   6840
      List            =   "sale invoice.frx":0012
      TabIndex        =   2
      Top             =   4320
      Width           =   2055
   End
   Begin VB.ListBox List3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1920
      ItemData        =   "sale invoice.frx":0014
      Left            =   4800
      List            =   "sale invoice.frx":0016
      TabIndex        =   1
      Top             =   4320
      Width           =   2055
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1920
      ItemData        =   "sale invoice.frx":0018
      Left            =   2880
      List            =   "sale invoice.frx":001A
      TabIndex        =   0
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10935
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   15615
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Save/Print"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   13200
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   485
         Left            =   13560
         TabIndex        =   61
         Top             =   8760
         Width           =   1815
      End
      Begin VB.TextBox Text12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   485
         Left            =   13680
         TabIndex        =   31
         Top             =   6360
         Width           =   1695
      End
      Begin VB.TextBox paidby 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13560
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   9360
         Width           =   1815
      End
      Begin VB.TextBox Text15 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   485
         Left            =   13560
         TabIndex        =   29
         Top             =   8280
         Width           =   1815
      End
      Begin VB.TextBox Txt88 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   485
         Left            =   13680
         TabIndex        =   28
         Top             =   6840
         Width           =   1695
      End
      Begin VB.TextBox TXT10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   485
         Index           =   0
         Left            =   11880
         TabIndex        =   27
         Text            =   "SGST(Rs.)"
         Top             =   6360
         Width           =   1875
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   485
         Index           =   1
         Left            =   10200
         TabIndex        =   26
         Text            =   "CGST(Rs.)"
         Top             =   6360
         Width           =   1755
      End
      Begin VB.TextBox sgmt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   485
         Left            =   11880
         TabIndex        =   25
         Top             =   6840
         Width           =   1875
      End
      Begin VB.TextBox cgmt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   485
         Left            =   10200
         TabIndex        =   24
         Top             =   6840
         Width           =   1755
      End
      Begin VB.TextBox Txt90 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   485
         Left            =   240
         TabIndex        =   23
         Top             =   6360
         Width           =   1280
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1680
         TabIndex        =   22
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox TXT1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   12840
         TabIndex        =   21
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox Txt9 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   7320
         TabIndex        =   20
         Top             =   3120
         Width           =   1935
      End
      Begin VB.TextBox Txt8 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   12120
         TabIndex        =   19
         Top             =   3000
         Width           =   1935
      End
      Begin VB.TextBox TXT7 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   12120
         TabIndex        =   18
         Top             =   2640
         Width           =   3015
      End
      Begin VB.TextBox TXT4 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   17
         Top             =   3480
         Width           =   1935
      End
      Begin VB.TextBox TXT3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   16
         Top             =   3120
         Width           =   1935
      End
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   485
         Left            =   8760
         TabIndex        =   15
         Top             =   6360
         Width           =   1455
      End
      Begin VB.TextBox Text13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   485
         Left            =   13560
         TabIndex        =   14
         Top             =   7800
         Width           =   1815
      End
      Begin VB.TextBox Text6 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   11280
         TabIndex        =   13
         Top             =   2640
         Width           =   3255
      End
      Begin VB.TextBox Text8 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   11280
         TabIndex        =   12
         Top             =   3360
         Width           =   3855
      End
      Begin VB.TextBox Text2 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1800
         TabIndex        =   11
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   10
         Top             =   2760
         Width           =   1935
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   485
         Left            =   13560
         TabIndex        =   9
         Top             =   7320
         Width           =   1815
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         Caption         =   "Paid By:"
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
         Left            =   11040
         TabIndex        =   60
         Top             =   9360
         Width           =   2175
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Amount Paid(Rs.):"
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
         Left            =   11400
         TabIndex        =   59
         Top             =   8400
         Width           =   1935
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Sub Total(Rs.):"
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
         Left            =   11760
         TabIndex        =   58
         Top             =   7920
         Width           =   1560
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Terms and Conditions:"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   57
         Top             =   7560
         Width           =   2415
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
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
         Height          =   315
         Index           =   0
         Left            =   1680
         TabIndex        =   56
         Top             =   6480
         Width           =   600
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         Caption         =   "SALE INVOICE"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   7080
         TabIndex        =   55
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         Caption         =   "STATE-BIHAR"
         BeginProperty Font 
            Name            =   "Constantia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   270
         Left            =   7440
         TabIndex        =   54
         Top             =   1440
         Width           =   1515
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         Caption         =   "Invoice No.:"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   240
         TabIndex        =   53
         Top             =   2160
         Width           =   1065
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         Caption         =   "Customer Name:"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   1080
         TabIndex        =   52
         Top             =   3120
         Width           =   1515
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         Caption         =   "Gender:"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   1080
         TabIndex        =   51
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         Caption         =   "Mobile No.:"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   10080
         TabIndex        =   50
         Top             =   3360
         Width           =   1050
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         Caption         =   "Address:"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   10080
         TabIndex        =   49
         Top             =   2640
         Width           =   795
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         Caption         =   "Invoice Date:"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   11520
         TabIndex        =   48
         Top             =   2160
         Width           =   1185
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   15240
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H8000000B&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         Height          =   495
         Left            =   240
         Top             =   6360
         Width           =   15135
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H80000011&
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
         Height          =   315
         Index           =   1
         Left            =   1560
         TabIndex        =   47
         Top             =   6480
         Width           =   600
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H8000000E&
         Height          =   495
         Left            =   240
         Top             =   6840
         Width           =   15135
      End
      Begin VB.Line Line7 
         X1              =   11580
         X2              =   11580
         Y1              =   6360
         Y2              =   7320
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Total Tax Amount:"
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
         Left            =   8160
         TabIndex        =   46
         Top             =   6960
         Width           =   1980
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         Caption         =   "Mobile No.-8434991679"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Index           =   0
         Left            =   6960
         TabIndex        =   45
         Top             =   1800
         Width           =   2505
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "S.N ENTERPRISER, PATNA"
         BeginProperty Font 
            Name            =   "Constantia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   6840
         TabIndex        =   44
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         Caption         =   "Sell Order No.:"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   43
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Product name"
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
         Height          =   375
         Left            =   240
         TabIndex        =   42
         Top             =   3960
         Width           =   2655
      End
      Begin VB.Label Label20 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "  MRP"
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
         Height          =   375
         Left            =   2880
         TabIndex        =   41
         Top             =   3960
         Width           =   2055
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "  Quantity"
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
         Height          =   375
         Left            =   4800
         TabIndex        =   40
         Top             =   3960
         Width           =   2055
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "  Total Price"
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
         Height          =   375
         Left            =   6840
         TabIndex        =   39
         Top             =   3960
         Width           =   2055
      End
      Begin VB.Label Label23 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "  CGST"
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
         Height          =   375
         Left            =   8880
         TabIndex        =   38
         Top             =   3960
         Width           =   2175
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "  SGST"
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
         Height          =   375
         Left            =   11040
         TabIndex        =   37
         Top             =   3960
         Width           =   2175
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "  Net Amount"
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
         Height          =   375
         Left            =   13200
         TabIndex        =   36
         Top             =   3960
         Width           =   2175
      End
      Begin VB.Image Image1 
         Height          =   1200
         Left            =   120
         Picture         =   "sale invoice.frx":001C
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2520
      End
      Begin VB.Line Line3 
         X1              =   240
         X2              =   15360
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   $"sale invoice.frx":690E
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   240
         TabIndex        =   35
         Top             =   7920
         Width           =   3735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         Caption         =   "Customer ID:"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   1080
         TabIndex        =   34
         Top             =   2760
         Width           =   1200
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Previous Dues Amount(Rs.):"
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
         Left            =   10320
         TabIndex        =   33
         Top             =   7440
         Width           =   3000
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "New Dues Amount (Rs.):"
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
         Index           =   3
         Left            =   10680
         TabIndex        =   32
         Top             =   8880
         Width           =   2640
      End
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Dues Amount(Rs.):"
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
      TabIndex        =   7
      Top             =   6480
      Width           =   2010
   End
End
Attribute VB_Name = "sale_invoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Set R = New ADODB.Recordset
SQL = "insert into sell_inv values('" + Text2.Text + "','" + Format(txt1.Text, "dd-MMM-yyyy") + "','" + Text13.Text + "','" + Text1.Text + "')"
Set R = C.Execute(SQL)
MsgBox "Invoice Saved"
Me.PrintForm
End Sub


Private Sub Form_Load()
CONN
Text2.Text = Sell.Text1.Text

Dim a As String
Set R = New ADODB.Recordset
SQL = "select max(to_number(SUBSTR(inv_no,6,LENGTH(inv_no))))from sell_inv"
Set R = C.Execute(SQL)
If IsNull(R.Fields(0)) Then
Text1.Text = "INV" & "00" & 1
Else
Text1.Text = "INV" & "00" & R.Fields(0) + 1
a = Text1.Text
End If
If (a = "INV0010") Then
Set R = New ADODB.Recordset
SQL = "select max(to_number(SUBSTR(inv_no,5,LENGTH(inv_no))))from sell_inv"
Set R = C.Execute(SQL)
Text1.Text = "INV" & "0" & R.Fields(0) + 1
End If
txt1.Text = Date
End Sub

