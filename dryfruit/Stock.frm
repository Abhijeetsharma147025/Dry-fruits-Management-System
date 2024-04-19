VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Stock 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock"
   ClientHeight    =   6990
   ClientLeft      =   6465
   ClientTop       =   690
   ClientWidth     =   10305
   DrawStyle       =   6  'Inside Solid
   BeginProperty Font 
      Name            =   "@Malgun Gothic"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   10305
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   7080
      Top             =   5520
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
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
      RecordSource    =   "select *from stock"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "@Malgun Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Stock.frx":0000
      Height          =   1695
      Left            =   0
      TabIndex        =   22
      Top             =   4560
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   2990
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "@Malgun Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "RNO"
         Caption         =   "Rack No."
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
         DataField       =   "P_ID"
         Caption         =   "Product Id"
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
         DataField       =   "AVL_QTY"
         Caption         =   "Available Quantity"
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
         DataField       =   "MIN_QTY"
         Caption         =   "Minimum Quantity"
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
         DataField       =   "MAX_QTY"
         Caption         =   "Maximum Quantity"
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
            WrapText        =   -1  'True
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2039.811
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1635.024
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1635.024
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1635.024
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3360
      TabIndex        =   21
      Top             =   960
      Width           =   3135
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Add Rack"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   20
      Top             =   960
      Width           =   1575
   End
   Begin VB.ComboBox Combo2 
      Height          =   495
      Left            =   3360
      TabIndex        =   19
      Top             =   1680
      Width           =   3135
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Search"
      Height          =   1575
      Left            =   6720
      TabIndex        =   16
      Top             =   2160
      Width           =   3375
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
         Left            =   600
         TabIndex        =   18
         Text            =   "product id"
         Top             =   600
         Width           =   1575
      End
      Begin VB.Image close2 
         BorderStyle     =   1  'Fixed Single
         Height          =   510
         Left            =   2880
         Picture         =   "Stock.frx":0015
         Top             =   0
         Width           =   510
      End
      Begin VB.Label Label7 
         Caption         =   "Search By"
         BeginProperty Font 
            Name            =   "@Malgun Gothic"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   17
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   0
      TabIndex        =   9
      Top             =   6240
      Width           =   10335
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFF00&
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7200
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFFF00&
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFFF00&
         Caption         =   "Report"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFF00&
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFF00&
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFF00&
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   1215
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   10320
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H8000000E&
      Height          =   495
      Left            =   3360
      TabIndex        =   8
      Top             =   3840
      Width           =   3135
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H8000000E&
      Height          =   495
      Left            =   3360
      TabIndex        =   7
      Top             =   3120
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H8000000E&
      Height          =   495
      Left            =   3360
      TabIndex        =   6
      Top             =   2400
      Width           =   3135
   End
   Begin VB.Line Line1 
      DrawMode        =   6  'Mask Pen Not
      X1              =   0
      X2              =   10335
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000E&
      Caption         =   "Maximum Quantity :"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   3840
      Width           =   2895
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "Minimum Quantity : "
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   2895
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "Avilable Quantity    :     "
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      Caption         =   "Product Id                  :"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
      Caption         =   "Rack No                      :"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Stock:"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10455
   End
End
Attribute VB_Name = "Stock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub close2_Click()
Frame2.Visible = False
End Sub

Public Function auto_p_id()
Set R = New ADODB.Recordset
SQL = "select *from stock"
Set R = C.Execute(SQL)
While R.EOF = False
Combo1.AddItem R.Fields(1)
R.MoveNext
Wend
End Function

Private Sub Combo1_Click()
Set R = New ADODB.Recordset
SQL = "select *from stock where P_ID='" + Combo1.Text + "'"
Set R = C.Execute(SQL)
Text1.Text = R.Fields(0)
Combo2.Text = R.Fields(1)
Text3.Text = R.Fields(2)
Text4.Text = R.Fields(3)
Text5.Text = R.Fields(4)
End Sub

Private Sub Command2_Click()
If Text1.Text = blank Or Combo2.Text = blank Or Text3.Text = blank Or Text4.Text = blank Or Text5.Text = blank Then
MsgBox "Please enter the details first!!"
Else
Set R = New ADODB.Recordset
SQL = "insert into stock values('" + Text1.Text + "','" + Combo2.Text + "'," + Text3.Text + "," + Text4.Text + "," + Text5.Text + ")"
Set R = C.Execute(SQL)
MsgBox "data saved"
Unload Me
Stock.Show
Stock.Top = 0
Stock.Left = 0

End If
End Sub

Private Sub Command3_Click()
Set R = New ADODB.Recordset
SQL = "update stock set  RNO='" + Text1.Text + "',P_Id='" + Combo2.Text + "',AVL_QTY=" + Text3.Text + ",min_qty=" + Text4.Text + ",max_qty=" + Text5.Text + "  where P_ID='" + Combo1.Text + "'"
Set R = C.Execute(SQL)
MsgBox "record updated"
Adodc1.Refresh
Combo3.Text = " "
Combo2.Text = " "
Text3.Text = " "
Text4.Text = " "
Text5.Text = " "
End Sub

Private Sub Command4_Click()
Frame2.Visible = True
End Sub

Private Sub Command6_Click()
Set R = New ADODB.Recordset
SQL = "Delete from stock where p_id='" + Combo1.Text + "'"
Set R = C.Execute(SQL)
MsgBox "Record deleted"
Adodc1.Refresh
Combo3.Clear
auto_p_id
Combo2.Text = " "
Text3.Text = " "
Text4.Text = " "
Text5.Text = " "
End Sub

Private Sub Command7_Click()
Unload Me
End Sub

Private Sub Command8_Click()
Dim i As String
'i = Combo3.ListCount
'Combo3.AddItem i + 1
Set R = New ADODB.Recordset
SQL = "select count (rno) from stock"
Set R = C.Execute(SQL)
i = R.Fields(0)


Text1.Text = i + 1

End Sub

Private Sub Form_Load()
CONN
Text1.Locked = True
Adodc1.Visible = False
Set R = New ADODB.Recordset
SQL = "select distinct p_id from product where p_id not in (select p_id from stock)"
Set R = C.Execute(SQL)
While R.EOF = False
Combo2.AddItem R.Fields(0)
R.MoveNext
Wend
auto_p_id
'Text1.SetFocus
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text2.SetFocus
End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text3.SetFocus
End If
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text4.SetFocus
End If
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text5.SetFocus
End If
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command2.SetFocus
End If
End Sub
