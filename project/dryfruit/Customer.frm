VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form customer 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7545
   ClientLeft      =   6465
   ClientTop       =   1905
   ClientWidth     =   11730
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   11730
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   4680
      Top             =   5400
      Width           =   1215
      _ExtentX        =   2143
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
      RecordSource    =   "select *from customer"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox tx4 
      Height          =   735
      Left            =   2880
      TabIndex        =   28
      Top             =   2400
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1296
      _Version        =   393217
      TextRTF         =   $"Customer.frx":0000
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Customer.frx":0084
      Height          =   1815
      Left            =   0
      TabIndex        =   27
      Top             =   4920
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   3201
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   26
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "C_ID"
         Caption         =   "Id"
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
         DataField       =   "C_NM"
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
         DataField       =   "C_MOB"
         Caption         =   "Mobile No."
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
         DataField       =   "C_ADD"
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
         DataField       =   "C_GENDER"
         Caption         =   "Gender"
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
         DataField       =   "C_EMAIL"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1950.236
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2039.811
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2039.811
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1635.024
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   3119.811
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      Height          =   420
      Left            =   7920
      TabIndex        =   26
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   975
      Left            =   0
      TabIndex        =   12
      Top             =   6720
      Width           =   11775
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Close"
         Height          =   540
         Left            =   7200
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Delete"
         Height          =   540
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Search"
         Height          =   540
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Update"
         Height          =   540
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Save"
         Height          =   540
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Add New "
         Height          =   540
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search"
      Height          =   1575
      Left            =   7800
      TabIndex        =   9
      Top             =   2400
      Width           =   3495
      Begin VB.ComboBox Combo1 
         Height          =   420
         Left            =   480
         TabIndex        =   11
         Text            =   "Select"
         Top             =   720
         Width           =   2055
      End
      Begin VB.Image close2 
         BorderStyle     =   1  'Fixed Single
         Height          =   510
         Left            =   3000
         Picture         =   "Customer.frx":0099
         Top             =   0
         Width           =   510
      End
      Begin VB.Label Label10 
         Caption         =   "Search by"
         Height          =   375
         Left            =   480
         TabIndex        =   10
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.TextBox Tx8 
      Height          =   420
      Left            =   2880
      TabIndex        =   8
      Top             =   4200
      Width           =   4575
   End
   Begin VB.TextBox Tx3 
      Height          =   420
      Left            =   2880
      MaxLength       =   10
      TabIndex        =   7
      Top             =   1680
      Width           =   4575
   End
   Begin VB.TextBox Tx2 
      Height          =   420
      Left            =   2880
      TabIndex        =   6
      Top             =   1080
      Width           =   4575
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   2880
      TabIndex        =   21
      Top             =   3360
      Width           =   4575
      Begin VB.OptionButton Option3 
         Caption         =   "Transgender"
         Height          =   615
         Left            =   2640
         TabIndex        =   24
         Top             =   0
         Width           =   1935
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Female"
         Height          =   615
         Left            =   1200
         TabIndex        =   23
         Top             =   0
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Male"
         Height          =   615
         Left            =   120
         TabIndex        =   22
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.Label Label8 
      Caption         =   "0"
      Height          =   255
      Left            =   8400
      TabIndex        =   19
      Top             =   2040
      Width           =   135
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   11760
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label7 
      Height          =   375
      Left            =   9360
      TabIndex        =   25
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Gender                 :   "
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   20
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Email Id                :"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   4200
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Address                :   "
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2520
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile No           :"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Name                    :"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Id :*"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "  Customer :"
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
      Width           =   11775
   End
End
Attribute VB_Name = "customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub close2_Click()
Frame1.Visible = False
End Sub

Private Sub Combo1_Click()
Set R = New ADODB.Recordset
SQL = "select *from customer where c_id='" + Combo1.Text + "'"
Set R = C.Execute(SQL)
Label7.Caption = R.Fields(0)
tx2.Text = R.Fields(1)
tx3.Text = R.Fields(2)
tx4.Text = R.Fields(3)
If (IsNull(R.Fields(5))) Then
tx8.Text = ""
Else
tx8.Text = R.Fields(5)
End If
Text1.Text = R.Fields(4)


End Sub

Private Sub Command1_Click()
Dim a As String
Set R = New ADODB.Recordset
SQL = "select max(to_number(SUBSTR(c_id,4,LENGTH(c_id))))from customer"
Set R = C.Execute(SQL)
If IsNull(R.Fields(0)) Then
Label7.Caption = "C" & "00" & 1
Else
Label7.Caption = "C" & "00" & R.Fields(0) + 1
a = Label7.Caption
End If
If (a = "C0010") Then
Set R = New ADODB.Recordset
SQL = "select max(to_number(SUBSTR(c_id,3,LENGTH(c_id))))from customer"
Set R = C.Execute(SQL)
Label7.Caption = "C" & "0" & R.Fields(0) + 1
End If
tx2.Text = " "
tx3.Text = ""
tx4.Text = ""
tx8.Text = ""
Text1.Text = " "
tx2.SetFocus

End Sub

Private Sub Command2_Click()
If Label7.Caption = blank Or tx2.Text = blank Or tx3.Text = blank Or tx4.Text = blank Or Text1.Text = blank Then
MsgBox "Please fill the details first!!"
Else
Set R = New ADODB.Recordset
SQL = "insert into customer values('" + Label7.Caption + "','" + tx2.Text + "','" + tx3.Text + "','" + tx4.Text + "','" + Text1.Text + "','" + tx8.Text + "'," + Label8.Caption + ")"
Set R = C.Execute(SQL)
MsgBox "data saved"
Adodc1.Refresh

Label7.Caption = " "
tx2.Text = " "
tx3.Text = " "
tx4.Text = " "
tx8.Text = " "
Text1.Text = " "
Option1.Value = 0
Option2.Value = 0
Option3.Value = 0
Combo1.Clear
auto_c_id
End If
Sell.Combo6.Clear
Sell.auto_c_id
End Sub


Private Sub Command3_Click()
Set R = New ADODB.Recordset
SQL = "update customer set c_id='" + Label7.Caption + "',c_nm='" + tx2.Text + "',c_mob='" + tx3.Text + "',c_add='" + tx4.Text + "',c_gender='" + Text1.Text + "',c_email='" + tx8.Text + "' where c_id='" + Label7.Caption + "'"
Set R = C.Execute(SQL)
MsgBox "record updated..."
Adodc1.Refresh
End Sub

Private Sub Command4_Click()
Frame1.Visible = True
End Sub

Private Sub Command6_Click()
Set R = New ADODB.Recordset
SQL = "delete from customer where c_id='" + Combo1.Text + "'"
Set R = C.Execute(SQL)
MsgBox "Customer record deleted...!!"
End Sub

Private Sub Command7_Click()
Unload Me
End Sub

Private Sub Form_Load()
CONN
Frame1.Visible = False
auto_c_id
Adodc1.Visible = False
Text1.Visible = False
Label8.Visible = False
End Sub

Private Sub Option1_Click()
Text1.Text = Option1.Caption
End Sub

Private Sub Option2_Click()
Text1.Text = Option2.Caption
End Sub

Private Sub Option3_Click()
Text1.Text = Option3.Caption
End Sub

Private Sub Text1_Change()
If Text1.Text = "Male" Then
Option1.Value = True
Option2.Value = False
Option3.Value = False
ElseIf Text1.Text = "Female" Then
Option1.Value = False
Option2.Value = True
Option3.Value = False
ElseIf Text1.Text = "Transgender" Then
Option1.Value = False
Option2.Value = False
Option3.Value = True
End If
End Sub

Private Sub Tx2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
tx3.SetFocus
End If
End Sub

Private Sub tx2_LostFocus()
tx2.Text = UCase(Mid(tx2.Text, 1, 1)) & Mid(tx2.Text, 2, Len(tx2.Text))
End Sub

Private Sub Tx3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
tx4.SetFocus
End If
End Sub

Private Sub tx3_LostFocus()
tx3.Text = UCase(Mid(tx3.Text, 1, 1)) & Mid(tx3.Text, 2, Len(tx3.Text))
End Sub

Private Sub Tx4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
tx4.SetFocus
End If
End Sub
Private Sub Tx5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
tx6.SetFocus
End If
End Sub
Private Sub Tx6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
tx7.SetFocus
End If
End Sub
Private Sub Tx7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
tx8.SetFocus
End If
End Sub

Private Sub tx4_LostFocus()
tx4.Text = UCase(Mid(tx4.Text, 1, 1)) & Mid(tx4.Text, 2, Len(tx4.Text))
End Sub

Private Sub Tx8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1.SetFocus
End If
End Sub

Public Function auto_c_id()
Set R = New ADODB.Recordset
SQL = "select *from customer order by c_id"
Set R = C.Execute(SQL)
While R.EOF = False
Combo1.AddItem R.Fields(0)
R.MoveNext
Wend
End Function

