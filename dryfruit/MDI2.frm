VERSION 5.00
Begin VB.MDIForm mdi 
   BackColor       =   &H8000000C&
   Caption         =   "Dry Fruits Managment System"
   ClientHeight    =   9645
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   19980
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDI2.frx":0000
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   0
      Picture         =   "MDI2.frx":2EEA3
      ScaleHeight     =   1275
      ScaleWidth      =   19920
      TabIndex        =   0
      Top             =   0
      Width           =   19980
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   17160
         Top             =   720
      End
      Begin VB.CommandButton Sup_but 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Supplier"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   1320
         Picture         =   "MDI2.frx":422C6
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   1335
      End
      Begin VB.CommandButton Pur_but 
         BackColor       =   &H8000000E&
         Caption         =   "Order"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   0
         Picture         =   "MDI2.frx":43155
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cust_but 
         BackColor       =   &H8000000E&
         Caption         =   "Customer"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   5280
         Picture         =   "MDI2.frx":43A09
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton stock_but 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Stock In"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   3960
         Picture         =   "MDI2.frx":4424F
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   1335
      End
      Begin VB.CommandButton Sale_but 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sales"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   6600
         Picture         =   "MDI2.frx":450F1
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   1335
      End
      Begin VB.CommandButton report_but 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reports"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   7920
         Picture         =   "MDI2.frx":45E9D
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   1335
      End
      Begin VB.CommandButton prod_but 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   2640
         Picture         =   "MDI2.frx":46FD6
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   17880
         TabIndex        =   10
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   12000
         TabIndex        =   9
         Top             =   720
         Width           =   5655
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "             DRY FRUITS DISTRIBUTOR MANAGEMENT SYSTEM"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   375
         Left            =   10920
         TabIndex        =   8
         Top             =   240
         Width           =   9255
      End
   End
   Begin VB.Menu Account 
      Caption         =   "Account"
      Begin VB.Menu rpluser 
         Caption         =   "Replace User"
      End
      Begin VB.Menu Frgtpsswd 
         Caption         =   "Forget Password"
      End
   End
   Begin VB.Menu Prd 
      Caption         =   "Product"
      Begin VB.Menu Prd_ct 
         Caption         =   "Product"
      End
      Begin VB.Menu St_ock 
         Caption         =   "Stock"
      End
   End
   Begin VB.Menu Supp_lier 
      Caption         =   "Supplier"
   End
   Begin VB.Menu Pur_chase 
      Caption         =   "Purchase"
      Begin VB.Menu Or_der 
         Caption         =   "Order"
      End
      Begin VB.Menu Purc_hase 
         Caption         =   "Purchase"
      End
   End
   Begin VB.Menu Cus_tomer 
      Caption         =   "Customer"
   End
   Begin VB.Menu S_ell 
      Caption         =   "Sell"
      Begin VB.Menu Orde_r 
         Caption         =   "Order"
      End
      Begin VB.Menu S_inv 
         Caption         =   "Sale Invoice"
      End
   End
   Begin VB.Menu Rep_ort 
      Caption         =   "Report"
      Begin VB.Menu Prd_rept 
         Caption         =   "Product Report"
      End
      Begin VB.Menu Cus_rept 
         Caption         =   "Customer Report"
      End
      Begin VB.Menu Stc_rpt 
         Caption         =   "Stock Report"
      End
      Begin VB.Menu Supp_rpt 
         Caption         =   "Supplier Report"
      End
      Begin VB.Menu Pur_rpt 
         Caption         =   "Purchase Report"
      End
      Begin VB.Menu Pur_st_rpt 
         Caption         =   "Purchase Status Report"
      End
      Begin VB.Menu Sal_rpt 
         Caption         =   "Sales Report"
      End
   End
   Begin VB.Menu Ab_out 
      Caption         =   "About"
   End
End
Attribute VB_Name = "mdi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Ab_out_Click()
Unload Me
Load About
About.Show
End Sub

Private Sub Cus_rept_Click()
Unload Me
Load Report
Report.Show
Report.Frame4.Visible = True
Report.Combo5.Clear
Report.Combo5.AddItem "ID"
Report.Combo5.AddItem "MOBILE"
Report.Combo5.AddItem "NAME"

End Sub

Private Sub Cus_tomer_Click()
Unload Me
Load customer
customer.Show
End Sub

Private Sub cust_but_Click()
Unload Me
Load customer
customer.Show
End Sub

Private Sub Frgtpsswd_Click()
Unload Me
Load forget
forget.Show
End Sub

Private Sub MDIForm_Load()
Label2.Caption = Format$(Now, "dd mmmm,yyyy" & Space(19) & "DDDD")
End Sub

Private Sub Or_der_Click()
Unload Me
Load order
order.Show
End Sub

Private Sub Orde_r_Click()
Unload Me
Load Sell
Sell.Show

End Sub

Private Sub Prd_ct_Click()
Unload Me
Load Product
Product.Show
End Sub

Private Sub Prd_rept_Click()
Unload Me
Load Report
Report.Show
Report.frame3.Visible = True
Report.Combo1.Clear
Report.Combo1.AddItem "ID"
Report.Combo1.AddItem "NAME"
Report.Combo1.AddItem "COMPANY"
Report.Combo1.AddItem "TYPE"
End Sub

Private Sub prod_but_Click()
Unload Me
Load Product
Product.Show
End Sub

Private Sub Pur_but_Click()
Unload Me
Load order
order.Show
End Sub

Private Sub Pur_rpt_Click()
Unload Me
Load Report
Report.Show
Report.Frame9.Visible = True

Report.Combo9.Clear
Report.Combo9.AddItem "Order No"
Report.Combo9.AddItem "Supplier Id"
Report.Combo9.AddItem "Date"
Report.Combo9.AddItem "Month"
Report.Combo9.AddItem "Between Dates"

End Sub

Private Sub Pur_st_rpt_Click()
Unload Me
Load Report
Report.Show
Report.Frame5.Visible = True
Report.Combo6.Clear
Report.Combo6.AddItem "Order No"
Report.Combo6.AddItem "Invoice No."
Report.Combo6.AddItem "Invoice Date"
Report.Combo6.AddItem "Month"
Report.Combo6.AddItem "Between Dates"
End Sub

Private Sub Purc_hase_Click()
Unload Me
Load purchase
purchase.Show
End Sub

Private Sub report_but_Click()
Unload Me
Load Report
Report.Show
End Sub

Private Sub rpluser_Click()
Unload Me
Load Replace_user
Replace_user.Show
End Sub

Private Sub S_inv_Click()
Unload Me
Load sale_invoice
sale_invoice.Show
sale_invoice.Top = 0
sale_invoice.Left = 0
End Sub

Private Sub Sal_rpt_Click()
Unload Me
Load Report
Report.Show
Report.Frame101.Visible = True
Report.sale1.Clear
Report.sale1.AddItem "Order Id"
Report.sale1.AddItem "Customer Id"
Report.sale1.AddItem "Date"
Report.sale1.AddItem "Month"
Report.sale1.AddItem "Between Dates"

End Sub

Private Sub Sale_but_Click()
Unload Me
Load Sell
Sell.Show

End Sub

Private Sub St_ock_Click()
Unload Me
Load Stock
Stock.Show
End Sub

Private Sub Stc_rpt_Click()
Unload Me
Load Report
Report.Show
Report.Frame2.Visible = True
Report.STK1.Clear

Report.STK1.AddItem "ID"
Report.STK1.AddItem "RACK NO"
End Sub

Private Sub stock_but_Click()
Unload Me
Load Stock
Stock.Show
End Sub

Private Sub Sup_but_Click()
Unload Me
Load supplier
supplier.Show

End Sub

Private Sub Supp_lier_Click()
Unload Me
Load supplier
supplier.Show
End Sub

Private Sub Supp_rpt_Click()
Unload Me
Load Report
Report.Show
Report.Frame7.Visible = True
Report.Combo12.Clear
Report.Combo12.AddItem "ID"
Report.Combo12.AddItem "NAME"
Report.Combo12.AddItem "MOBILE NO."
End Sub

Private Sub Timer1_Timer()
Label3.Caption = Format$(Time$, "hh:mm:ss AM/PM")
End Sub
