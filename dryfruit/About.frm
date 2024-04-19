VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form About 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10110
   BeginProperty Font 
      Name            =   "Palatino Linotype"
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
   ScaleHeight     =   5655
   ScaleWidth      =   10110
   ShowInTaskbar   =   0   'False
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   9551
      _Version        =   393217
      TextRTF         =   $"About.frx":0000
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
