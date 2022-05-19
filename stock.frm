VERSION 5.00
Begin VB.Form stockdetails 
   BackColor       =   &H008080FF&
   Caption         =   "Stock details"
   ClientHeight    =   6090
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11085
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
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   780
      Left            =   9480
      TabIndex        =   2
      Top             =   0
      Width           =   1575
   End
   Begin VB.PictureBox DataGrid1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   0
      ScaleHeight     =   5235
      ScaleWidth      =   10995
      TabIndex        =   1
      Top             =   840
      Width           =   11055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Stock details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "stockdetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim con As New ADODB.Connection
Dim rec As New ADODB.Recordset

Private Sub Command1_Click()
home.Show
Me.Hide
End Sub

Private Sub Form_Load()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=A:\giftshop\Database2.mdb;Persist Security Info=False"
rec.CursorLocation = adUseClient
rec.Open "Select * from stock", con, adOpenDynamic, adLockOptimistic
Set DataGrid1.DataSource = rec
End Sub

