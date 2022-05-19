VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form billing 
   BackColor       =   &H008080FF&
   Caption         =   "Billing "
   ClientHeight    =   6720
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   10425
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      BackColor       =   &H000000C0&
      Caption         =   "Back to Home"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14280
      TabIndex        =   23
      Top             =   0
      Width           =   2175
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3135
      Left            =   0
      TabIndex        =   22
      Top             =   3720
      Width           =   16335
      _ExtentX        =   28813
      _ExtentY        =   5530
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   30
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txttotal 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   9120
      TabIndex        =   21
      Top             =   6960
      Width           =   2535
   End
   Begin VB.TextBox txtgst 
      Height          =   405
      Left            =   11160
      TabIndex        =   20
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox txtprice 
      Height          =   405
      Left            =   11160
      TabIndex        =   19
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox txtqty 
      Height          =   405
      Left            =   5640
      TabIndex        =   18
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   6120
      TabIndex        =   17
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txtmobileno 
      Height          =   405
      Left            =   2640
      MaxLength       =   10
      TabIndex        =   16
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox txtitem 
      Height          =   405
      Left            =   2640
      TabIndex        =   15
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox txtname 
      Height          =   405
      Left            =   2640
      TabIndex        =   14
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Receipt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12000
      TabIndex        =   13
      Top             =   6960
      Width           =   1815
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   7680
      TabIndex        =   12
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdditem 
      Caption         =   "Add Item"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   840
      TabIndex        =   11
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   2640
      TabIndex        =   10
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdpre 
      Caption         =   "Prev"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   4560
      TabIndex        =   9
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label lblid 
      BackColor       =   &H008080FF&
      Height          =   495
      Left            =   2640
      TabIndex        =   24
      Top             =   360
      Width           =   2175
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   16080
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line1 
      X1              =   8040
      X2              =   8040
      Y1              =   0
      Y2              =   2880
   End
   Begin VB.Label lbldate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile No"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   600
      TabIndex        =   8
      Top             =   2280
      Width           =   1110
   End
   Begin VB.Label lbldate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Price"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   7
      Left            =   7320
      TabIndex        =   7
      Top             =   7080
      Width           =   1350
   End
   Begin VB.Label lbldate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gst/taxes"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   9360
      TabIndex        =   6
      Top             =   1560
      Width           =   1035
   End
   Begin VB.Label lbldate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Price"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   9360
      TabIndex        =   5
      Top             =   840
      Width           =   1140
   End
   Begin VB.Label lbldate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   5040
      TabIndex        =   4
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lbldate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill no"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   600
      TabIndex        =   3
      Top             =   480
      Width           =   720
   End
   Begin VB.Label lbldate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   600
      TabIndex        =   2
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lbldate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Customer name"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   1905
   End
   Begin VB.Label lblbilling 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "billing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rec As New ADODB.Recordset
Private Sub cmdAdditem_Click(Index As Integer)
txtitem = ""
txtprice = ""
txtqty = ""
txtgst = ""
End Sub

Private Sub cmdnew_Click(Index As Integer)
rec.AddNew
End Sub

Private Sub cmdnext_Click(Index As Integer)
rec.MoveNext
If rec.EOF Then rec.MoveLast
End Sub

Private Sub cmdpre_Click(Index As Integer)
rec.MovePrevious
If rec.BOF Then rec.MoveFirst
End Sub

Private Sub cmdsave_Click(Index As Integer)
If txtname.Text = "" Or txtitem.Text = "" Or txtprice.Text = "" Or txtqty.Text = "" Then
MsgBox "Fill the complite records please"
Else
rec.Update
MsgBox "records is saved"
End If

End Sub

Private Sub Command3_Click()
home.Show
Me.Hide
End Sub

Private Sub Form_Load()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=A:\giftshop\Database2.mdb;Persist Security Info=False"
rec.CursorLocation = adUseClient
rec.Open "Select * from CustomerBill", con, adOpenDynamic, adLockOptimistic
Set DataGrid1.DataSource = rec
Set lblid.DataSource = rec
lblid.DataField = "BillNo"
Set txtname.DataSource = rec
txtname.DataField = "Cust_Name"
Set txtitem.DataSource = rec
txtitem.DataField = "Item"
Set txtprice.DataSource = rec
txtprice.DataField = "TotalPrice"
Set txtqty.DataSource = rec
txtqty.DataField = "ItemQty"
Set txtgst.DataSource = rec
txtgst.DataField = "Gst"
Set txtmobileno.DataSource = rec
txtmobileno.DataField = "Mob"
End Sub
