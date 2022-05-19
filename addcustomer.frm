VERSION 5.00
Begin VB.Form addemployee 
   BackColor       =   &H008080FF&
   Caption         =   "Employee registration"
   ClientHeight    =   7365
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   11295
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdfirst 
      BackColor       =   &H008080FF&
      Caption         =   "First"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   3240
      TabIndex        =   17
      Top             =   5760
      Width           =   2775
   End
   Begin VB.CommandButton cmdlast 
      BackColor       =   &H008080FF&
      Caption         =   "Last"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   8040
      TabIndex        =   16
      Top             =   5760
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9840
      TabIndex        =   15
      Top             =   0
      Width           =   1455
   End
   Begin VB.TextBox txthra 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8400
      TabIndex        =   14
      Top             =   3600
      Width           =   2655
   End
   Begin VB.TextBox txtsalary 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8400
      TabIndex        =   13
      Top             =   2640
      Width           =   2655
   End
   Begin VB.TextBox txtnumber 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      MaxLength       =   10
      TabIndex        =   12
      Top             =   3600
      Width           =   2655
   End
   Begin VB.TextBox txtname 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      TabIndex        =   11
      Top             =   2640
      Width           =   2655
   End
   Begin VB.CommandButton cmdnext 
      BackColor       =   &H008080FF&
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
      Height          =   735
      Index           =   1
      Left            =   9600
      TabIndex        =   9
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdpre 
      BackColor       =   &H008080FF&
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
      Height          =   735
      Index           =   0
      Left            =   8040
      TabIndex        =   8
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H000000FF&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   4920
      TabIndex        =   6
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdadd 
      BackColor       =   &H008080FF&
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   3240
      TabIndex        =   4
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label lblid 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   10
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackColor       =   &H008080FF&
      Caption         =   "Employee Id"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   1080
      TabIndex        =   7
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackColor       =   &H008080FF&
      Caption         =   "Hra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   6240
      TabIndex        =   5
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H008080FF&
      Caption         =   "Salary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   6240
      TabIndex        =   3
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H008080FF&
      Caption         =   "Mobile number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H008080FF&
      Caption         =   "Employee name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   1080
      TabIndex        =   1
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "ADD NEW EMPLOYEE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "addemployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rec As New ADODB.Recordset
Private Sub cmdadd_Click(Index As Integer)
rec.AddNew
End Sub

Private Sub cmdfirst_Click(Index As Integer)
rec.MoveFirst
End Sub

Private Sub cmdlast_Click(Index As Integer)
rec.MoveLast
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
If txtname.Text = "" Or txtnumber.Text = "" Or txtsalary.Text = "" Or txthra.Text = "" Then
MsgBox " fill the compliet records please"
Else
rec.Update
MsgBox "Record is saved"
End If
End Sub

Private Sub Command1_Click()
home.Show
Me.Hide
End Sub

Private Sub Form_Load()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=A:\giftshop\Database2.mdb;Persist Security Info=False"
rec.CursorLocation = adUseClient
rec.Open "Select * from Employee", con, adOpenDynamic, adLockOptimistic
Set lblid.DataSource = rec
lblid.DataField = "ID"
Set txtname.DataSource = rec
txtname.DataField = "name"
Set txtnumber.DataSource = rec
txtnumber.DataField = "number"
Set txtsalary.DataSource = rec
txtsalary.DataField = "salary"
Set txthra.DataSource = rec
txthra.DataField = "hra"
End Sub

