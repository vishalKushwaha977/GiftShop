VERSION 5.00
Begin VB.Form Login 
   BackColor       =   &H80000016&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   6420
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   10560
   ForeColor       =   &H00FF80FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3793.148
   ScaleMode       =   0  'User
   ScaleWidth      =   9915.269
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "register"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   8
      Top             =   5400
      Width           =   4095
   End
   Begin VB.TextBox txtusername 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   7440
      TabIndex        =   1
      Top             =   2640
      Width           =   2685
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "login"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   5400
      TabIndex        =   4
      Top             =   4560
      Width           =   1740
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   5
      Top             =   4560
      Width           =   1740
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      IMEMode         =   3  'DISABLE
      Left            =   7440
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   3480
      Width           =   2685
   End
   Begin VB.Image Image1 
      Height          =   3645
      Left            =   840
      Picture         =   "frmLogin.frx":0000
      Top             =   1680
      Width           =   2940
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "User Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   5040
      TabIndex        =   7
      Top             =   1680
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "      Gift shop management system"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   975
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   10335
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&User Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   5040
      TabIndex        =   0
      Top             =   2640
      Width           =   2400
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   5040
      TabIndex        =   2
      Top             =   3480
      Width           =   2400
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
If txtusername.Text = "vishal" Or txtusername.Text = "neelesh" Then
    'check for correct password
    If txtpassword = "12345" Or txtpassword.Text = "123456" Then
        'place code to here to pass the
        'success to the calling sub
        'setting a global var is the easiest
        LoginSucceeded = True
        home.Show
        Me.Hide
        txtpassword.Text = ""
        txtusername.Text = ""
    Else
        MsgBox "Invalid Password, try again!", vbExclamation, "Login"
        txtpassword.SetFocus
    End If
    Else
        MsgBox "Invalid Username"
        txtpassword.Text = ""
        txtusername.Text = ""
    End If
End Sub

Private Sub Command1_Click()
register.Show
Me.Hide
End Sub
