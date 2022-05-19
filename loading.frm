VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form loading 
   BackColor       =   &H00C0C0C0&
   Caption         =   "loading"
   ClientHeight    =   6660
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   ScaleHeight     =   6660
   ScaleWidth      =   10605
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   1095
      Left            =   960
      TabIndex        =   2
      Top             =   5400
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   1931
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.PictureBox Picture1 
      Height          =   4335
      Left            =   0
      Picture         =   "loading.frx":0000
      ScaleHeight     =   4275
      ScaleWidth      =   6315
      TabIndex        =   1
      Top             =   720
      Width           =   6375
   End
   Begin VB.Timer Timer1 
      Interval        =   60
      Left            =   120
      Top             =   5160
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "   Gift shop management        system"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   2895
      Left            =   6960
      TabIndex        =   0
      Top             =   1440
      Width           =   3135
   End
End
Attribute VB_Name = "loading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
If ProgressBar1.Value = 100 Then
  Timer1.Enabled = False
  Me.Hide
  Load Login
  Login.Show
Else
  ProgressBar1.Value = ProgressBar1.Value + 10
  Label1.Caption = "Loading... " & ProgressBar1.Value & "% Completed"
End If
End Sub

