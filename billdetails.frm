VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form billinformation 
   BackColor       =   &H008080FF&
   Caption         =   "Bill Informatio"
   ClientHeight    =   7785
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12660
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   12660
   StartUpPosition =   3  'Windows Default
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   4575
      Left            =   600
      OleObjectBlob   =   "billdetails.frx":0000
      TabIndex        =   0
      Top             =   3000
      Width           =   11535
   End
   Begin VB.Label Label2 
      Caption         =   "Bill No"
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "                                    Bill Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   8175
   End
End
Attribute VB_Name = "billinformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

