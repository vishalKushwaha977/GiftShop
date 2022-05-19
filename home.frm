VERSION 5.00
Begin VB.MDIForm home 
   BackColor       =   &H8000000C&
   Caption         =   "Home"
   ClientHeight    =   8325
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   16080
   LinkTopic       =   "Home"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   8895
      Left            =   0
      Picture         =   "home.frx":0000
      ScaleHeight     =   8895
      ScaleWidth      =   16080
      TabIndex        =   0
      Top             =   0
      Width           =   16080
   End
   Begin VB.Menu mnuhome 
      Caption         =   "Home"
   End
   Begin VB.Menu mnubilling 
      Caption         =   "Billing"
   End
   Begin VB.Menu mnudetails 
      Caption         =   "Details"
      Begin VB.Menu mnucusdetails 
         Caption         =   "Customer details"
      End
      Begin VB.Menu mnustockdetails 
         Caption         =   "Stock details"
      End
      Begin VB.Menu mnusaldetails 
         Caption         =   "Sales details"
      End
      Begin VB.Menu mnuempdetails 
         Caption         =   "Employee details"
      End
      Begin VB.Menu mnusuppdetails 
         Caption         =   "Supplier details"
      End
   End
   Begin VB.Menu mnuemployee 
      Caption         =   "Employee"
      Begin VB.Menu mnuaddEmployee 
         Caption         =   "Add Employee"
      End
   End
   Begin VB.Menu munsupplier 
      Caption         =   "Supplier"
      Begin VB.Menu mnuaddsuppler 
         Caption         =   "Add Supplier"
      End
   End
   Begin VB.Menu mnustock 
      Caption         =   "Stock"
      Begin VB.Menu mnuaddstock 
         Caption         =   "Add Stock"
      End
   End
   Begin VB.Menu mnutransaction 
      Caption         =   "Transaction"
      Begin VB.Menu mnuBillllinfo 
         Caption         =   "Billing  info"
      End
      Begin VB.Menu mnupurchaseinfo 
         Caption         =   "Purchase info"
      End
   End
   Begin VB.Menu mnureport 
      Caption         =   "Report"
      Begin VB.Menu mnustockreport 
         Caption         =   "Stock Report"
      End
      Begin VB.Menu mnucustomerreport 
         Caption         =   "Customer Report"
      End
      Begin VB.Menu mnumaintreport 
         Caption         =   "Maintenance Report"
      End
      Begin VB.Menu mnusupplerreport 
         Caption         =   "Suppler Report"
      End
      Begin VB.Menu mnubillreport 
         Caption         =   "Bills Report"
      End
      Begin VB.Menu mnusalreport 
         Caption         =   "Sales Report"
      End
   End
   Begin VB.Menu mnulogout 
      Caption         =   "LogOut"
   End
End
Attribute VB_Name = "home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub mnuaddEmployee_Click()
addemployee.Show
End Sub

Private Sub mnuaddstock_Click()
addstock.Show
End Sub
Private Sub mnuaddsuppler_Click()
addSupplier.Show
End Sub

Private Sub mnubilling_Click()
billing.Show
End Sub

Private Sub mnuBillllinfo_Click()
billinginformation.Show
End Sub

Private Sub mnucusdetails_Click()
customerdetails.Show
End Sub

Private Sub mnuempdetails_Click()
empdetails.Show
End Sub


Private Sub mnulogout_Click()
Login.Show
home.Hide
End Sub

Private Sub mnupurchaseinfo_Click()
purchaseinformation.Show
End Sub

Private Sub mnusellinfo_Click()
selling.Show
End Sub

Private Sub mnusaldetails_Click()
Salesdetails.Show
End Sub

Private Sub mnustockdetails_Click()
stockdetails.Show
End Sub

Private Sub mnuvendetails_Click()
venders.Show
End Sub

Private Sub mnusuppdetails_Click()
Supplierdetails.Show
End Sub
