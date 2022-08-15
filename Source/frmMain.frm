VERSION 5.00
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Home Finance"
   ClientHeight    =   4410
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7050
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu fileOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu FileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu FileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuPurchase 
      Caption         =   "&Purchase"
      Begin VB.Menu PurchaseAdd 
         Caption         =   "&Add Purchase"
      End
      Begin VB.Menu PurchaseEdit 
         Caption         =   "&Edit Purchase"
      End
      Begin VB.Menu PurchaseDelete 
         Caption         =   "&Delete Purchase"
      End
   End
   Begin VB.Menu mnuItem 
      Caption         =   "&Item"
      Begin VB.Menu ItemAdd 
         Caption         =   "&Add Item"
      End
      Begin VB.Menu ItemEdit 
         Caption         =   "&Edit Item"
      End
      Begin VB.Menu ItemDelete 
         Caption         =   "&Delete Item"
      End
   End
   Begin VB.Menu mnuVendor 
      Caption         =   "&Vendor"
      Begin VB.Menu VendorAdd 
         Caption         =   "&Add Vendor"
      End
      Begin VB.Menu VendorEdit 
         Caption         =   "&Edit Vendor"
      End
      Begin VB.Menu VendorDelete 
         Caption         =   "&Delete Vendor"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu ReportsPurchases 
         Caption         =   "&Purchases"
      End
      Begin VB.Menu ReportsItems 
         Caption         =   "&Items"
      End
      Begin VB.Menu ReportsVendors 
         Caption         =   "&Vendors"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MDIForm_Load()
    PurchaseDelete.Enabled = False
    ItemDelete.Enabled = False
End Sub

Private Sub ItemAdd_Click()
    Load frmItems
End Sub

Private Sub PurchaseAdd_Click()
    Load frmPurchases
    frmPurchases.btnAddNew_Click
End Sub

Private Sub PurchaseEdit_Click()
    Load frmPurchases
    frmPurchases.btnEdit_Click
End Sub

Private Sub ReportsItems_Click()
    frmRptItems.Show
End Sub

Private Sub ReportsPurchases_Click()
    frmRptPurchases.Show
End Sub

Private Sub VendorAdd_Click()
    Load frmVendors
End Sub
