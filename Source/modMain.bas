Attribute VB_Name = "modMain"
'application varialbe
Dim myFrmMain As frmMain

'sql variables
Public Cn As Connection
Public rs As Recordset
Public SQLText As String

'class variables
Public clsPurchases As Purchases
Public clsItems As Items
Public clsVendors As Vendors
Public clsUnits As Units

'Enum variables
Public Enum SaveMode
    Adding
    Editing
    None
End Enum

Sub Main()
    ' make connection to the database
    Set Cn = New Connection
    Cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\homeFinance.mdb;Persist Security Info=False"
    Cn.CursorLocation = adUseClient
    Cn.Open
    
    'start new class instances
    Set clsPurchases = New Purchases
    Set clsItems = New Items
    Set clsVendors = New Vendors
    Set clsUnits = New Units
    
    'Start Application
    Set myFrmMain = New frmMain
    myFrmMain.Show
End Sub







