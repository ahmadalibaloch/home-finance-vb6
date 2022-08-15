VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Temp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function Conversion()
    Dim Rs2 As Recordset
    
    Dim PDate As Date
    Dim ItemID As String
    Dim Qty As String
    Dim UnitID As String
    Dim Rate As String
    Dim VendorID As String
    Dim LoanCash As String
    
    
    SQLText = "Select * from temp"
    Set Rs = Cn.Execute(SQLText)
    Rs.MoveFirst
    
    While Not Rs.EOF
        PDate = Rs.Fields("Date")
        
        SQLText = "select ItemID from Items where Item='" & Rs.Fields("Name") & "'"
        Set Rs2 = Cn.Execute(SQLText)
        ItemID = Rs2.Fields("ItemID")
        
        Qty = Format(Rs.Fields("Qty"), "#.00")
        
        SQLText = "select UnitID from Units where Unit='" & Rs.Fields("Pkt/Kg") & "'"
        Set Rs2 = Cn.Execute(SQLText)
        UnitID = Rs2.Fields("UnitID")
        
        Rate = Rs.Fields("Price")
        
        SQLText = "select VendorID from Vendors where Vendor='" & Rs.Fields("From") & "'"
        Set Rs2 = Cn.Execute(SQLText)
        VendorID = Rs2.Fields("VendorID")
        
        LoanCash = Rs.Fields("As")
        
        clsPurchases.AddPurchase PDate, ItemID, (Qty), UnitID, (Rate), VendorID, LoanCash
        
        Rs.MoveNext
    Wend
End Function
