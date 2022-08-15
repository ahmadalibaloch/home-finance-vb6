VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Purchases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function PurchaseTotal(Optional Year As String, Optional Month As String, Optional Day As String, Optional VendorID As String, Optional LoanCash As String) As Recordset
    Dim Condition As String
    
    'Condition for VendorID and LoanCash
    If VendorID = "" And LoanCash = "" Then
        Condition = " AND PurID=PurID"
    ElseIf VendorID = "" And LoanCash <> "" Then
        Condition = " AND LoanCash='" & LoanCash & "'"
    ElseIf VendorID <> "" And LoanCash = "" Then
        Condition = " AND Purchases.VendorID=" & VendorID
    ElseIf VendorID <> "" And LoanCash <> "" Then
        Condition = " AND Purchases.VendorID=" & VendorID & " AND LoanCash='" & LoanCash & "'"
    End If
    
    'Build SQLText
    If Year = "" And Month = "" And Day = "" Then
        SQLText = "SELECT YEAR(PDate) as _Year,SUM(Qty*Rate) as Total FROM Purchases WHERE" & Right$(Condition, Len(Condition) - 4) & " GROUP BY YEAR(PDate)"
    ElseIf Year <> "" And Month = "" And Day = "" Then
        SQLText = "SELECT MONTH(PDate) as _Month, SUM(Qty*Rate) as Total FROM Purchases" & _
        " WHERE YEAR(PDate)=" & Year & Condition & _
        " GROUP BY MONTH(PDate)"
    ElseIf Year <> "" And Month <> "" And Day = "" Then
        SQLText = "SELECT DAY(PDate) as _Day, SUM(Qty*Rate) as Total FROM Purchases" & _
        " WHERE YEAR(PDate)=" & Year & " AND MONTH(PDate)=" & Month & Condition & _
        " GROUP BY DAY(PDate)"
    ElseIf Year <> "" And Month <> "" And Day <> "" Then
        SQLText = "SELECT Items.Item, Purchases.Qty, Units.Unit, Purchases.Rate, Purchases.Rate*Purchases.Qty AS Price, Vendors.Vendor, Purchases.LoanCash" & _
        " FROM ((Purchases INNER JOIN Items ON Purchases.ItemID = Items.ItemID) INNER JOIN Units ON Purchases.UnitID = Units.UnitID) INNER JOIN Vendors ON Purchases.VendorID = Vendors.VendorID" & _
        " WHERE YEAR(PDate)=" & Year & " AND MONTH(PDate)=" & Month & " AND DAY(PDate)=" & Day & Condition
    Else
        SQLText = ""
        Set PurchaseTotal = Nothing
        Exit Function
    End If
        
    Set PurchaseTotal = Cn.Execute(SQLText)
    
End Function

Public Function DailyPurchases(PurDate As Date) As Recordset
    SQLText = "SELECT Purchases.PurID, Purchases.PDate, Purchases.ItemID, Items.Item, Purchases.Qty, Purchases.UnitID, Units.Unit, Purchases.Rate, Purchases.Rate*Purchases.Qty AS Price, Purchases.VendorID, Vendors.Vendor, Purchases.LoanCash" & _
    " FROM ((Purchases INNER JOIN Items ON Purchases.ItemID = Items.ItemID) INNER JOIN Units ON Purchases.UnitID = Units.UnitID) INNER JOIN Vendors ON Purchases.VendorID = Vendors.VendorID" & _
    " WHERE PDate=#" & PurDate & "#"
    On Error GoTo DailyPurchasesError
    Set DailyPurchases = Cn.Execute(SQLText)
    Exit Function
DailyPurchasesError:
    MsgBox Err.Description
        
End Function

Public Function AddPurchase(PDate As Date, ItemID As String, Qty As Single, UnitID As String, Rate As Single, VendorID As String, LoanCash As String)
    SQLText = "INSERT INTO Purchases (PDate,ItemID,Qty,UnitID,Rate,VendorID,LoanCash)" & _
    " VALUES (#" & PDate & "#, " & ItemID & "," & Qty & "," & UnitID & "," & Rate & "," & VendorID & ",'" & LoanCash & "')"
    On Error GoTo AddPurchasesError
    Cn.Execute SQLText
    Exit Function
AddPurchasesError:
    MsgBox Err.Description
End Function
Public Function EditPurchase(PurID As String, PDate As Date, ItemID As String, Qty As Single, UnitID As String, Rate As Single, VendorID As String, LoanCash As String)
    SQLText = "UPDATE Purchases" & _
    " set PDate=#" & PDate & "#," & _
    " ItemID=" & ItemID & "," & _
    " Qty=" & Qty & "," & _
    " UnitID=" & UnitID & "," & _
    " Rate=" & Rate & "," & _
    " VendorID=" & VendorID & "," & _
    " LoanCash='" & LoanCash & "'" & _
    " WHERE PurID=" & PurID
    On Error GoTo EditPurchaseError
    Cn.Execute SQLText
    Exit Function
EditPurchaseError:
    MsgBox Err.Description
    
End Function
Public Function DelPurchase(PurID As String)
    SQLText = "DELETE FROM Purchases WHERE PurID=" & PurID
    On Error GoTo DelPurchaseError
    Cn.Execute SQLText
    Exit Function
DelPurchaseError:
    MsgBox Err.Description
End Function
Public Function LastItemPurchase(ItemID As String) As Recordset
    Dim LastPurID As Long
    'Select last PurID for Item
    SQLText = "SELECT MAX(PurID) as LastPurID FROM Purchases WHERE ItemID=" & ItemID
    On Error GoTo LastItemPurchaseError
    Set Rs = Cn.Execute(SQLText)
    If Not (Rs.EOF Or Rs.BOF Or IsNull(Rs.Fields("LastPurID"))) Then
        LastPurID = Rs.Fields("LastPurID")
        'select last record for last PurID
        SQLText = "SELECT Rate,UnitID,VendorID,LoanCash FROM Purchases WHERE PurID=" & LastPurID
        Set LastItemPurchase = Cn.Execute(SQLText)
    Else
        Set LastItemPurchase = Nothing
    End If
    Exit Function
LastItemPurchaseError:
    Set LastItemPurchase = Nothing
    
    
End Function
