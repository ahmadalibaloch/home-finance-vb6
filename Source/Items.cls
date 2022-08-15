VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Items"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function AllItems() As Recordset
    SQLText = "SELECT * FROM Items"
    Set AllItems = Cn.Execute(SQLText)
End Function

Public Function AddItem(Item As String, Des As String) As String
    On Error GoTo AddItemError
    SQLText = "INSERT INTO Items (Item,Des) values('" & Item & "','" & Des & "')"
    Cn.Execute SQLText
    Exit Function
AddItemError:
    AddItem = Err.Description
End Function

Public Function DelItem(ItemID As String)
    SQLText = "delete from Items where ItemID = " & ItemID
    Cn.Execute SQLText
End Function

Public Function EditItem(ItemID As String, Item As String, Des As String) As String
    On Error GoTo EditItemError
    SQLText = "Update items set item = '" & Item & "', Des='" & Des & "' where ItemID = " & ItemID
    Cn.Execute SQLText
    Exit Function
EditItemError:
    EditItem = Err.Description
End Function
Public Function PeriodicUsage(FromDate As String, ToDate As String) As Recordset
    Dim FirstDate, LastDate As Date
    Dim TotDays As Integer
    Dim TotMonths As Single
    FirstDate = CDate(FromDate)
    LastDate = CDate(ToDate)
    TotDays = DateDiff("d", FirstDate, LastDate) + 1
    TotMonths = TotDays / 30
    
    SQLText = "SELECT Purchases.ItemID, Items.Item, COUNT(Purchases.ItemID) AS NoOfPur" & _
    ", Format(Sum(Purchases.Qty),'#0.00') & ' ' & Units.Unit as TotConsm" & _
    ", Format(Sum(Purchases.Qty)/" & TotMonths & ",'#0.00') & ' ' & Units.Unit as MonthConsm" & _
    ", Format(Sum(Purchases.Qty)/" & TotDays & ",'#0.00') & ' ' & Units.Unit as DailConsm" & _
    ", Format(SUM(Purchases.Qty*Purchases.Rate),'#0.00') AS TotExp" & _
    ", Format(SUM(Purchases.Qty*Purchases.Rate)/" & TotMonths & ",'#0.00') AS MonthExp" & _
    ", Format(SUM(Purchases.Qty*Purchases.Rate)/" & TotDays & ",'#0.00') AS DailExp" & _
    " FROM (Purchases INNER JOIN Items ON Purchases.ItemID = Items.ItemID) INNER JOIN Units" & _
    " ON Purchases.UnitID = Units.UnitID" & _
    " WHERE PDate between #" & FromDate & "# AND #" & ToDate & "#" & _
    " GROUP BY Purchases.ItemID, Items.Item, Units.Unit" & _
    " ORDER BY Sum(Qty*rate) DESC"
    Set PeriodicUsage = Cn.Execute(SQLText)
End Function
Public Function AllDates() As Recordset
    SQLText = "SELECT DISTINCT PDate FROM Purchases ORDER BY PDate"
    Set AllDates = Cn.Execute(SQLText)
End Function
Public Function ItemPurDetail(ItemID As String, FromDate As String, ToDate As String) As Recordset
    SQLText = "SELECT PDate,Qty & ' ' & Unit as Qty,Rate,Format(Qty*Rate,'#0.00') as Price, Vendor, LoanCash" & _
    " FROM (Purchases INNER JOIN Units ON Purchases.UnitID = Units.UnitID) INNER JOIN Vendors ON Purchases.VendorID = Vendors.VendorID" & _
    " WHERE ItemID = " & ItemID & " AND PDate BETWEEN #" & FromDate & "# AND #" & ToDate & "#" & _
    " ORDER BY PDate"
    Set ItemPurDetail = Cn.Execute(SQLText)
    
End Function




