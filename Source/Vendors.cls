VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Vendors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function AllVendors() As Recordset
    SQLText = "SELECT * FROM Vendors"
    Set AllVendors = Cn.Execute(SQLText)
End Function
Public Function VendorsForLoanCash(LoanCash As String) As Recordset
    Dim Condition As String
    If LoanCash = "" Then
        Condition = ""
    Else
        Condition = " Where Purchases.LoanCash = '" & LoanCash & "'"
    End If
    SQLText = "SELECT DISTINCT Vendors.Vendor,Purchases.VendorID" & _
    " FROM Purchases INNER JOIN Vendors ON Purchases.VendorID = Vendors.VendorID" & Condition
    Set VendorsForLoanCash = Cn.Execute(SQLText)
    
End Function

Public Function AddVendor(Vendor As String, Address As String, Phone As String) As String
    On Error GoTo AddVendorError
    SQLText = "INSERT INTO Vendors (Vendor,Address,Phone) values('" & Vendor & "','" & Adress & "','" & Phone & "')"
    Cn.Execute SQLText
    Exit Function
AddVendorError:
    AddVendor = Err.Description
End Function

Public Function DelVendor(VendorID As String)
    SQLText = "delete from Vendors where VendorID = " & VendorID
    Cn.Execute SQLText
End Function

Public Function EditVendor(VendorID As String, Vendor As String, Address As String, Phone As String) As String
    On Error GoTo EditVendorError
    SQLText = "Update Vendors set Vendor = '" & Vendor & "', Address='" & Address & "', Phone='" & Phone & "' where VendorID = " & VendorID
    Cn.Execute SQLText
    Exit Function
EditVendorError:
    EditVendor = Err.Description
End Function
