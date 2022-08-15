VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmPurchases 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6315
   ScaleWidth      =   8610
   Begin VB.TextBox txtTotal 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   7065
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   5760
      Width           =   1305
   End
   Begin VB.CommandButton btnEdit 
      Caption         =   "&Edit"
      Height          =   360
      Left            =   1815
      TabIndex        =   8
      Top             =   5775
      Width           =   1245
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   1815
      TabIndex        =   22
      Top             =   5775
      Width           =   1245
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "&Save"
      Height          =   360
      Left            =   390
      TabIndex        =   7
      Top             =   5760
      Width           =   1245
   End
   Begin VB.CommandButton btnDelete 
      Caption         =   "&Delete"
      Height          =   360
      Left            =   3330
      TabIndex        =   9
      Top             =   5775
      Width           =   1245
   End
   Begin VB.CommandButton btnAddNew 
      Caption         =   "Add &New"
      Height          =   360
      Left            =   390
      TabIndex        =   21
      Top             =   5760
      Width           =   1245
   End
   Begin VB.Frame Frame3 
      Caption         =   " Item "
      Height          =   2340
      Left            =   180
      TabIndex        =   16
      Top             =   735
      Width           =   4215
      Begin VB.TextBox txtQty 
         Appearance      =   0  'Flat
         DataField       =   "Qty"
         Height          =   285
         Left            =   1215
         TabIndex        =   1
         Top             =   810
         Width           =   1335
      End
      Begin VB.TextBox txtPrice 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1185
         TabIndex        =   4
         Top             =   1860
         Width           =   1335
      End
      Begin VB.TextBox txtRate 
         Appearance      =   0  'Flat
         DataField       =   "Rate"
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   1335
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo dcomUnits 
         DataField       =   "UnitID"
         Height          =   315
         Left            =   2655
         TabIndex        =   2
         Top             =   810
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Appearance      =   0
         Style           =   2
         ListField       =   "Unit"
         BoundColumn     =   "UnitID"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcomItems 
         DataField       =   "ItemID"
         Height          =   315
         Left            =   1215
         TabIndex        =   0
         Top             =   300
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Appearance      =   0
         Style           =   2
         ListField       =   "Item"
         BoundColumn     =   "ItemID"
         Text            =   ""
      End
      Begin VB.Label Label1 
         Caption         =   "Item :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   135
         TabIndex        =   20
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Quantity :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   135
         TabIndex        =   19
         Top             =   810
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Rate :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   135
         TabIndex        =   18
         Top             =   1335
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Price :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   135
         TabIndex        =   17
         Top             =   1860
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " From "
      Height          =   2340
      Left            =   4515
      TabIndex        =   13
      Top             =   735
      Width           =   3885
      Begin VB.Frame Frame1 
         Caption         =   " Loan / Cash "
         Height          =   855
         Left            =   990
         TabIndex        =   11
         Top             =   1035
         Width           =   2070
         Begin VB.OptionButton optLoan 
            Appearance      =   0  'Flat
            Caption         =   "Loan"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   360
            Width           =   735
         End
         Begin VB.OptionButton optCash 
            Appearance      =   0  'Flat
            Caption         =   "Cash"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1230
            TabIndex        =   6
            Top             =   300
            Width           =   705
         End
      End
      Begin MSDataListLib.DataCombo dcomVendors 
         DataField       =   "VendorID"
         Height          =   315
         Left            =   1305
         TabIndex        =   5
         Top             =   375
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Appearance      =   0
         Style           =   2
         ListField       =   "Vendor"
         BoundColumn     =   "VendorID"
         Text            =   ""
      End
      Begin VB.Label Label5 
         Caption         =   "Vendor :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   285
         TabIndex        =   15
         Top             =   405
         Width           =   975
      End
   End
   Begin MSDataGridLib.DataGrid DGrid 
      Height          =   2400
      Left            =   195
      TabIndex        =   12
      Top             =   3210
      Width           =   8190
      _ExtentX        =   14446
      _ExtentY        =   4233
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   12
      BeginProperty Column00 
         DataField       =   "PurID"
         Caption         =   "PurID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "PDate"
         Caption         =   "PDate"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "ItemID"
         Caption         =   "ItemID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Item"
         Caption         =   "Item"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Qty"
         Caption         =   "Qty"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "UnitID"
         Caption         =   "UnitID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Unit"
         Caption         =   "Unit"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "Rate"
         Caption         =   "Rate"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "Price"
         Caption         =   "Price"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00;(0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "VendorID"
         Caption         =   "VendorID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "Vendor"
         Caption         =   "Vendor"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "LoanCash"
         Caption         =   "LoanCash"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Object.Visible         =   0   'False
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column09 
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   794.835
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   315
      TabIndex        =   10
      Top             =   240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   20381697
      CurrentDate     =   39112
   End
   Begin VB.Label lblTotItems 
      Caption         =   "Items = 0"
      Height          =   225
      Left            =   5040
      TabIndex        =   26
      Top             =   5805
      Width           =   1245
   End
   Begin VB.Label lblSaveInfo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "lblSaveInfo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   7170
      TabIndex        =   25
      Top             =   345
      Width           =   1185
   End
   Begin VB.Label Label6 
      Caption         =   "Total = "
      Height          =   255
      Left            =   6495
      TabIndex        =   24
      Top             =   5805
      Width           =   585
   End
End
Attribute VB_Name = "frmPurchases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents rsDailyPurchases As Recordset
Attribute rsDailyPurchases.VB_VarHelpID = -1
Dim rsAllItems As Recordset
Dim rsAllUnits As Recordset
Dim rsAllVendors As Recordset
Dim rsLastItemPurchase As Recordset
Dim mySaveMode As SaveMode

Private Sub Form_Load()
    DTPicker1.Value = Date
    mySaveMode = None
    BindPopControls DTPicker1.Value
    ShowSave False
    Lock_Controls True
        
End Sub

Public Sub btnAddNew_Click()
    txtQty = ""
    Set txtQty.DataSource = Nothing
    Set txtRate.DataSource = Nothing
    mySaveMode = Adding
    Lock_Controls False
    ShowSave True
    lblSaveInfo.Caption = "Add New"
    dcomItems.SetFocus
End Sub

Private Sub btnCancel_Click()
    mySaveMode = None
    BindPopControls DTPicker1.Value
    ShowSave False
    Lock_Controls True
   lblSaveInfo.Caption = ""

End Sub

Public Sub btnDelete_Click()
    Dim Yes As VbMsgBoxResult
    Yes = vbYes
    If Not (rsDailyPurchases.EOF Or rsDailyPurchases.BOF) Then
        If Yes = MsgBox("Record will be deleted Permanently. Sure?", vbYesNo) Then
            clsPurchases.DelPurchase rsDailyPurchases.Fields("PurID")
            BindPopControls DTPicker1.Value
        End If
    End If
End Sub

Public Sub btnEdit_Click()
    mySaveMode = Editing
    ShowSave True
    Lock_Controls False
    lblSaveInfo.Caption = "Edit"
End Sub

Private Sub btnSave_Click()
    On Error GoTo SaveError
    If mySaveMode = Adding Then
        clsPurchases.AddPurchase DTPicker1.Value, dcomItems.BoundText, txtQty, dcomUnits.BoundText, txtRate, dcomVendors.BoundText, TelLoanCash
    ElseIf mySaveMode = Editing And Not (rsDailyPurchases.EOF Or rsDailyPurchases.BOF) Then
        clsPurchases.EditPurchase rsDailyPurchases.Fields("PurID"), DTPicker1.Value, dcomItems.BoundText, txtQty, dcomUnits.BoundText, txtRate, dcomVendors.BoundText, TelLoanCash
    End If
    mySaveMode = None
    ShowSave False
    Lock_Controls True
    BindPopControls DTPicker1.Value
    lblSaveInfo.Caption = ""
    btnAddNew.SetFocus
    Exit Sub
SaveError:
    MsgBox Err.Description & ". Some Fields may be Empty."

End Sub


Private Sub dcomItems_Change()
    On Error GoTo ClearControls
    If mySaveMode = Adding Then
        Set rsLastItemPurchase = clsPurchases.LastItemPurchase(dcomItems.BoundText)
        If Not (rsLastItemPurchase.EOF Or rsLastItemPurchase.BOF) Then
            txtQty = ""
            txtRate = rsLastItemPurchase.Fields("Rate")
            dcomUnits.BoundText = rsLastItemPurchase.Fields("UnitID")
            dcomVendors.BoundText = rsLastItemPurchase.Fields("VendorID")
            SetLoanCash rsLastItemPurchase.Fields("LoanCash")
        End If
    End If
    Exit Sub
ClearControls:
    txtRate = ""
    dcomUnits.BoundText = ""
    dcomVendors.BoundText = ""
    optLoan.Value = False
    optCash.Value = False
End Sub

Private Sub BindPopControls(pramDate As Date)
    'Fetch DailyPurchases Recordset
    Set rsDailyPurchases = clsPurchases.DailyPurchases(pramDate)
    
    'Move to the last Record and Calculate Total
    Dim Total As Single
    txtTotal = ""
    If Not (rsDailyPurchases.EOF Or rsDailyPurchases.BOF) Then
        While Not rsDailyPurchases.EOF
            Total = Total + rsDailyPurchases.Fields("Price")
            rsDailyPurchases.MoveNext
        Wend
        txtTotal = Format(Total, "#.00")
        lblTotItems.Caption = "Items = " & rsDailyPurchases.RecordCount

    End If
    
    'Bind DGrid
    Set DGrid.DataSource = rsDailyPurchases

    'Fetch other Recordsets
    Set rsAllItems = clsItems.AllItems()
    Set rsAllUnits = clsUnits.AllUnits()
    Set rsAllVendors = clsVendors.AllVendors()
    
    'Bind ComboBoxes
    Set dcomItems.RowSource = rsAllItems
    Set dcomUnits.RowSource = rsAllUnits
    Set dcomVendors.RowSource = rsAllVendors

    Set dcomItems.DataSource = rsDailyPurchases
    Set dcomUnits.DataSource = rsDailyPurchases
    Set dcomVendors.DataSource = rsDailyPurchases
    
    'Bind Text Boxes
    Set txtQty.DataSource = rsDailyPurchases
    Set txtRate.DataSource = rsDailyPurchases
    
    'Bind LoanCash Option Boxes
    If Not (rsDailyPurchases.EOF Or rsDailyPurchases.BOF) Then
        SetLoanCash rsDailyPurchases.Fields("LoanCash")
    End If
    lblSaveInfo.Caption = ""
    
End Sub

Private Sub SetLoanCash(LoanCash As String)
    If LoanCash = "Loan" Then
        optLoan.Value = True
    Else
        optCash.Value = True
    End If
End Sub

Private Function TelLoanCash() As String
    If optLoan.Value Then
        TelLoanCash = "Loan"
    Else
        TelLoanCash = "Cash"
    End If
End Function
Private Sub rsDailyPurchases_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If Not (rsDailyPurchases.EOF Or rsDailyPurchases.BOF) Then
        SetLoanCash rsDailyPurchases.Fields("LoanCash")
    End If
End Sub
Private Sub txtPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Calc3rdValue
    End If
End Sub
Private Sub txtRate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Calc3rdValue
    End If
End Sub
Private Sub txtQty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Calc3rdValue
    End If
End Sub

Private Sub txtQty_Change()
    txtPrice = CalcPrice
End Sub
Private Sub txtRate_Change()
    txtPrice = CalcPrice
End Sub
Private Sub Calc3rdValue()
    If Val(txtPrice) > 0 And Val(txtQty) > 0 Then
        txtRate = Format(Val(txtPrice) / Val(txtQty), "#.00")
    ElseIf Val(txtPrice) > 0 And Val(txtRate) > 0 Then
        txtQty = Format(Val(txtPrice) / Val(txtRate), "#.00")
    ElseIf Val(txtQty) > 0 And Val(txtRate) > 0 Then
        txtPrice = Format(Val(txtQty) * Val(txtRate), "#.00")
    End If
End Sub

Private Sub ShowSave(vValue As Boolean)
    btnAddNew.Visible = Not vValue
    btnEdit.Visible = Not vValue
    btnDelete.Visible = Not vValue
    btnSave.Visible = vValue
    btnCancel.Visible = vValue
End Sub
Private Sub Lock_Controls(LValue As Boolean)
    dcomItems.Locked = LValue
    dcomUnits.Locked = LValue
    dcomVendors.Locked = LValue
    txtQty.Locked = LValue
    txtRate.Locked = LValue
    txtPrice.Locked = LValue
    Frame1.Enabled = Not LValue
End Sub
Private Function CalcPrice() As Single
    CalcPrice = Format(Val(txtQty) * Val(txtRate), "#.00")
End Function
Private Sub DTPicker1_Change()
    On Error Resume Next
    BindPopControls DTPicker1.Value
    ShowSave False
    Lock_Controls True
End Sub
