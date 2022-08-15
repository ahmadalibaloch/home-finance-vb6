VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRptPurchases 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchases Report"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6735
   ScaleWidth      =   10350
   Begin VB.CommandButton btnRefresh 
      Caption         =   "&Refresh"
      Height          =   315
      Left            =   2055
      TabIndex        =   13
      Top             =   2310
      Width           =   1110
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   2400
      Left            =   165
      TabIndex        =   7
      Top             =   375
      Width           =   3255
      Begin VB.CheckBox chkVendor 
         Appearance      =   0  'Flat
         Caption         =   "Vendor"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   180
         TabIndex        =   12
         Top             =   1515
         Width           =   930
      End
      Begin VB.OptionButton optLoanCash 
         Appearance      =   0  'Flat
         Caption         =   "Cash + Loan"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   150
         TabIndex        =   10
         Top             =   1080
         Value           =   -1  'True
         Width           =   1260
      End
      Begin VB.OptionButton optCash 
         Appearance      =   0  'Flat
         Caption         =   "Cash"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   150
         TabIndex        =   9
         Top             =   705
         Width           =   915
      End
      Begin VB.OptionButton optLoan 
         Appearance      =   0  'Flat
         Caption         =   "Loan"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   150
         TabIndex        =   8
         Top             =   330
         Width           =   930
      End
      Begin MSDataListLib.DataCombo dcomVendor 
         DataField       =   "VendorID"
         Height          =   315
         Left            =   1200
         TabIndex        =   11
         Top             =   1500
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Style           =   2
         ListField       =   "Vendor"
         BoundColumn     =   "VendorID"
         Text            =   ""
      End
   End
   Begin MSDataGridLib.DataGrid DGridDay 
      Height          =   6090
      Left            =   7965
      TabIndex        =   2
      Top             =   450
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   10742
      _Version        =   393216
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "_Day"
         Caption         =   "Date"
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
         DataField       =   "Total"
         Caption         =   "Total"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   599.811
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   929.764
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGridMonth 
      Height          =   2685
      Left            =   5625
      TabIndex        =   1
      Top             =   465
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   4736
      _Version        =   393216
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "_Month"
         Caption         =   "Month"
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
         DataField       =   "Total"
         Caption         =   "Total"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   599.811
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   929.764
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGridYear 
      Height          =   2685
      Left            =   3750
      TabIndex        =   0
      Top             =   465
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   4736
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "_Year"
         Caption         =   "Year"
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
         DataField       =   "Total"
         Caption         =   "Total"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   599.811
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   929.764
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGridPurchase 
      Height          =   3240
      Left            =   165
      TabIndex        =   3
      Top             =   3285
      Width           =   7620
      _ExtentX        =   13441
      _ExtentY        =   5715
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
      ColumnCount     =   7
      BeginProperty Column00 
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
      BeginProperty Column01 
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
      BeginProperty Column02 
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
      BeginProperty Column03 
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
      BeginProperty Column04 
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
      BeginProperty Column05 
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
      BeginProperty Column06 
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
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   794.835
         EndProperty
      EndProperty
   End
   Begin VB.Label lblDay 
      AutoSize        =   -1  'True
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   510
      TabIndex        =   6
      Top             =   3030
      Width           =   120
   End
   Begin VB.Label lblMonth 
      AutoSize        =   -1  'True
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   8340
      TabIndex        =   5
      Top             =   195
      Width           =   120
   End
   Begin VB.Label lblYear 
      AutoSize        =   -1  'True
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   5970
      TabIndex        =   4
      Top             =   210
      Width           =   120
   End
End
Attribute VB_Name = "frmRptPurchases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents rsYearlyTotal As Recordset
Attribute rsYearlyTotal.VB_VarHelpID = -1
Dim WithEvents rsMonthlyTotal As Recordset
Attribute rsMonthlyTotal.VB_VarHelpID = -1
Dim WithEvents rsDayTotal As Recordset
Attribute rsDayTotal.VB_VarHelpID = -1
Dim WithEvents rsPurchase As Recordset
Attribute rsPurchase.VB_VarHelpID = -1
Dim rsVendors As Recordset
Private Sub Form_Load()
    'Fill dcomVendor
    PopDcomVendors
    
    ' Step by Step sequence
    BiginSequence
End Sub

Private Sub BiginSequence()
    'Fetch YearlyTotal
    Set rsYearlyTotal = clsPurchases.PurchaseTotal(, , , VendorID, LoanCash)
    rsYearlyTotal.MoveLast
    
    ' Bind DGridYear
    Set DGridYear.DataSource = rsYearlyTotal

End Sub
Private Sub PopDcomVendors()
    Set rsVendors = clsVendors.VendorsForLoanCash(LoanCash)
    Set dcomVendor.RowSource = rsVendors
    Set dcomVendor.DataSource = Nothing
    If Not (rsVendors.EOF Or rsVendors.BOF) Then
        dcomVendor.BoundText = rsVendors.Fields("VendorID")
    End If
End Sub
Private Sub rsYearlyTotal_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    'Write lblYear
    lblYear.Caption = rsYearlyTotal.Fields("_Year")
    
    'Fetch MonthlyTotal
    Set rsMonthlyTotal = clsPurchases.PurchaseTotal(rsYearlyTotal.Fields("_Year"), , , VendorID, LoanCash)
    rsMonthlyTotal.MoveLast
    
    'Bind DGridMonth
    Set DGridMonth.DataSource = rsMonthlyTotal
    
End Sub
Private Sub rsMonthlyTotal_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    'Write lblMonth
    lblMonth.Caption = FindMonth(rsMonthlyTotal.Fields("_Month"))
    'Fetch DayTotal
    Set rsDayTotal = clsPurchases.PurchaseTotal(rsYearlyTotal.Fields("_Year"), rsMonthlyTotal.Fields("_Month"), , VendorID, LoanCash)
    rsDayTotal.MoveLast
    
    'Bind DGridDay
    Set DGridDay.DataSource = rsDayTotal

End Sub

Private Sub rsDayTotal_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    'Write lblDay
    lblDay.Caption = FindDate(rsYearlyTotal.Fields("_Year"), rsMonthlyTotal.Fields("_Month"), rsDayTotal.Fields("_Day"))
    'Fetch Purchase for the Day
    Set rsPurchase = clsPurchases.PurchaseTotal(rsYearlyTotal.Fields("_Year"), rsMonthlyTotal.Fields("_Month"), rsDayTotal.Fields("_Day"), VendorID, LoanCash)
    
    'Bind DGridPurchase
    Set DGridPurchase.DataSource = rsPurchase
End Sub
Private Function VendorID() As String
    If dcomVendor.Enabled Then
        VendorID = dcomVendor.BoundText
    Else
        VendorID = ""
    End If
End Function
Private Function LoanCash() As String
    If optLoan.Value Then
        LoanCash = "Loan"
    ElseIf optCash.Value Then
        LoanCash = "Cash"
    Else
        LoanCash = ""
    End If
End Function
Private Sub btnRefresh_Click()
    BiginSequence
End Sub

Private Sub chkVendor_Click()
    dcomVendor.Enabled = chkVendor.Value
End Sub
Private Sub optCash_Click()
    PopDcomVendors
End Sub

Private Sub optLoan_Click()
    PopDcomVendors
End Sub

Private Sub optLoanCash_Click()
    PopDcomVendors
End Sub

Private Function FindMonth(Month As Integer) As String
    FindMonth = Format(DateSerial(2000, Month, 1), "mmmm")
End Function

Private Function FindDate(Year As Integer, Month As Integer, Day As Integer) As String
    FindDate = Format(DateSerial(Year, Month, Day), "dd, mmmm, yyyy")
End Function

