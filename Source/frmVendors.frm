VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmVendors 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vendors"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5040
   ScaleWidth      =   7305
   Begin VB.TextBox txtPhone 
      Appearance      =   0  'Flat
      DataField       =   "Phone"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3240
      TabIndex        =   6
      Top             =   2865
      Width           =   3300
   End
   Begin HomeFinance.RecNavigator RecNavigator1 
      Height          =   870
      Left            =   3195
      TabIndex        =   5
      Top             =   3885
      Width           =   3525
      _extentx        =   6218
      _extenty        =   1535
   End
   Begin VB.TextBox txtVendor 
      Appearance      =   0  'Flat
      DataField       =   "Vendor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3240
      TabIndex        =   1
      Top             =   840
      Width           =   3300
   End
   Begin VB.TextBox txtAddress 
      Appearance      =   0  'Flat
      DataField       =   "Address"
      Height          =   960
      Left            =   3225
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1515
      Width           =   3285
   End
   Begin MSDataGridLib.DataGrid dGrid 
      Height          =   4560
      Left            =   255
      TabIndex        =   2
      Top             =   180
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   8043
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
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
      ColumnCount     =   4
      BeginProperty Column00 
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
      BeginProperty Column01 
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
      BeginProperty Column02 
         DataField       =   "Address"
         Caption         =   "Address"
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
         DataField       =   "Phone"
         Caption         =   "Phone"
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
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "Phone # :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3225
      TabIndex        =   7
      Top             =   2595
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Vendor Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3195
      TabIndex        =   4
      Top             =   555
      Width           =   1275
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Address :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3240
      TabIndex        =   3
      Top             =   1275
      Width           =   810
   End
End
Attribute VB_Name = "frmVendors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsAllVendor As Recordset

Private Sub Form_Load()
    RecNavigator1_BindControls
    
End Sub

Private Sub RecNavigator1_BindControls()
    Set rsAllVendor = clsVendors.AllVendors()
    
    Set RecNavigator1.DataSource = rsAllVendor
    rsAllVendor.MoveLast
    
    Set dGrid.DataSource = rsAllVendor
    
    'Bind Text Boxes
    Set txtVendor.DataSource = rsAllVendor
    Set txtAddress.DataSource = rsAllVendor
    Set txtPhone.DataSource = rsAllVendor

End Sub


Private Sub RecNavigator1_WillInsert()
    txtVendor = ""
    txtAddress = ""
    txtPhone = ""
    Set txtVendor.DataSource = Nothing
    Set txtAddress.DataSource = Nothing
    Set txtPhone.DataSource = Nothing

End Sub

Private Sub RecNavigator1_Delete()
    clsVendor.DelVendor rsAllVendor.Fields("VendorID")
End Sub

Private Sub RecNavigator1_Update(Cancel As Boolean)
    Dim Er As String
    Er = clsVendors.EditVendor(rsAllVendor.Fields("VendorID"), txtVendor, txtAddress, txtPhone)
    If Er <> "" Then
        Cancel = True
        MsgBox Er
    End If
End Sub
Private Sub RecNavigator1_Insert(Cancel As Boolean)
    Dim Er As String
    Er = clsVendors.AddVendor(txtVendor, txtAddress, txtPhone)
    If Er <> "" Then
        Cancel = True
        MsgBox Er
    End If
End Sub

Private Sub RecNavigator1_LockControls(Value As Boolean)
    txtVendor.Locked = Value
    txtAddress.Locked = Value
    txtPhone.Locked = Value
End Sub

