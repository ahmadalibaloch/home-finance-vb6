VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmItems 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Items"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4965
   ScaleWidth      =   7245
   Begin HomeFinance.RecNavigator RecNavigator1 
      Height          =   885
      Left            =   3150
      TabIndex        =   5
      Top             =   3885
      Width           =   3585
      _ExtentX        =   6324
      _ExtentY        =   1561
   End
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      DataField       =   "Des"
      Height          =   960
      Left            =   3240
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1965
      Width           =   3285
   End
   Begin VB.TextBox txtItem 
      Appearance      =   0  'Flat
      DataField       =   "Item"
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
      TabIndex        =   0
      Top             =   1305
      Width           =   3300
   End
   Begin MSDataGridLib.DataGrid dGrid 
      Height          =   4560
      Left            =   240
      TabIndex        =   2
      Top             =   240
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "ItemID"
         Caption         =   "Item ID"
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
      SplitCount      =   1
      BeginProperty Split0 
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
            ColumnAllowSizing=   0   'False
            Object.Visible         =   0   'False
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column01 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Description :"
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
      Left            =   3240
      TabIndex        =   4
      Top             =   1725
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Item Name :"
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
      Left            =   3195
      TabIndex        =   3
      Top             =   1005
      Width           =   1080
   End
End
Attribute VB_Name = "frmItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsAllItems As Recordset
Attribute rsAllItems.VB_VarHelpID = -1

Private Sub Form_Load()
    RecNavigator1_BindControls
    
End Sub

Private Sub RecNavigator1_BindControls()
    Set rsAllItems = clsItems.AllItems()
    
    Set RecNavigator1.DataSource = rsAllItems
    rsAllItems.MoveLast
    
    Set dGrid.DataSource = rsAllItems
    
    'Bind Text Boxes
    Set txtItem.DataSource = rsAllItems
    Set txtDescription.DataSource = rsAllItems

End Sub


Private Sub RecNavigator1_WillInsert()
    txtItem = ""
    txtDescription = ""
    Set txtItem.DataSource = Nothing
    Set txtDescription.DataSource = Nothing

End Sub

Private Sub RecNavigator1_Delete()
    clsItems.DelItem rsAllItems.Fields("ItemID")
End Sub

Private Sub RecNavigator1_Update(Cancel As Boolean)
    Dim Er As String
    Er = clsItems.EditItem(rsAllItems.Fields("ItemID"), txtItem, txtDescription)
    If Er <> "" Then
        Cancel = True
        MsgBox Er
    End If
End Sub
Private Sub RecNavigator1_Insert(Cancel As Boolean)
    Dim Er As String
    Er = clsItems.AddItem(txtItem, txtDescription)
    If Er <> "" Then
        Cancel = True
        MsgBox Er
    End If
End Sub

Private Sub RecNavigator1_LockControls(Value As Boolean)
    txtItem.Locked = Value
    txtDescription.Locked = Value
End Sub

