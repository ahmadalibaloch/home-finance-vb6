VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRptItems 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Report"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7260
   ScaleWidth      =   9945
   Begin VB.CommandButton btnClose 
      Caption         =   "&Close"
      Height          =   360
      Left            =   8415
      TabIndex        =   11
      Top             =   6570
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Caption         =   "Time Spane"
      Height          =   1080
      Left            =   270
      TabIndex        =   4
      Top             =   135
      Width           =   5880
      Begin VB.CommandButton btnChoose 
         Caption         =   "&Choose"
         Height          =   360
         Left            =   4410
         TabIndex        =   5
         Top             =   240
         Width           =   1245
      End
      Begin MSDataListLib.DataCombo dcEnd 
         Height          =   315
         Left            =   2640
         TabIndex        =   6
         Top             =   270
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "PDate"
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo dcStart 
         Height          =   315
         Left            =   705
         TabIndex        =   7
         Top             =   255
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "PDate"
         Text            =   "DataCombo1"
      End
      Begin VB.Label lblTotDays 
         AutoSize        =   -1  'True
         Caption         =   "Label3"
         Height          =   195
         Left            =   210
         TabIndex        =   10
         Top             =   720
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "From :"
         Height          =   180
         Left            =   195
         TabIndex        =   9
         Top             =   300
         Width           =   645
      End
      Begin VB.Label Label2 
         Caption         =   "To :"
         Height          =   180
         Left            =   2280
         TabIndex        =   8
         Top             =   315
         Width           =   375
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4560
      Left            =   285
      TabIndex        =   0
      Top             =   1320
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   8043
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
      ColumnCount     =   9
      BeginProperty Column00 
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
      BeginProperty Column02 
         DataField       =   "NoOfPur"
         Caption         =   "No.of Pur"
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
         DataField       =   "TotConsm"
         Caption         =   "Tot Consm"
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
         DataField       =   "MonthConsm"
         Caption         =   "Month Consm"
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
         DataField       =   "DailConsm"
         Caption         =   "Daily Consm"
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
         DataField       =   "TotExp"
         Caption         =   "Tot Exp"
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
         DataField       =   "MonthExp"
         Caption         =   "Month Exp"
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
         DataField       =   "DailExp"
         Caption         =   "Daily Exp"
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
            ColumnWidth     =   599.811
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1695.118
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   900.284
         EndProperty
      EndProperty
   End
   Begin VB.Label lblDlyAvg 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   195
      Left            =   315
      TabIndex        =   3
      Top             =   6600
      Width           =   480
   End
   Begin VB.Label lblMonAvg 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   195
      Left            =   315
      TabIndex        =   2
      Top             =   6330
      Width           =   480
   End
   Begin VB.Label lblTotExp 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   195
      Left            =   315
      TabIndex        =   1
      Top             =   6060
      Width           =   480
   End
End
Attribute VB_Name = "frmRptItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsAllDates As Recordset
Public rsPeriodicUsage As Recordset
Public DBStartDate, DBEndDate As Date
Dim mFirstDate, mLastDate As Date
Dim CanSetNewDate As Boolean



Private Sub DataGrid1_DblClick()
     frmItemPurDetail.Show vbModal
End Sub

Private Sub dcEnd_Change()
    If CanSetNewDate Then
        mLastDate = dcEnd.Text
        PopDGrid
    End If
End Sub

Private Sub dcStart_Change()
    If CanSetNewDate Then
        mFirstDate = dcStart.Text
        PopDGrid
    End If
End Sub

Private Sub Form_Load()
    Set rsAllDates = clsItems.AllDates()
    PopDCombs
    SetDates
    SetDComText
    CanSetNewDate = True
    PopDGrid
    
End Sub
Private Sub PopDGrid()
    Dim TotDays As Integer
    Dim Months, Days As Integer
    Dim TotExp As Long
    Set rsPeriodicUsage = clsItems.PeriodicUsage(CStr(mFirstDate), CStr(mLastDate))
    
    rsPeriodicUsage.MoveFirst
    While Not rsPeriodicUsage.EOF
        TotExp = TotExp + rsPeriodicUsage.Fields("TotExp").Value
        rsPeriodicUsage.MoveNext
    Wend
    rsPeriodicUsage.MoveFirst
    
    Set DataGrid1.DataSource = rsPeriodicUsage

    TotDays = DateDiff("d", mFirstDate, mLastDate) + 1
    Months = Fix(TotDays / 30)
    Days = TotDays Mod 30
    
    lblTotDays.Caption = TotDays & " Day(s) = " & Months & " Month(s) and " & Days & " Day(s)."
    lblTotExp.Caption = "Total Expenses for " & TotDays & " Day(s) = " & TotExp & " Rupees."
    lblMonAvg.Caption = "Monthly Average = " & Format(TotExp / (TotDays / 30), "#0.00") & " Rupees / month."
    lblDlyAvg.Caption = "Daily Average = " & Format(TotExp / TotDays, "#0.00") & " Rupees / day."
End Sub




Private Sub PopDCombs()
    Set dcStart.RowSource = rsAllDates
    Set dcEnd.RowSource = rsAllDates
    Set dcStart.DataSource = Nothing
    Set dcEnd.DataSource = Nothing
End Sub

Private Sub SetDates()
    DBStartDate = rsAllDates.Fields(0).Value
    rsAllDates.MoveLast
    DBEndDate = rsAllDates.Fields(0).Value
    mFirstDate = DBStartDate
    mLastDate = DBEndDate

End Sub
Private Sub SetDComText()
    dcStart.Text = mFirstDate
    dcEnd.Text = mLastDate

End Sub

Public Property Get FirstDate() As Date
    FirstDate = mFirstDate
End Property

Public Property Let FirstDate(ByVal Value As Date)
    If Value < DBStartDate Then
        mFirstDate = DBStartDate
    Else
        mFirstDate = Value
    End If
    
    dcStart.Text = mFirstDate
End Property

Public Property Get LastDate() As Date
    LastDate = mLastDate
End Property

Public Property Let LastDate(ByVal Value As Date)
    If Value > DBEndDate Then
        mLastDate = DBEndDate
    Else
        mLastDate = Value
    End If

    dcEnd.Text = mLastDate
End Property
Private Sub btnChoose_Click()
    frmDateChooser.Show vbModal
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub
