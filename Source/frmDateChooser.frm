VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDateChooser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Date Chooser"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   10815
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnOk 
      Caption         =   "&Ok"
      Height          =   360
      Left            =   9435
      TabIndex        =   1
      Top             =   6690
      Width           =   1245
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   6600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10830
      _ExtentX        =   19103
      _ExtentY        =   11642
      _Version        =   393216
      ForeColor       =   65535
      BackColor       =   -2147483633
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxSelCount     =   366
      MonthColumns    =   4
      MonthRows       =   3
      MonthBackColor  =   8421376
      MultiSelect     =   -1  'True
      ShowToday       =   0   'False
      StartOfWeek     =   20381697
      TitleBackColor  =   49344
      TrailingForeColor=   33023
      CurrentDate     =   39183
   End
End
Attribute VB_Name = "frmDateChooser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FirstDate, LastDate As Date

Private Sub btnOk_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    FirstDate = frmRptItems.DBStartDate
    LastDate = frmRptItems.DBEndDate
    MonthView1.SelStart = frmRptItems.FirstDate
    MonthView1.SelEnd = frmRptItems.LastDate
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmRptItems.FirstDate = MonthView1.SelStart
    frmRptItems.LastDate = MonthView1.SelEnd
End Sub

Private Sub MonthView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    btnOk.SetFocus
End Sub

Private Sub MonthView1_SelChange(ByVal StartDate As Date, ByVal EndDate As Date, Cancel As Boolean)
    If StartDate < FirstDate Then
        MonthView1.SelStart = FirstDate
    End If
    If EndDate > LastDate Then
        MonthView1.SelEnd = LastDate
    End If
End Sub
