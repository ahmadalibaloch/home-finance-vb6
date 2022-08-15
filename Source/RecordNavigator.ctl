VERSION 5.00
Begin VB.UserControl RecNavigator 
   ClientHeight    =   870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3510
   ScaleHeight     =   870
   ScaleWidth      =   3510
   Begin VB.CommandButton cmdAddNew 
      Caption         =   "Add &New"
      Height          =   345
      Left            =   45
      TabIndex        =   8
      Top             =   465
      Width           =   1065
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   345
      Left            =   2400
      TabIndex        =   7
      Top             =   465
      Width           =   1065
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Save"
      Height          =   345
      Left            =   45
      TabIndex        =   6
      Top             =   465
      Width           =   1065
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   1230
      TabIndex        =   5
      Top             =   465
      Width           =   1065
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   345
      Left            =   1230
      TabIndex        =   4
      Top             =   465
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1050
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   75
      Width           =   1410
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   ">>"
      Height          =   315
      Left            =   2985
      TabIndex        =   3
      Top             =   60
      Width           =   465
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      Height          =   315
      Left            =   2520
      TabIndex        =   2
      Top             =   60
      Width           =   465
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<"
      Height          =   315
      Left            =   510
      TabIndex        =   1
      Top             =   60
      Width           =   465
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "<<"
      Height          =   315
      Left            =   45
      TabIndex        =   0
      Top             =   60
      Width           =   465
   End
End
Attribute VB_Name = "RecNavigator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents DataSource As ADODB.Recordset
Attribute DataSource.VB_VarHelpID = -1
Private ForInsert As Boolean
Private myBookMark As Variant
Event Update(ByRef Cancel As Boolean)
Event Insert(ByRef Cancel As Boolean)
Event Delete()
Event LockControls(Value As Boolean)
Event WillInsert()
Event BindControls()
Private Sub UserControl_Initialize()
    ShowSave False
    RaiseEvent BindControls
End Sub
Private Sub cmdAddNew_Click()
    RaiseEvent WillInsert
    ForInsert = True
    EnableScroll False
    ShowSave True
    RaiseEvent LockControls(False)
    Text1.ForeColor = vbRed
    Text1 = "Add New"
    cmdCancel.SetFocus
End Sub

Private Sub cmdEdit_Click()
    ForInsert = False
    RaiseEvent LockControls(False)
    myBookMark = DataSource.Bookmark
    EnableScroll False
    ShowSave True
    Text1.ForeColor = vbRed
    Text1 = "Edit"
    cmdCancel.SetFocus
End Sub

Private Sub cmdUpdate_Click()
    Dim Cancel As Boolean
    If ForInsert Then
        RaiseEvent Insert(Cancel)
    Else
        RaiseEvent Update(Cancel)
    End If
    If Cancel Then Exit Sub
    RaiseEvent BindControls
    If ForInsert Then
        DataSource.MoveLast
    Else
        DataSource.Bookmark = myBookMark
        myBookMark = Empty
    End If
    RaiseEvent LockControls(True)
    ForInsert = False
    EnableScroll True
    ShowSave False
End Sub

Private Sub cmdCancel_Click()
    ForInsert = False
    EnableScroll True
    ShowSave False
    RaiseEvent BindControls
    If Not (myBookMark = Empty) Then DataSource.Bookmark = myBookMark
    myBookMark = Empty
    RaiseEvent LockControls(True)
    ShowInfo
End Sub

Private Sub cmdDelete_Click()
    Dim Cancel As Boolean
    Dim Response As VbMsgBoxResult
    myBookMark = DataSource.Bookmark
    Response = MsgBox("Record will be deleted permanently", vbYesNo)
    If Response = vbYes Then
        RaiseEvent Delete
    Else
        Exit Sub
    End If
    RaiseEvent BindControls
    With DataSource
        If .RecordCount = 0 Then
            Exit Sub
        ElseIf .EOF And Not .BOF Then
            .MovePrevious
        ElseIf Not .EOF And .BOF Then
            .MoveNext
        End If
    End With
    
End Sub

Private Sub cmdFirst_Click()
    If DataSource.RecordCount > 0 Then DataSource.MoveFirst
    
End Sub

Private Sub cmdLast_Click()
    If DataSource.RecordCount > 0 Then DataSource.MoveLast
    
End Sub

Private Sub cmdNext_Click()
    If DataSource.RecordCount > 0 Then
        DataSource.MoveNext
        If DataSource.EOF Then DataSource.MoveLast
    End If
    
End Sub

Private Sub cmdPrevious_Click()
    If DataSource.RecordCount > 0 Then
        DataSource.MovePrevious
        If DataSource.BOF Then DataSource.MoveFirst
    End If
    
End Sub

Private Sub DataSource_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    ShowInfo
End Sub
Public Sub ShowInfo()
    If DataSource.RecordCount = 0 Then Text1 = "No Record"
    Text1 = DataSource.AbsolutePosition & " of " & DataSource.RecordCount
    Text1.ForeColor = vbBlack
End Sub
Private Sub EnableScroll(Value As Boolean)
  cmdLast.Enabled = Value
  cmdNext.Enabled = Value
  cmdFirst.Enabled = Value
  cmdPrevious.Enabled = Value
  
End Sub

Private Sub ShowSave(Value As Boolean)
    cmdAddNew.Visible = Not Value
    cmdEdit.Visible = Not Value
    cmdUpdate.Visible = Value
    cmdCancel.Visible = Value
    cmdDelete.Visible = Not Value
       
End Sub
