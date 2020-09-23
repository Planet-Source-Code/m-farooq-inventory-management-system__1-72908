VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmAcOpening 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8040
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   11310
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicSrchGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   7470
      Left            =   2775
      ScaleHeight     =   7440
      ScaleWidth      =   5490
      TabIndex        =   0
      Top             =   510
      Visible         =   0   'False
      Width           =   5520
      Begin VB.TextBox TxtGrdSrch 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         Height          =   330
         Left            =   105
         TabIndex        =   1
         Top             =   750
         Width           =   5310
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MshSearch 
         Height          =   6180
         Left            =   90
         TabIndex        =   2
         Top             =   1185
         Width           =   5310
         _ExtentX        =   9366
         _ExtentY        =   10901
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         BackColorFixed  =   -2147483645
         BackColorBkg    =   15790320
         GridColorFixed  =   0
         FocusRect       =   2
         GridLinesFixed  =   1
         SelectionMode   =   1
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search Account"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1853
         TabIndex        =   4
         Top             =   450
         Width           =   1635
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "List of Accounts"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   5490
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8070
      Left            =   -15
      ScaleHeight     =   8040
      ScaleWidth      =   10650
      TabIndex        =   5
      Top             =   -15
      Width           =   10680
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   915
         Left            =   -15
         ScaleHeight     =   885
         ScaleWidth      =   10650
         TabIndex        =   6
         Top             =   7140
         Width           =   10680
         Begin VB.CommandButton cmdExit 
            BackColor       =   &H00F9F9F9&
            Caption         =   "&Exit"
            BeginProperty Font 
               Name            =   "Garamond"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   135
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   540
            Width           =   3120
         End
         Begin VB.CommandButton cmdCancel 
            BackColor       =   &H00F9F9F9&
            Caption         =   "&Cancel"
            BeginProperty Font 
               Name            =   "Garamond"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1785
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   150
            Width           =   1470
         End
         Begin VB.CommandButton cmdSave 
            BackColor       =   &H00F9F9F9&
            Caption         =   "&Save Record"
            BeginProperty Font 
               Name            =   "Garamond"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   135
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   150
            Width           =   1470
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgAcOpening 
         Height          =   6495
         Left            =   90
         TabIndex        =   10
         Top             =   555
         Width           =   10470
         _ExtentX        =   18468
         _ExtentY        =   11456
         _Version        =   393216
         Cols            =   7
         BackColorFixed  =   -2147483645
         ForeColorFixed  =   16777215
         BackColorBkg    =   16777215
         Appearance      =   0
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   7
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Opening Balances"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   4057
         TabIndex        =   12
         Top             =   52
         Width           =   2550
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Opening Balances"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   4065
         TabIndex        =   11
         Top             =   30
         Width           =   2550
      End
      Begin VB.Image Image1 
         Height          =   405
         Left            =   0
         Picture         =   "frmAcOpening.frx":0000
         Stretch         =   -1  'True
         Top             =   15
         Width           =   10665
      End
   End
End
Attribute VB_Name = "frmAcOpening"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Duplicate As Boolean
Dim RowSelect As Integer
Dim rCount As Integer
Dim mSqlQry As String
Dim TransactionId As Single
Dim TranMainCode As Single
Dim mDate As Date

Private Sub cmdDelete_Click()
    If MsgBox("Do you want to delete this Complete Voucher?", vbQuestion + vbYesNo, "Delete Voucher") = vbYes Then
        If MsgBox("System will be unable to recover the loss data. Continue ?", vbQuestion + vbYesNo, "Delete Voucher") = vbYes Then
'Deleting data
            Con.Execute "Delete from transaction_Detail where TranId = " & Val(txtId) & ""
            Con.Execute "Delete from transaction_Main where TransId = " & Val(txtId) & ""
            MsgBox "Voucher Deleted", vbInformation, "Message"
            cmdCancel_Click
        End If
    End If
    
End Sub

Private Sub cmdSave_Click()
Dim fgaAcOpeningRow As Integer

'Checking Data to Save
    If Val(fgAcOpening.Rows) < 3 Then
        MsgBox "No data to save", vbInformation, "Message"
        Exit Sub
    End If
    
    'Getting Maximum Code for Transaction Main
        MaxNumber "TransId", "Max_Codes"
        TranMainCode = Val(MaxNmbr)
        
        TransactionRef = "OpBal"
    
    'Inserting Data to TransactionMain Table
            Con.Execute "Insert into Transaction_Main (TransId,TransDate,TransType,Posted,TransRef) values (" & Val(TranMainCode) & ", '" & Date & "' , '" & TransactionType & "' ,'" & "N" & "', '" & TransactionRef & "')"
            
    '----------------------------- Inserting to Transaction Detail -----------------------------
        
            For fgacopeningRow = 1 To fgAcOpening.Rows - 2
    
    'Getting Maximum Code for transaction_Detail
        MaxNumber "TransDetId", "Max_Codes"
        TransactionId = Val(MaxNmbr)
                    
                Con.Execute "Insert into Transaction_Detail (TranDetId, TranId, AcId, Descript, DrAmount, CrAmount) values (" & Val(TransactionId) & ", " & Val(txtId) & ", " & Val(fgAcOpening.TextMatrix(fgacopeningRow, 5)) & ", '" & fgAcOpening.TextMatrix(fgacopeningRow, 2) & "', " & Val(fgAcOpening.TextMatrix(fgacopeningRow, 3)) & ", " & Val(fgAcOpening.TextMatrix(fgacopeningRow, 4)) & ")"
           
    'Updating MaxCode for Transaction Detail
        UpdateMaxNumber "TransDetId", Val(TransactionId)
    
            Next
    
    'Updating Transaction Main Id
        UpdateMaxNumber "TransId", Val(txtId)
    
    MsgBox "Record saved", vbInformation, "Done"
    cmdCancel_Click
    
    Else
'If existing record
        TranMainCode = Val(txtId)
    
'Deleting old records from TransactionMain and Transaction Detail
        Con.Execute "Delete from Transaction_Detail where TranId = " & Val(txtId) & ""
        Con.Execute "Delete from Transaction_Main where TransId = " & Val(txtId) & ""
    
'Checking Data to Save
    If Val(fgAcOpening.Rows) < 3 Then
        MsgBox "No data to save", vbInformation, "Message"
        Exit Sub
    End If
    
'Validation for Equal Balances of Debit and Credit Side
    If Not Val(txtDebit) = Val(txtCredit) Then
        MsgBox "Debit and Credit sides should be equal", vbCritical, "Message"
        Exit Sub
    End If
        
    'Inserting Data to TransactionMain Table
            
         Con.Execute "insert into Transaction_Main (TransId,TransDate, TransType, Posted,TransRef) values (" & Val(txtId) & ", '" & Date & "' , '" & TransactionType & "' ,'" & "N" & "', '" & TransactionRef & "')"
            
    '----------------------------- Inserting to Transaction Detail -----------------------------
        
            For fgacopeningRow = 1 To fgAcOpening.Rows - 2
    
    'Getting Maximum Code for transaction_Detail
        MaxNumber "TransDetId", "Max_Codes"
        TransactionId = Val(MaxNmbr)
                    
                Con.Execute "Insert into Transaction_Detail (TranDetId, TranId, AcId, Descript, DrAmount, CrAmount) values (" & Val(TransactionId) & ", " & Val(txtId) & ", " & Val(fgAcOpening.TextMatrix(fgacopeningRow, 5)) & ", '" & fgAcOpening.TextMatrix(fgacopeningRow, 2) & "', " & Val(fgAcOpening.TextMatrix(fgacopeningRow, 3)) & ", " & Val(fgAcOpening.TextMatrix(fgacopeningRow, 4)) & ")"
           
    'Updating MaxCode for Transaction Detail
        UpdateMaxNumber "TransDetId", Val(TransactionId)
    
            Next
    
    MsgBox "Record updated", vbInformation, "Done"
    cmdCancel_Click
    
    
    End If
End Sub

Private Sub fgacopening_Click()
    RowSelect = fgAcOpening.RowSel
End Sub

Private Sub fgacopening_KeyDown(KeyCode As Integer, Shift As Integer)
'If CTRL + SPACE is pressed
    If fgAcOpening.Col = 1 Then
        If KeyCode = 32 And Shift = 2 Then
            PicSrchGrid.Visible = True
            TxtGrdSrch.Text = ""

            FillGridAccounts
            SetGridAccounts
            
            MshSearch.Col = 0
            MshSearch.Row = 1
            MshSearch.SetFocus
        End If
    End If

'Delete Row from Grid
    If KeyCode = vbKeyDelete Then
        
        If fgAcOpening.Rows > 2 And fgAcOpening.TextMatrix(fgAcOpening.Row, 1) <> "" Then
            If MsgBox("Do you want to delete this line>", vbQuestion + vbYesNo, "Delete Line") = vbYes Then
                       
                fgAcOpening.RemoveItem RowSelect
            End If
            
        Else
            MsgBox "Blank or Last line can not be deleted", vbCritical, "Message "
            Exit Sub
        End If
        
    End If
End Sub

Private Sub fgacopening_KeyPress(KeyAscii As Integer)
    EditGridAcOp fgAcOpening, KeyAscii
End Sub



Private Sub Form_Load()

'Setting up flexgrid data
    Call GridSetting

'Calling Exist Data
    Call ExistData

 'Calling Data for navigation
    Set RsNAV = New ADODB.Recordset
    If RsNAV.State = 1 Then RsNAV.Close
    RsNAV.Open "Select * from Account_Openings Order By Id", Con, adOpenStatic, adLockOptimistic

'Defining Transaction Type
    TransactionType = "Opening Balance"

End Sub

Private Sub Form_Resize()
    Me.Left = 3600
    Me.Top = 1400


End Sub


Public Sub GridSetting()
'Setting of Transaction Grid
    With fgAcOpening
        
        .ColWidth(0) = 250  'fixed
        .ColWidth(1) = 2500 'Title
        .ColWidth(2) = 2925 'Descript
        .ColWidth(3) = 1475 'Qty
        .ColWidth(4) = 1475 'Dr
        .ColWidth(5) = 1475 'Cr
        .ColWidth(6) = 0    'AcId
    
        .RowHeight(0) = 400
    
        .Rows = 2
        
        .TextMatrix(0, 1) = "Account Name"
        .TextMatrix(0, 2) = "Description"
        .TextMatrix(0, 3) = "Quantity"
        .TextMatrix(0, 4) = "Debit"
        .TextMatrix(0, 5) = "Credit"
        
        .ColAlignmentFixed(1) = 4
        .ColAlignmentFixed(2) = 4
        .ColAlignmentFixed(3) = 4
        .ColAlignmentFixed(4) = 4
        .ColAlignmentFixed(5) = 4
    
    End With
End Sub

Public Sub SetGridAccounts()
'Setting of Search Grid
    With MshSearch
        .ColWidth(1) = 2500
        .ColWidth(2) = 1500
        
        .TextMatrix(0, 0) = "ID"
        .TextMatrix(0, 1) = "Account Title"
        .TextMatrix(0, 2) = "Account Type"
        
        .RowHeight(0) = 400
    
        .ColAlignmentFixed(0) = 4
        .ColAlignmentFixed(1) = 4
        .ColAlignmentFixed(2) = 4
    
    End With
    
End Sub

Public Sub FillGridAccounts()
'Filling all Accounts Data in Search grid
    SQLQRY = "Select AcId, AcTitle, AcType from ViewHeadWise"
    
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
        
        RS.Open SQLQRY, Con, adOpenStatic, adLockReadOnly
            
        Set MshSearch.DataSource = RS
End Sub

Private Sub TxtGrdSrch_Change()
    Call SearchRecord
End Sub
Public Sub SearchRecord()
'Filling the Search grid with Critarial Data
    
    Dim SearchedRowCount As Integer
    
    If PicSrchGrid.Visible = True Then
        
        MshSearch.Rows = 2
        MshSearch.Row = 0
        
        SearchedRowCount = 0
        
        SQLQRY = "Select AcId, AcTitle, AcType from ViewHeadWise where AcTitle Like '" & TxtGrdSrch.Text & "%'"
            
        Set RS = New ADODB.Recordset
        If RS.State = 1 Then RS.Close
            RS.Open SQLQRY, Con, adOpenStatic, adLockReadOnly
                
                If RS.RecordCount <= 0 Then
                    MshSearch.Clear
                    
                    With MshSearch
                        .TextMatrix(0, 0) = "ID"
                        .TextMatrix(0, 1) = "Account Title"
                        .TextMatrix(0, 2) = "Account Type"
                    End With
                    
                    
                    MshSearch.Rows = 2
                    Exit Sub
                End If
                
                For SearchedRowCount = 1 To RS.RecordCount
                    
                    If SearchedRowCount >= MshSearch.Rows - 1 Then
                        MshSearch.Rows = MshSearch.Rows + 1
                        MshSearch.Row = MshSearch.Row + 1
                    End If
                    
                    MshSearch.TextMatrix(SearchedRowCount, 0) = RS(0)
                    MshSearch.TextMatrix(SearchedRowCount, 1) = RS(1)
                    MshSearch.TextMatrix(SearchedRowCount, 2) = RS(2)
                    
                    
                    RS.MoveNext
                Next
                
                MshSearch.Col = 0
                MshSearch.Row = 0
                MshSearch.ColAlignment(0) = 3
    End If
    
End Sub

Private Sub MshSearch_DblClick()
    If MshSearch.Row = 0 Then
        Exit Sub
    End If
    
'Checking for duplicate entry in grid
        Call CheckDuplicate
        
        If Duplicate = True Then
            Exit Sub
        End If
        
        fgAcOpening.TextMatrix(fgAcOpening.Row, 1) = MshSearch.TextMatrix(MshSearch.Row, 1)
        fgAcOpening.TextMatrix(fgAcOpening.Row, 6) = MshSearch.TextMatrix(MshSearch.Row, 0)
        
        PicSrchGrid.Visible = False
        fgAcOpening.Col = 1
        fgAcOpening.SetFocus
End Sub

Public Sub CheckDuplicate()
Dim dRow As Integer
   
    For dRow = 1 To fgAcOpening.Rows - 2
        If fgAcOpening.TextMatrix(dRow, 6) = MshSearch.TextMatrix(MshSearch.Row, 0) Then
            MsgBox "Account Title aleready selected", vbInformation, "Message"
            Duplicate = True
            Exit Sub
        Else
            Duplicate = False
        End If
    Next
    
        
End Sub

Private Sub MshSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        MshSearch_DblClick
    End If
    
    If KeyAscii = 27 Then
        PicSrchGrid.Visible = False
        fgAcOpening.SetFocus
        Exit Sub
    End If
    
    If KeyAscii = 8 Then
        If TxtGrdSrch.Text <> "" Then TxtGrdSrch.Text = Left$(TxtGrdSrch.Text, (Len(TxtGrdSrch.Text) - 1))
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
    Else
        TxtGrdSrch.Text = TxtGrdSrch.Text + Chr$(KeyAscii)
    End If
    
End Sub

Private Sub TxtGrdSrch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        PicSrchGrid.Visible = False
        fgAcOpening.SetFocus
        Exit Sub
    End If
End Sub

Public Sub ExistData()

'Getting data from TransactionDetail
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
        
        RS.Open "Select * from ViewAccountOpenings", Con, adOpenStatic, adLockOptimistic
            
            If RS.RecordCount <= 0 Then
                Exit Sub
            End If
            
            fgAcOpening.Rows = 2
            fgAcOpening.Row = 1
            
            For rCount = 1 To RS.RecordCount
                fgAcOpening.TextMatrix(fgAcOpening.Row, 1) = RS(0)  'Title
                fgAcOpening.TextMatrix(fgAcOpening.Row, 2) = RS(1)  'Description
                fgAcOpening.TextMatrix(fgAcOpening.Row, 3) = RS(2)  'Qty
                fgAcOpening.TextMatrix(fgAcOpening.Row, 4) = RS(3)  'Debit
                fgAcOpening.TextMatrix(fgAcOpening.Row, 5) = RS(4)  'Credit
                fgAcOpening.TextMatrix(fgAcOpening.Row, 6) = RS(5)  'Ac Code
                
                
                fgAcOpening.Rows = fgAcOpening.Rows + 1
                fgAcOpening.Row = fgAcOpening.Row + 1
                                
                RS.MoveNext
            Next

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdNew_Click()
    mDate = Now
                
    fgAcOpening.Row = fgAcOpening.Rows - 1
    fgAcOpening.Col = 1
    fgAcOpening.SetFocus
    
End Sub
Private Sub cmdCancel_Click()
    Clear Me
    
    fgAcOpening.Clear
    
    GridSetting
            
    Call ExistData
  
End Sub
