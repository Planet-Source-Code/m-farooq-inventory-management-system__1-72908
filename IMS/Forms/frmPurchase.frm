VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPurchase 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8625
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   9525
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   9525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox PicSrchGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   7470
      Left            =   1950
      ScaleHeight     =   7440
      ScaleWidth      =   5490
      TabIndex        =   8
      Top             =   870
      Visible         =   0   'False
      Width           =   5520
      Begin VB.TextBox TxtGrdSrch 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         Height          =   330
         Left            =   105
         TabIndex        =   9
         Top             =   750
         Width           =   5310
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MshSearch 
         Height          =   6180
         Left            =   90
         TabIndex        =   10
         Top             =   1185
         Width           =   5310
         _ExtentX        =   9366
         _ExtentY        =   10901
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         BackColorFixed  =   13408563
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
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   0
         Width           =   5490
      End
   End
   Begin VB.TextBox txtId 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   420
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   1980
      Width           =   1650
   End
   Begin VB.PictureBox picFind 
      Appearance      =   0  'Flat
      BackColor       =   &H00CC9933&
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   2565
      ScaleHeight     =   1785
      ScaleWidth      =   4455
      TabIndex        =   13
      Top             =   3390
      Visible         =   0   'False
      Width           =   4485
      Begin VB.OptionButton optId 
         Appearance      =   0  'Flat
         BackColor       =   &H00CC9933&
         Caption         =   "Find By Purchase ID"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   225
         TabIndex        =   17
         Top             =   495
         Width           =   2025
      End
      Begin VB.TextBox txtFindId 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   225
         TabIndex        =   16
         Top             =   780
         Width           =   3945
      End
      Begin VB.CommandButton cmdFindCancel 
         BackColor       =   &H00F9F9F9&
         Caption         =   "Cancel"
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
         Left            =   2190
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1260
         Width           =   1470
      End
      Begin VB.CommandButton cmdFind 
         BackColor       =   &H00F9F9F9&
         Caption         =   "Find Record"
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
         Left            =   675
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1260
         Width           =   1470
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Find Purchase"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   1462
         TabIndex        =   18
         Top             =   15
         Width           =   1530
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Find Purchase"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   270
         Left            =   1470
         TabIndex        =   19
         Top             =   15
         Width           =   1530
      End
      Begin VB.Image Image2 
         Height          =   270
         Index           =   1
         Left            =   0
         Picture         =   "frmPurchase.frx":0000
         Stretch         =   -1  'True
         Top             =   15
         Width           =   4455
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   -15
      ScaleHeight     =   450
      ScaleWidth      =   9525
      TabIndex        =   5
      Top             =   8175
      Width           =   9555
      Begin VB.TextBox txtTotal 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   7830
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   105
         Width           =   1485
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Press CTRL+SPACE in Supplier Name and Product Name for List"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   90
         TabIndex        =   36
         Top             =   150
         Width           =   6105
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   6570
         TabIndex        =   7
         Top             =   150
         Width           =   1140
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F0F0&
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   -15
      ScaleHeight     =   450
      ScaleWidth      =   9525
      TabIndex        =   0
      Top             =   2340
      Width           =   9555
      Begin VB.TextBox TxtAcCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2250
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   90
         Width           =   1485
      End
      Begin VB.TextBox TxtName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3720
         TabIndex        =   1
         Top             =   90
         Width           =   5565
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Code and Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   210
         Left            =   90
         TabIndex        =   4
         Top             =   127
         Width           =   2220
      End
   End
   Begin MSComCtl2.DTPicker Dtp1 
      Height          =   300
      Left            =   6195
      TabIndex        =   21
      Top             =   1980
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   529
      _Version        =   393216
      Format          =   66125824
      CurrentDate     =   40158
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgPurchase 
      Height          =   5265
      Left            =   0
      TabIndex        =   2
      Top             =   2865
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   9287
      _Version        =   393216
      Cols            =   7
      BackColorFixed  =   13408563
      ForeColorFixed  =   16777215
      BackColorBkg    =   15790320
      FocusRect       =   2
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
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P U R C H A S E S"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   3045
      TabIndex        =   34
      Top             =   1005
      Width           =   3435
   End
   Begin VB.Label lblnav 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   11.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   1
      Left            =   915
      TabIndex        =   33
      Top             =   405
      Width           =   225
   End
   Begin VB.Label lblnav 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   11.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   2
      Left            =   1245
      TabIndex        =   32
      Top             =   405
      Width           =   225
   End
   Begin VB.Label lblnav 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   11.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   0
      Left            =   585
      TabIndex        =   31
      Top             =   405
      Width           =   225
   End
   Begin VB.Label lblnav 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   11.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   3
      Left            =   1575
      TabIndex        =   30
      Top             =   405
      Width           =   225
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   3285
      TabIndex        =   29
      Top             =   375
      Width           =   390
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   4185
      TabIndex        =   28
      Top             =   375
      Width           =   450
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Find"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   7335
      TabIndex        =   27
      Top             =   375
      Width           =   375
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   6255
      TabIndex        =   26
      Top             =   375
      Width           =   570
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   8220
      TabIndex        =   25
      Top             =   375
      Width           =   330
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   5145
      TabIndex        =   24
      Top             =   375
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   0
      Left            =   -15
      Top             =   0
      Width           =   9555
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   7020
      TabIndex        =   23
      Top             =   1755
      Width           =   1275
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   600
      TabIndex        =   22
      Top             =   1770
      Width           =   1305
   End
   Begin VB.Image Image 
      Height          =   465
      Left            =   0
      Picture         =   "frmPurchase.frx":0E2D
      Stretch         =   -1  'True
      Top             =   240
      Width           =   9525
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P U R C H A S E S"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3060
      TabIndex        =   35
      Top             =   1020
      Width           =   3435
   End
End
Attribute VB_Name = "frmPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Duplicate As Boolean
Dim RowSelect As Integer
Dim FgRow As Integer
Dim rCount As Integer
Dim mSqlQry As String
Dim PurchaseId As Single
Dim PurMainCode As Single
Dim ObjectFoucs As Boolean 'True if Focus on TxtName(Supplier Name) . False if Focus on Grid
Dim OpQty As Single
Dim OpAmount As Single
Dim AvgCost As Single


Private Sub Form_Resize()
    Me.Top = 1500
    Me.Left = 4200

End Sub

Private Sub lbl_Click(Index As Integer)
    Select Case Index
        Case 0                          'New
        'Clearing all controls for New Data
            Clear Me
            fgPurchase.Clear
            
            modWriteInGrid.SetPSGRID fgPurchase
            
            Call AutoId
            
            Dtp1.SetFocus
            Dtp1.Value = Now
                    
            Modes True, False, Me
        
        'Lock Navigation
            LockNav Me
    
        Case 1                          'Save
            Dim TranDetId As Single
            Dim PurchaseRef As String
            Dim TranMainId As Single
            
            'Checking Data to Save
                If Val(fgPurchase.Rows) < 3 Then
                    MsgBox "No data to save", vbInformation, "Message"
                    Exit Sub
                End If
                
            'If Account not selected
                If TxtAcCode.Text = "" Or Val(TxtAcCode) = 0 Then
                    MsgBox "Select Account Name", vbCritical, "Message"
                    TxtName.SetFocus
                    Exit Sub
                End If
                
            'If New Purchase
            
                If lbl(0).Enabled = False Then
                    
                    TransactionRef = "Purchase-" & Val(txtId)
            
            'Getting Maximum Code for Purchase Main
                    MaxNumber "PurId", "Max_Codes"
                    PurMainCode = Val(MaxNmbr)
                    
            'Getting Maximum Code for Transaction Main
                    MaxNumber "TransId", "Max_Codes"
                    TranMainId = Val(MaxNmbr)
                    
            'Inserting Data to TransactionMain Table
                        Con.Execute "insert into Transaction_Main (TransId,TransDate,TransType,Posted,TransRef) values (" & Val(TranMainId) & ", '" & Format(Dtp1.Value, "mm/dd/yyyy") & "' , '" & TransactionType & "' ,'" & "N" & "', '" & TransactionRef & "')"
                
            '---------------------INSERTING DATA INTO TRANSACTION_DETAIL for CREDIT ENTRY
            'Getting Maximum Code for Transaction_Detail
                    MaxNumber "TransDetId", "Max_Codes"
                    TranDetId = Val(MaxNmbr)
            
                    PurchaseRef = "Purchase # " & Val(txtId)
            'Inserting Data to Transaction Detail
                            Con.Execute "Insert into Transaction_Detail (TranDetId, TranId, AcId, Descript, DrAmount, CrAmount) values (" & Val(TranDetId) & ", " & Val(TranMainId) & ", " & Val(TxtAcCode) & ", '" & PurchaseRef & "', 0 , " & Val(txtTotal) & ")"
                       
            'Updating MaxCode for Transaction Detail
                    UpdateMaxNumber "TransDetId", Val(TranDetId)
                
                
            '----------------------------- Inserting to Purchase Main -----------------------------
                
            'Inserting Data to PurchaseMain Table
                        Con.Execute "insert into Purchase_Main (PurId,PurDate,TotalAmount) values (" & Val(txtId) & ", '" & Format(Dtp1.Value, "mm/dd/yyyy") & "' , " & Val(txtTotal) & " )"
                        
            '----------------------------- Inserting to Purchase Detail -----------------------------
                    
                        For FgRow = 1 To fgPurchase.Rows - 2
                
            'Getting Maximum Code for Purchase_Detail
                    MaxNumber "PurDetId", "Max_Codes"
                    PurchaseId = Val(MaxNmbr)
            
            'Inserting Data to Purchase Detail
                            Con.Execute "Insert into Purchase_Detail (PurDetId, PurId, AcId, ProdId, Qty, Price) values (" & Val(PurchaseId) & ", " & Val(txtId) & ", " & Val(TxtAcCode) & ", '" & fgPurchase.TextMatrix(FgRow, 5) & "', " & Val(fgPurchase.TextMatrix(FgRow, 2)) & ", " & Val(fgPurchase.TextMatrix(FgRow, 3)) & ")"
            'Updating MaxCode for Purchase Detail
                    UpdateMaxNumber "PurDetId", Val(PurchaseId)
                
                
            '---------------------INSERTING DATA INTO TRANSACTION_DETAIL for DEBIT ENTRY
            'Getting Maximum Code for Transaction_Detail
                    MaxNumber "TransDetId", "Max_Codes"
                    TranDetId = Val(MaxNmbr)
            
            
                    PurchaseRef = "Purchase # " & Val(txtId)
            'Inserting Data to Transaction Detail
                            Con.Execute "Insert into Transaction_Detail (TranDetId, TranId, AcId, Descript, DrAmount, CrAmount) values (" & Val(TranDetId) & ", " & Val(TranMainId) & ", " & Val(fgPurchase.TextMatrix(FgRow, 5)) & ", '" & PurchaseRef & "', " & Val(fgPurchase.TextMatrix(FgRow, 4)) & ", 0)"
                       
            'Updating MaxCode for Transaction Detail
                    UpdateMaxNumber "TransDetId", Val(TranDetId)
                
                
                
                        Next
                
            'Updating Purchase Main Id
                    UpdateMaxNumber "PurId", Val(txtId)
                
            'Updating Transaction Main Id
                    UpdateMaxNumber "TransId", Val(TranMainId)
                
            'Getting Average Cost for the specifig product
                    Call GetTheAverage
                
                MsgBox "Record saved", vbInformation, "Done"
                lbl_Click (2)
                
                
                
                Else
            
            Dim MTranId As Single
            'If existing record
                    PurMainCode = Val(txtId)
                    TransactionRef = "Purchase-" & Val(txtId)
                
            'Deleting old records from PurchaseMain and Purchase Detail
                    Con.Execute "Delete from Purchase_Detail where PurId = " & Val(txtId) & ""
                    Con.Execute "Delete from Purchase_Main where PurId = " & Val(txtId) & ""
                    
                    Set RS = New ADODB.Recordset
                    If RS.State = 1 Then RS.Close
                        RS.Open "Select * from Transaction_Main where TransRef = '" & TransactionRef & "'", Con, adOpenStatic, adLockOptimistic
                            
                            MTranId = Val(RS(0))
                    
                    Con.Execute "Delete from Transaction_Detail Where TranId = " & Val(MTranId) & ""
                    Con.Execute "Delete from Transaction_Main where TransRef = '" & TransactionRef & "'"
            
            
            'Checking Data to Save
                If Val(fgPurchase.Rows) < 3 Then
                    MsgBox "No data to save", vbInformation, "Message"
                    Exit Sub
                End If
                
            'If Account not selected
                If TxtAcCode.Text = "" Or Val(TxtAcCode) = 0 Then
                    MsgBox "Select Account Name", vbCritical, "Message"
                    TxtName.SetFocus
                    Exit Sub
                End If
                    
            'Inserting Data to TransactionMain Table
                        Con.Execute "insert into Transaction_Main (TransId,TransDate,TransType,Posted,TransRef) values (" & Val(MTranId) & ", '" & Format(Dtp1.Value, "mm/dd/yyyy") & "' , '" & TransactionType & "' ,'" & "N" & "', '" & TransactionRef & "')"
                
            '---------------------INSERTING DATA INTO TRANSACTION_DETAIL for CREDIT ENTRY
            'Getting Maximum Code for Transaction_Detail
                    MaxNumber "TransDetId", "Max_Codes"
                    TranDetId = Val(MaxNmbr)
            
                    PurchaseRef = "Purchase # " & Val(txtId)
            'Inserting Data to Transaction Detail
                            Con.Execute "Insert into Transaction_Detail (TranDetId, TranId, AcId, Descript, DrAmount, CrAmount) values (" & Val(TranDetId) & ", " & Val(MTranId) & ", " & Val(TxtAcCode) & ", '" & PurchaseRef & "', 0 , " & Val(txtTotal) & ")"
                       
            'Updating MaxCode for Transaction Detail
                    UpdateMaxNumber "TransDetId", Val(TranDetId)
                
                
            '----------------------------- Inserting to Purchase Main -----------------------------
                
            'Inserting Data to PurchaseMain Table
                        Con.Execute "insert into Purchase_Main (PurId,PurDate,TotalAmount) values (" & Val(txtId) & ", '" & Format(Dtp1.Value, "mm/dd/yyyy") & "' , " & Val(txtTotal) & " )"
                        
            '----------------------------- Inserting to Purchase Detail -----------------------------
                    
                        For FgRow = 1 To fgPurchase.Rows - 2
                
            'Getting Maximum Code for Purchase_Detail
                    MaxNumber "PurDetId", "Max_Codes"
                    PurchaseId = Val(MaxNmbr)
            
            'Inserting Data to Purchase Detail
                            Con.Execute "Insert into Purchase_Detail (PurDetId, PurId, AcId, ProdId, Qty, Price) values (" & Val(PurchaseId) & ", " & Val(txtId) & ", " & Val(TxtAcCode) & ", '" & fgPurchase.TextMatrix(FgRow, 5) & "', " & Val(fgPurchase.TextMatrix(FgRow, 2)) & ", " & Val(fgPurchase.TextMatrix(FgRow, 3)) & ")"
                       
            'Updating MaxCode for Purchase Detail
                    UpdateMaxNumber "PurDetId", Val(PurchaseId)
                
            '---------------------INSERTING DATA INTO TRANSACTION_DETAIL for DEBIT ENTRY (Product in grid Entry)
            'Getting Maximum Code for Transaction_Detail
                    MaxNumber "TransDetId", "Max_Codes"
                    TranDetId = Val(MaxNmbr)
            
            
                    PurchaseRef = "Purchase # " & Val(txtId)
            'Inserting Data to Transaction Detail
                            Con.Execute "Insert into Transaction_Detail (TranDetId, TranId, AcId, Descript, DrAmount, CrAmount) values (" & Val(TranDetId) & ", " & Val(MTranId) & ", " & Val(fgPurchase.TextMatrix(FgRow, 5)) & ", '" & PurchaseRef & "', " & Val(fgPurchase.TextMatrix(FgRow, 4)) & ", 0)"
                       
            'Updating MaxCode for Transaction Detail
                    UpdateMaxNumber "TransDetId", Val(TranDetId)
                
                        Next
                
            'Getting Average Cost for the specifig product
                    Call GetTheAverage
                
                MsgBox "Record Updated", vbInformation, "Done"
                lbl_Click (2)
                
                End If
    
        Case 2                                      'Cancel
                Clear Me
                
            'UnLock Navigation
                UnLockNav Me
                
                fgPurchase.Clear
                
                modWriteInGrid.SetPSGRID fgPurchase
                        
                Call ExistData
                
                Modes False, True, Me
                
                Dtp1.SetFocus
                
                If RsNAV.RecordCount <= 0 Then
                    Exit Sub
                Else
                    RsNAV.Requery
                    RsNAV.MoveFirst
                End If
            
            
        
        Case 3                                      'Delete
            Dim mRef As String
            Dim mTransID As Integer
            mRef = "Purchase-" & Val(txtId)
                            
                If MsgBox("Do you want to delete this purchase?", vbQuestion + vbYesNo, "Delete Voucher") = vbYes Then
                    If MsgBox("System will be unable to recover the loss data. Continue ?", vbQuestion + vbYesNo, "Delete Voucher") = vbYes Then
            'Deleting data from Transaction Main and Detail
                        Set RS = New ADODB.Recordset
                        If RS.State = 1 Then RS.Close
                            RS.Open "Select TransId from Transaction_Main where TransRef = '" & mRef & "'", Con, adOpenStatic, adLockOptimistic
                                mTransID = Val(RS(0))
                        
                        Con.Execute "Delete from Transaction_Main where TransID = " & mTransID & ""
                        Con.Execute "Delete from Transaction_Detail where TranId = " & mTransID & ""
                        
            'Deleting data from Prucase Main and Detial
                        Con.Execute "Delete from Purchase_Detail where PurId = " & Val(txtId) & ""
                        Con.Execute "Delete from Purchase_Main where PurId = " & Val(txtId) & ""
                        MsgBox "Record Deleted", vbInformation, "Message"
                        lbl_Click (2)
                    End If
                End If
    
        Case 4                                      'Find
                picFind.Visible = True
                    optId.Value = True
                    txtFindId.SetFocus
                    txtFindId.Text = ""
    
        Case 5                                      'Exit
                Unload Me
    End Select
        
End Sub

Private Sub lblnav_Click(Index As Integer)
    Select Case Index
        
        Case 0              'Move First
            On Error Resume Next
            RsNAV.MoveFirst
            
            If RsNAV.BOF = True Then
                MsgBox "First Record", vbInformation, "Message"
                RsNAV.MoveFirst
                Exit Sub
            Else
                Call NAVData
            End If
        
        Case 1              'Move previous
            On Error Resume Next
            RsNAV.MovePrevious
        
            If RsNAV.BOF = True Then
                MsgBox "First Record", vbInformation, "Message"
                RsNAV.MoveFirst
            Else
                Call NAVData
            End If
    
    
        Case 2              'Move Next
            On Error Resume Next
            RsNAV.MoveNext
        
            If RsNAV.EOF = True Then
                MsgBox "Last Record", vbInformation, "Message"
                RsNAV.MoveLast
            Else
                Call NAVData
            End If
            
        Case 3              'Move last
            On Error Resume Next
            RsNAV.MoveLast
            
            If RsNAV.EOF = True Then
                MsgBox "Last Record", vbInformation, "Message"
                RsNAV.MoveLast
            Else
                Call NAVData
            End If
        
        
        
        
        End Select
        
End Sub

Private Sub lblnav_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 0, 1, 2, 3
            MouseOver lblnav(Index)
    End Select
    
End Sub

Private Sub cmdFind_Click()
    If txtFindId.Text = "" Or Val(txtFindId) = 0 Then
        MsgBox "Enter Purchase ID", vbCritical, "Message.."
        txtFindId.SetFocus
        Exit Sub
    End If
    
    Call FindRecord

End Sub

Private Sub cmdFindCancel_Click()
    picFind.Visible = False
    Dtp1.SetFocus
    
End Sub




Private Sub Dtp1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        TxtName.SetFocus
    End If
End Sub

Private Sub fgpurchase_Click()
    RowSelect = fgPurchase.RowSel
End Sub


Private Sub fgPurchase_GotFocus()
    ObjectFoucs = False
End Sub

Private Sub fgpurchase_KeyDown(KeyCode As Integer, Shift As Integer)
'If CTRL + SPACE is pressed
    If fgPurchase.Col = 1 Then
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
        
        If fgPurchase.Rows > 2 And fgPurchase.TextMatrix(fgPurchase.Row, 1) <> "" Then
            If MsgBox("Do you want to delete this line?", vbQuestion + vbYesNo, "Delete Line") = vbYes Then
                txtTotal.Text = Val(txtTotal) - Val(fgPurchase.TextMatrix(fgPurchase.Row, 4))
                        
                fgPurchase.RemoveItem RowSelect
            End If
            
        Else
            MsgBox "Blank or Last line can not be deleted", vbCritical, "Message "
            Exit Sub
        End If
        
    End If

End Sub

Private Sub fgpurchase_KeyPress(KeyAscii As Integer)
    EditGridPS fgPurchase, KeyAscii
End Sub

Private Sub fgPurchase_LeaveCell()
    Select Case fgPurchase.Col
        Case 1
            Call Calculatetotal
            
        Case 2, 3
'Total of Row i.e Qty * Price
                fgPurchase.TextMatrix(fgPurchase.Row, 4) = Val(fgPurchase.TextMatrix(fgPurchase.Row, 2)) * Val(fgPurchase.TextMatrix(fgPurchase.Row, 3))
            
    End Select


End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'Changing Control focus on Enter
    ChangeFocusOnEnter KeyAscii, Me

End Sub

Private Sub Form_Load()

    
'Setting up flexgrid
    SetPSGRID fgPurchase

'Call Existing Data
    Call ExistData

'Calculating Total
    Call Calculatetotal

 'Calling Data for navigation
    Set RsNAV = New ADODB.Recordset
    If RsNAV.State = 1 Then RsNAV.Close
    RsNAV.Open "Select * from Purchase_Main Order By PurId", Con, adOpenStatic, adLockOptimistic

'Defining Purchase Type
    TransactionType = "Purchase"

' Setting ObjectFocus variable to True
    ObjectFoucs = True

End Sub


Public Sub FillGridAccounts()
'Filling all Accounts Data in Search grid
    If ObjectFoucs = True Then  ' If focus on TxtName
        SQLQry = "SELECT ViewHeadWise.AcId, ViewHeadWise.AcTitle, ViewHeadWise.AcType FROM ViewHeadWise Where AcType NOT IN ('Product')"
    Else    ' If focus on Product Grid
        SQLQry = "SELECT ViewHeadWise.AcId, ViewHeadWise.AcTitle, ViewHeadWise.AcType FROM ViewHeadWise Where AcType IN ('Product')"
    End If
    
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
        
        RS.Open SQLQry, Con, adOpenStatic, adLockReadOnly
            
        Set MshSearch.DataSource = RS
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseNormalOnLbl
End Sub

Private Sub Image_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseNormalOnLbl
End Sub

Private Sub lbl_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 0, 1, 2, 3, 4, 5
            MouseOver lbl(Index)
            
    End Select

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
        
    If ObjectFoucs = True Then  ' If focus on TxtName
        SQLQry = "SELECT ViewHeadWise.AcId, ViewHeadWise.AcTitle, ViewHeadWise.AcType FROM ViewHeadWise Where AcType NOT IN ('Product') And AcTitle Like '" & TxtGrdSrch.Text & "%'"
    Else    ' If focus on Product Grid
        SQLQry = "SELECT ViewHeadWise.AcId, ViewHeadWise.AcTitle, ViewHeadWise.AcType FROM ViewHeadWise Where AcType IN ('Product') And AcTitle Like '" & TxtGrdSrch.Text & "%'"
    End If
        
        Set RS = New ADODB.Recordset
        If RS.State = 1 Then RS.Close
            RS.Open SQLQry, Con, adOpenStatic, adLockReadOnly
                
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
                
'Fetching data into Search Grid
                Set MshSearch.DataSource = RS
                
''                For SearchedRowCount = 1 To RS.RecordCount
''
''                    If SearchedRowCount >= MshSearch.Rows - 1 Then
''                        MshSearch.Rows = MshSearch.Rows + 1
''                        MshSearch.Row = MshSearch.Row + 1
''                    End If
''
''                    MshSearch.TextMatrix(SearchedRowCount, 0) = RS(0)
''                    MshSearch.TextMatrix(SearchedRowCount, 1) = RS(1)
''                    MshSearch.TextMatrix(SearchedRowCount, 2) = RS(2)
''
''
''                    RS.MoveNext
''                Next
                
                MshSearch.Col = 0
                MshSearch.Row = 0
                MshSearch.ColAlignment(0) = 3
    End If
    
End Sub

Private Sub MshSearch_DblClick()
    If MshSearch.Row = 0 Then
        Exit Sub
    End If
    
    If ObjectFoucs = True Then
        TxtAcCode.Text = Val(MshSearch.TextMatrix(MshSearch.Row, 0))
        TxtName.Text = MshSearch.TextMatrix(MshSearch.Row, 1)
        TxtName.SetFocus
        ObjectFoucs = False
    Else
'Checking for duplicate entry in grid
        Call CheckDuplicate
        
        If Duplicate = True Then
            Exit Sub
        End If
        fgPurchase.TextMatrix(fgPurchase.Row, 1) = MshSearch.TextMatrix(MshSearch.Row, 1)
        fgPurchase.TextMatrix(fgPurchase.Row, 5) = MshSearch.TextMatrix(MshSearch.Row, 0)
        fgPurchase.SetFocus
    End If
    
        
        PicSrchGrid.Visible = False
        fgPurchase.Col = 1
        
End Sub

Public Sub CheckDuplicate()
Dim dRow As Integer
   
    For dRow = 1 To fgPurchase.Rows - 2
        If fgPurchase.TextMatrix(dRow, 5) = MshSearch.TextMatrix(MshSearch.Row, 0) Then
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
        fgPurchase.SetFocus
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
        fgPurchase.SetFocus
        Exit Sub
    End If
End Sub

Public Sub ExistData()

'Getting data from Purchase
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
    
        RS.Open "SELECT PurId, PurDate, TotalAmount from Purchase_Main", Con, adOpenStatic, adLockOptimistic
            
            If RS.EOF = True Then
                Exit Sub
            End If
            
            txtId.Text = Val(RS(0))
            Dtp1.Value = RS(1)
        RS.Close
            
'Getting Supplier Data
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
        RS.Open "SELECT Accounts.AcId, Accounts.AcTitle, Purchase_Detail.PurId FROM Accounts INNER JOIN Purchase_Detail ON Accounts.AcId = Purchase_Detail.AcId where PurId = " & Val(txtId) & "", Con, adOpenStatic, adLockOptimistic
            TxtAcCode.Text = Val(RS(0))
            TxtName.Text = RS(1)
            
            
'Getting data from Purchase_Detail
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
        
       mSqlQry = "SELECT * from ViewPurchase where PurId = " & Val(txtId) & ""

        
        RS.Open mSqlQry, Con, adOpenStatic, adLockOptimistic
            
            If RS.RecordCount <= 0 Then
                Exit Sub
            End If
            
            fgPurchase.Rows = 2
            fgPurchase.Row = 1
            
            For rCount = 1 To RS.RecordCount
                fgPurchase.TextMatrix(fgPurchase.Row, 1) = RS(0)                   'Prod Name
                fgPurchase.TextMatrix(fgPurchase.Row, 2) = Val(RS(1))              'Qty
                fgPurchase.TextMatrix(fgPurchase.Row, 3) = Val(RS(2))              'Price
                fgPurchase.TextMatrix(fgPurchase.Row, 4) = Val(RS(1)) * Val(RS(2)) 'Amount
                fgPurchase.TextMatrix(fgPurchase.Row, 5) = Val(RS(3))              'Prod Code
                
                fgPurchase.Rows = fgPurchase.Rows + 1
                fgPurchase.Row = fgPurchase.Row + 1
                                
                RS.MoveNext
            Next

End Sub

Public Sub NAVData()

'Getting data from Purchase
            
            txtId.Text = Val(RsNAV(0))
            Dtp1.Value = RsNAV(1)
            
'Getting Supplier Data
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
        RS.Open "SELECT Accounts.AcId, Accounts.AcTitle, Purchase_Detail.PurId FROM Accounts INNER JOIN Purchase_Detail ON Accounts.AcId = Purchase_Detail.AcId where PurId = " & Val(txtId) & "", Con, adOpenStatic, adLockOptimistic
            TxtAcCode.Text = Val(RS(0))
            TxtName.Text = RS(1)
            
            
'Getting data from Purchase_Detail
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
        
       mSqlQry = "SELECT * from ViewPurchase where PurId = " & Val(txtId) & ""

        
        RS.Open mSqlQry, Con, adOpenStatic, adLockOptimistic
            
            If RS.RecordCount <= 0 Then
                Exit Sub
            End If
            
            fgPurchase.Rows = 2
            fgPurchase.Row = 1
            
            For rCount = 1 To RS.RecordCount
                fgPurchase.TextMatrix(fgPurchase.Row, 1) = RS(0)                   'Prod Name
                fgPurchase.TextMatrix(fgPurchase.Row, 2) = Val(RS(1))              'Qty
                fgPurchase.TextMatrix(fgPurchase.Row, 3) = Val(RS(2))              'Price
                fgPurchase.TextMatrix(fgPurchase.Row, 4) = Val(RS(1)) * Val(RS(2)) 'Amount
                fgPurchase.TextMatrix(fgPurchase.Row, 5) = Val(RS(3))              'Prod Code
                
                fgPurchase.Rows = fgPurchase.Rows + 1
                fgPurchase.Row = fgPurchase.Row + 1
                                
                RS.MoveNext
            Next

            Calculatetotal
    
End Sub


Public Sub AutoId()
'Calling MaxNumber function to get Auto Id for the record
    MaxNumber "PurId", "Max_Codes"
    txtId.Text = Val(MaxNmbr)
End Sub

Public Sub Calculatetotal()
    Dim mRow As Integer
                
        txtTotal.Text = ""
        For mRow = 1 To fgPurchase.Rows - 1
            txtTotal.Text = Val(txtTotal) + Val(fgPurchase.TextMatrix(mRow, 4))
        Next mRow
End Sub

Private Sub TxtAcCode_GotFocus()
    TxtName.SetFocus
End Sub

Private Sub TxtName_GotFocus()
    ObjectFoucs = True
End Sub

Private Sub TxtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 32 And Shift = 2 Then
        PicSrchGrid.Visible = True
        TxtGrdSrch.Text = ""
    
        FillGridAccounts
        SetGridAccounts
        
        MshSearch.Col = 0
        MshSearch.Row = 1
        MshSearch.SetFocus
    End If
    
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

Private Sub TxtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        fgPurchase.SetFocus
    Else
        KeyAscii = 0
        Exit Sub
    End If
    
        
End Sub

Public Sub FindRecord()
            mSqlQry = "Select * from Purchase_Main Where PurId = " & Val(txtFindId) & ""
            
            Set RS = New ADODB.Recordset
            If RS.State = 1 Then RS.Close
                RS.Open mSqlQry, Con, adOpenStatic, adLockOptimistic
                    If RS.RecordCount <= 0 Then
                        MsgBox "No purchase found with this ID", vbCritical, "Message.."
                        txtFindId.SetFocus
                        Exit Sub
                    Else
                        ShowFoundData
                    End If
    
End Sub


Public Sub ShowFoundData()
'Getting data from Purchase
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
    
        RS.Open "SELECT PurId, PurDate, TotalAmount from Purchase_Main Where PurId = " & Val(txtFindId) & "", Con, adOpenStatic, adLockOptimistic
            
            If RS.EOF = True Then
                Exit Sub
            End If
            
            txtId.Text = Val(RS(0))
            Dtp1.Value = RS(1)
        RS.Close
            
'Getting Supplier Data
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
        RS.Open "SELECT Accounts.AcId, Accounts.AcTitle, Purchase_Detail.PurId FROM Accounts INNER JOIN Purchase_Detail ON Accounts.AcId = Purchase_Detail.AcId where PurId = " & Val(txtId) & "", Con, adOpenStatic, adLockOptimistic
            TxtAcCode.Text = Val(RS(0))
            TxtName.Text = RS(1)
            
            
'Getting data from Purchase_Detail
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
        
       mSqlQry = "SELECT * from ViewPurchase where PurId = " & Val(txtId) & ""

        
        RS.Open mSqlQry, Con, adOpenStatic, adLockOptimistic
            
            If RS.RecordCount <= 0 Then
                Exit Sub
            End If
            
            fgPurchase.Rows = 2
            fgPurchase.Row = 1
            
            For rCount = 1 To RS.RecordCount
                fgPurchase.TextMatrix(fgPurchase.Row, 1) = RS(0)                   'Prod Name
                fgPurchase.TextMatrix(fgPurchase.Row, 2) = Val(RS(1))              'Qty
                fgPurchase.TextMatrix(fgPurchase.Row, 3) = Val(RS(2))              'Price
                fgPurchase.TextMatrix(fgPurchase.Row, 4) = Val(RS(1)) * Val(RS(2)) 'Amount
                fgPurchase.TextMatrix(fgPurchase.Row, 5) = Val(RS(3))              'Prod Code
                
                fgPurchase.Rows = fgPurchase.Rows + 1
                fgPurchase.Row = fgPurchase.Row + 1
                                
                RS.MoveNext
            Next

    picFind.Visible = False
    Dtp1.SetFocus

End Sub

Public Sub MouseNormalOnLbl()
'Commands Lable
    lbl(0).ForeColor = vbWhite
    lbl(0).Font.Underline = False
    
    lbl(1).ForeColor = vbWhite
    lbl(1).Font.Underline = False
    
    lbl(2).ForeColor = vbWhite
    lbl(2).Font.Underline = False
    
    lbl(3).ForeColor = vbWhite
    lbl(3).Font.Underline = False
    
    lbl(4).ForeColor = vbWhite
    lbl(4).Font.Underline = False
    
    lbl(5).ForeColor = vbWhite
    lbl(5).Font.Underline = False
    
'Navigation Label
    lblnav(0).ForeColor = vbWhite
    lblnav(0).Font.Underline = False

    lblnav(1).ForeColor = vbWhite
    lblnav(1).Font.Underline = False
    
    lblnav(2).ForeColor = vbWhite
    lblnav(2).Font.Underline = False
    
    lblnav(3).ForeColor = vbWhite
    lblnav(3).Font.Underline = False

End Sub

Public Sub GetTheAverage()
Dim CurQty As Single
Dim CurAmount As Single
Dim mTotalQty As Single
Dim mTotalAmount As Single
Dim myRow As Integer

            
        For myRow = 1 To fgPurchase.Rows - 2
            Set RsMisc = New ADODB.Recordset
            If RsMisc.State = 1 Then RsMisc.Close
                RsMisc.Open "Select Qty, Amount, ProdId from Product_Openings where ProdId = " & Val(fgPurchase.TextMatrix(myRow, 5)) & "", Con, adOpenStatic, adLockOptimistic
                    If Not IsNull(RsMisc(0)) Then
                        OpQty = Val(RsMisc(0))
                        OpAmount = Val(RsMisc(1))
                    Else
                        OpQty = 0
                        OpAmount = 0
                    End If
            
            
            CurQty = Val(fgPurchase.TextMatrix(myRow, 2))
            CurAmount = Val(fgPurchase.TextMatrix(myRow, 4))
            
            mTotalQty = Val(OpQty) + Val(CurQty)
            mTotalAmount = Val(OpAmount) + Val(CurAmount)
            
            AvgCost = Round(Val(mTotalAmount) / Val(mTotalQty), 2)
            
            Con.Execute "Update Avg_Cost set PurAvg = " & Val(AvgCost) & " Where ProdId = " & Val(fgPurchase.TextMatrix(myRow, 5)) & ""
        
        Next

End Sub
