VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTransaction 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   0  'None
   ClientHeight    =   8610
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   10695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picFind 
      Appearance      =   0  'Flat
      BackColor       =   &H00CC9933&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   3105
      ScaleHeight     =   1785
      ScaleWidth      =   4455
      TabIndex        =   9
      Top             =   3398
      Visible         =   0   'False
      Width           =   4485
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
         TabIndex        =   13
         Top             =   1260
         Width           =   1470
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
         TabIndex        =   12
         Top             =   1260
         Width           =   1470
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
         TabIndex        =   11
         Top             =   780
         Width           =   3945
      End
      Begin VB.OptionButton optId 
         Appearance      =   0  'Flat
         BackColor       =   &H00CC9933&
         Caption         =   "Find By Transaction ID"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   225
         TabIndex        =   10
         Top             =   495
         Width           =   2025
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Find Transaction"
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
         Left            =   1335
         TabIndex        =   15
         Top             =   15
         Width           =   1785
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Find Transaction"
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
         Left            =   1350
         TabIndex        =   14
         Top             =   15
         Width           =   1785
      End
      Begin VB.Image Image2 
         Height          =   270
         Index           =   1
         Left            =   0
         Picture         =   "frmTransaction.frx":0000
         Stretch         =   -1  'True
         Top             =   15
         Width           =   4455
      End
   End
   Begin VB.PictureBox PicSrchGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   7470
      Left            =   2587
      ScaleHeight     =   7440
      ScaleWidth      =   5490
      TabIndex        =   4
      Top             =   885
      Visible         =   0   'False
      Width           =   5520
      Begin VB.TextBox TxtGrdSrch 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         Height          =   330
         Left            =   105
         TabIndex        =   5
         Top             =   750
         Width           =   5310
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MshSearch 
         Height          =   6180
         Left            =   90
         TabIndex        =   6
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
         TabIndex        =   8
         Top             =   0
         Width           =   5490
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
         TabIndex        =   7
         Top             =   450
         Width           =   1635
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F0F0&
      ForeColor       =   &H80000008&
      Height          =   8565
      Left            =   0
      ScaleHeight     =   8535
      ScaleWidth      =   10650
      TabIndex        =   0
      Top             =   15
      Width           =   10680
      Begin VB.TextBox txtId 
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
         ForeColor       =   &H80000003&
         Height          =   285
         Left            =   1635
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1725
         Width           =   1500
      End
      Begin VB.TextBox txtDebit 
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
         ForeColor       =   &H80000003&
         Height          =   285
         Left            =   7185
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   7800
         Width           =   1500
      End
      Begin VB.TextBox txtCredit 
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
         ForeColor       =   &H80000003&
         Height          =   285
         Left            =   8700
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   7800
         Width           =   1500
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgTransaction 
         Height          =   5490
         Left            =   90
         TabIndex        =   1
         Top             =   2250
         Width           =   10470
         _ExtentX        =   18468
         _ExtentY        =   9684
         _Version        =   393216
         Cols            =   6
         BackColorFixed  =   13408563
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
         _Band(0).Cols   =   6
      End
      Begin MSComCtl2.DTPicker Dtp1 
         Height          =   300
         Left            =   7320
         TabIndex        =   17
         Top             =   1695
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   529
         _Version        =   393216
         Format          =   16515072
         CurrentDate     =   40158
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Press CTRL+SPACE in Account Name for List"
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
         Left            =   3180
         TabIndex        =   32
         Top             =   8310
         Width           =   3885
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T R A N S A C T I O N"
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
         Left            =   3255
         TabIndex        =   30
         Top             =   945
         Width           =   4140
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   0
         Left            =   -15
         Top             =   0
         Width           =   10680
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
         Left            =   5220
         TabIndex        =   29
         Top             =   375
         Width           =   615
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
         Left            =   8295
         TabIndex        =   28
         Top             =   375
         Width           =   330
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
         Left            =   6330
         TabIndex        =   27
         Top             =   375
         Width           =   570
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
         Left            =   7410
         TabIndex        =   26
         Top             =   375
         Width           =   375
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
         Left            =   4260
         TabIndex        =   25
         Top             =   375
         Width           =   450
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
         Left            =   3360
         TabIndex        =   24
         Top             =   375
         Width           =   390
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
         Left            =   1590
         TabIndex        =   23
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
         Left            =   600
         TabIndex        =   22
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
         Left            =   1260
         TabIndex        =   21
         Top             =   405
         Width           =   225
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
         Left            =   930
         TabIndex        =   20
         Top             =   405
         Width           =   225
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Date"
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
         Left            =   5595
         TabIndex        =   19
         Top             =   1740
         Width           =   1485
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Id"
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
         Left            =   255
         TabIndex        =   18
         Top             =   1755
         Width           =   1245
      End
      Begin VB.Image Image 
         Height          =   465
         Left            =   0
         Picture         =   "frmTransaction.frx":0E2D
         Stretch         =   -1  'True
         Top             =   240
         Width           =   10650
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T R A N S A C T I O N"
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
         Left            =   3285
         TabIndex        =   31
         Top             =   960
         Width           =   4140
      End
   End
End
Attribute VB_Name = "frmTransaction"
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

Private Sub lbl_Click(Index As Integer)
    Select Case Index
        Case 0                                      'New
            'Clearing all controls for New Data
                Clear Me
                fgTransaction.Clear
                
                GridSetting
                
                Call AutoId
                
                Dtp1.SetFocus
                Dtp1.Value = Now
                        
                Modes True, False, Me
            
            'Lock Navigation
                LockNav Me
        
        Case 1                                      'Save
                Dim FgTransactionRow As Integer
                
                'Checking Data to Save
                    If Val(fgTransaction.Rows) < 3 Then
                        MsgBox "No data to save", vbInformation, "Message"
                        Exit Sub
                    End If
                    
                'Validation for Equal Balances of Debit and Credit Side
                    If Not Val(txtDebit) = Val(txtCredit) Then
                        MsgBox "Debit and Credit sides should be equal", vbCritical, "Message"
                        Exit Sub
                    End If
                    
                    
                'If New Transaction
                    If lbl(0).Enabled = False Then
                    
                    'Getting Maximum Code for Transaction Main
                                
                        MaxNumber "TransId", "Max_Codes"
                        TranMainCode = Val(MaxNmbr)
                        
                        TransactionRef = "Voucher-" & Val(txtId)
                    
                    'Inserting Data to TransactionMain Table
                            
                            Con.Execute "insert into Transaction_Main (TransId,TransDate,TransType,Posted,TransRef) values (" & Val(txtId) & ", '" & Format(Dtp1.Value, "mm/dd/yyyy") & "' , '" & TransactionType & "' ,'" & "N" & "', '" & TransactionRef & "')"
                            
                    '----------------------------- Inserting to Transaction Detail -----------------------------
                        
                            For FgTransactionRow = 1 To fgTransaction.Rows - 2
                    
                    'Getting Maximum Code for transaction_Detail
                        MaxNumber "TransDetId", "Max_Codes"
                        TransactionId = Val(MaxNmbr)
                                mdescript = fgTransaction.TextMatrix(FgTransactionRow, 2) + " V/No " & Val(txtId)
                                Con.Execute "Insert into Transaction_Detail (TranDetId, TranId, AcId, Descript, DrAmount, CrAmount) values (" & Val(TransactionId) & ", " & Val(txtId) & ", " & Val(fgTransaction.TextMatrix(FgTransactionRow, 5)) & ", '" & mdescript & "', " & Val(fgTransaction.TextMatrix(FgTransactionRow, 3)) & ", " & Val(fgTransaction.TextMatrix(FgTransactionRow, 4)) & ")"
                           
                    'Updating MaxCode for Transaction Detail
                        UpdateMaxNumber "TransDetId", Val(TransactionId)
                    
                            Next
                    
                    'Updating Transaction Main Id
                        UpdateMaxNumber "TransId", Val(txtId)
                    
                    MsgBox "Record saved", vbInformation, "Done"
                    lbl_Click (2)
                    
                    Else
                'If existing record
                        TranMainCode = Val(txtId)
                    
                'Deleting old records from TransactionMain and Transaction Detail
                        Con.Execute "Delete from Transaction_Detail where TranId = " & Val(txtId) & ""
                        Con.Execute "Delete from Transaction_Main where TransId = " & Val(txtId) & ""
                    
                'Checking Data to Save
                    If Val(fgTransaction.Rows) < 3 Then
                        MsgBox "No data to save", vbInformation, "Message"
                        Exit Sub
                    End If
                    
                'Validation for Equal Balances of Debit and Credit Side
                    If Not Val(txtDebit) = Val(txtCredit) Then
                        MsgBox "Debit and Credit sides should be equal", vbCritical, "Message"
                        Exit Sub
                    End If
                        
                    'Inserting Data to TransactionMain Table
                            
                         Con.Execute "insert into Transaction_Main (TransId,TransDate, TransType, Posted,TransRef) values (" & Val(txtId) & ", '" & Format(Dtp1.Value, "mm/dd/yyyy") & "' , '" & TransactionType & "' ,'" & "N" & "', '" & TransactionRef & "')"
                            
                    '----------------------------- Inserting to Transaction Detail -----------------------------
                        
                            For FgTransactionRow = 1 To fgTransaction.Rows - 2
                    
                    'Getting Maximum Code for transaction_Detail
                        MaxNumber "TransDetId", "Max_Codes"
                        TransactionId = Val(MaxNmbr)
                                mdescript = fgTransaction.TextMatrix(FgTransactionRow, 2) + " V/No " & Val(txtId)
                                Con.Execute "Insert into Transaction_Detail (TranDetId, TranId, AcId, Descript, DrAmount, CrAmount) values (" & Val(TransactionId) & ", " & Val(txtId) & ", " & Val(fgTransaction.TextMatrix(FgTransactionRow, 5)) & ", '" & mdescript & "', " & Val(fgTransaction.TextMatrix(FgTransactionRow, 3)) & ", " & Val(fgTransaction.TextMatrix(FgTransactionRow, 4)) & ")"
                           
                    'Updating MaxCode for Transaction Detail
                        UpdateMaxNumber "TransDetId", Val(TransactionId)
                    
                            Next
                    
                    MsgBox "Record updated", vbInformation, "Done"
                    lbl_Click (2)
                    
                    
                    End If
        
        Case 2                                                              'Cancel
                    Clear Me
                    
                'UnLock Navigation
                    UnLockNav Me
                    
                    fgTransaction.Clear
                    
                    GridSetting
                            
                    Call ExistData
                    
                    Modes False, True, Me
                    
                    Dtp1.SetFocus
                    
                    If RsNAV.RecordCount <= 0 Then
                        Exit Sub
                    Else
                        RsNAV.Requery
                        RsNAV.MoveFirst
                    End If
                
                   
        Case 3                                                              'Delete
                     If MsgBox("Do you want to delete this Complete Voucher?", vbQuestion + vbYesNo, "Delete Voucher") = vbYes Then
                        If MsgBox("System will be unable to recover the loss data. Continue ?", vbQuestion + vbYesNo, "Delete Voucher") = vbYes Then
                'Deleting data
                            Con.Execute "Delete from transaction_Detail where TranId = " & Val(txtId) & ""
                            Con.Execute "Delete from transaction_Main where TransId = " & Val(txtId) & ""
                            MsgBox "Voucher Deleted", vbInformation, "Message"
                            lbl_Click (2)
                        End If
                    End If
               
                
        Case 4                                                              'Find
                    picFind.Visible = True
                       optId.Value = 1
                       txtFindId.SetFocus
       
        Case 5                                                              'Exit
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

Private Sub Image_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseNormalOnLbl
End Sub

Private Sub lbl_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 0, 1, 2, 3, 4, 5
            MouseOver lbl(Index)
            
    End Select

End Sub

Private Sub cmdFind_Click()
    If txtFindId.Text = "" Or Val(txtFindId) = 0 Then
        MsgBox "Enter Transaction ID", vbCritical, "Message.."
        txtFindId.SetFocus
        txtFindId.Text = ""
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
        fgTransaction.SetFocus
    End If
End Sub

Private Sub fgTransaction_Click()
    RowSelect = fgTransaction.RowSel
End Sub

Private Sub fgTransaction_KeyDown(KeyCode As Integer, Shift As Integer)
'If CTRL + SPACE is pressed
    If fgTransaction.Col = 1 Then
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
        
        If fgTransaction.Rows > 2 And fgTransaction.TextMatrix(fgTransaction.Row, 1) <> "" Then
            If MsgBox("Do you want to delete this line>", vbQuestion + vbYesNo, "Delete Line") = vbYes Then
                txtDebit.Text = Val(txtDebit) - Val(fgTransaction.TextMatrix(fgTransaction.Row, 3))
                txtCredit.Text = Val(txtCredit) - Val(fgTransaction.TextMatrix(fgTransaction.Row, 4))
                        
                fgTransaction.RemoveItem RowSelect
            End If
            
        Else
            MsgBox "Blank or Last line can not be deleted", vbCritical, "Message "
            Exit Sub
        End If
        
    End If
End Sub

Private Sub fgTransaction_KeyPress(KeyAscii As Integer)
    EditGrid fgTransaction, KeyAscii
End Sub

Private Sub fgTransaction_LeaveCell()
'========================================= W O R K I N G ==============================

Dim RowNum As Integer
Dim mDebit As Single
Dim mCredit As Single

    Select Case fgTransaction.Col
    
        Case 1
            txtDebit.Text = ""
            txtCredit.Text = ""
                For RowNum = 1 To fgTransaction.Rows - 1
                    mDebit = Val(mDebit) + Val(fgTransaction.TextMatrix(RowNum, 3))
                    mCredit = Val(mCredit) + Val(fgTransaction.TextMatrix(RowNum, 4))
                Next
            
                txtDebit.Text = mDebit
                txtCredit.Text = mCredit
        
    End Select

End Sub

Private Sub Form_Load()

'Setting up flexgrid data
    Call GridSetting

'Calling Exist Data
    Call ExistData

 'Calling Data for navigation
    Set RsNAV = New ADODB.Recordset
    If RsNAV.State = 1 Then RsNAV.Close
    RsNAV.Open "Select * from Transaction_Main Where TransType = '" & "Voucher" & "' Order By TransId", Con, adOpenStatic, adLockOptimistic

'Defining Transaction Type
    TransactionType = "Voucher"

End Sub

Private Sub Form_Resize()
    Me.Left = 3600
    Me.Top = 1400


End Sub


Public Sub GridSetting()
'Setting of Transaction Grid
    With fgTransaction
        
        .ColWidth(0) = 250
        .ColWidth(1) = 2500
        .ColWidth(2) = 4400
        .ColWidth(3) = 1475
        .ColWidth(4) = 1475
        .ColWidth(5) = 0
        .ColWidth(6) = 0
    
        .RowHeight(0) = 400
    
        .Rows = 2
        
        .TextMatrix(0, 1) = "Account Name"
        .TextMatrix(0, 2) = "Description"
        .TextMatrix(0, 3) = "Debit"
        .TextMatrix(0, 4) = "Credit"
        
        .ColAlignmentFixed(1) = 4
        .ColAlignmentFixed(2) = 4
        .ColAlignmentFixed(3) = 4
        .ColAlignmentFixed(4) = 4
    
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
    SQLQry = "Select AcId, AcTitle, AcType from ViewHeadWise"
    
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
        
        RS.Open SQLQry, Con, adOpenStatic, adLockReadOnly
            
        Set MshSearch.DataSource = RS
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseNormalOnLbl
End Sub

Private Sub txtFindId_KeyPress(KeyAscii As Integer)
    ONU KeyAscii, txtFindId
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
        
        SQLQry = "Select AcId, AcTitle, AcType from ViewHeadWise where AcTitle Like '" & TxtGrdSrch.Text & "%'"
            
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
        
        fgTransaction.TextMatrix(fgTransaction.Row, 1) = MshSearch.TextMatrix(MshSearch.Row, 1)
        fgTransaction.TextMatrix(fgTransaction.Row, 5) = MshSearch.TextMatrix(MshSearch.Row, 0)
        
        PicSrchGrid.Visible = False
        fgTransaction.Col = 1
        fgTransaction.SetFocus
End Sub

Public Sub CheckDuplicate()
Dim dRow As Integer
   
    For dRow = 1 To fgTransaction.Rows - 2
        If fgTransaction.TextMatrix(dRow, 5) = MshSearch.TextMatrix(MshSearch.Row, 0) Then
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
        fgTransaction.SetFocus
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
        fgTransaction.SetFocus
        Exit Sub
    End If
End Sub

Public Sub ExistData()

'Getting data from Purchase
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
    
        RS.Open "SELECT TransId, TransDate, TransType from Transaction_Main Where TransType = '" & "Voucher" & "'", Con, adOpenStatic, adLockOptimistic
            
            If RS.EOF = True Then
                Exit Sub
            End If
            
            txtId.Text = Val(RS(0))
            Dtp1.Value = RS(1)
        RS.Close
            
'Getting data from TransactionDetail
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
        
       mSqlQry = "SELECT Accounts.AcTitle, Transaction_Detail.Descript, Transaction_Detail.DrAmount, Transaction_Detail.CrAmount, Accounts.AcId, Transaction_Detail.TranDetId, Transaction_Main.TransId" & _
        " FROM Transaction_Main INNER JOIN (Accounts INNER JOIN Transaction_Detail ON Accounts.AcId = Transaction_Detail.AcId) ON Transaction_Main.TransId = Transaction_Detail.TranId" & _
        " WHERE Transaction_Main.TransId = " & Val(txtId) & ""

        
        RS.Open mSqlQry, Con, adOpenStatic, adLockOptimistic
            
            If RS.RecordCount <= 0 Then
                Exit Sub
            End If
            
            fgTransaction.Rows = 2
            fgTransaction.Row = 1
            
            For rCount = 1 To RS.RecordCount
                fgTransaction.TextMatrix(fgTransaction.Row, 1) = RS(0)
                fgTransaction.TextMatrix(fgTransaction.Row, 2) = RS(1)
                fgTransaction.TextMatrix(fgTransaction.Row, 3) = RS(2)
                fgTransaction.TextMatrix(fgTransaction.Row, 4) = RS(3)
                fgTransaction.TextMatrix(fgTransaction.Row, 5) = RS(4)
                
                txtDebit.Text = Val(txtDebit) + Val(RS(2))
                txtCredit.Text = Val(txtCredit) + Val(RS(3))
                
                fgTransaction.Rows = fgTransaction.Rows + 1
                fgTransaction.Row = fgTransaction.Row + 1
                                
                RS.MoveNext
            Next

End Sub

Private Sub PicFirst_Click()
    On Error Resume Next
    RsNAV.MoveFirst
    lblnav.Caption = "1"
    
    If RsNAV.BOF = True Then
        MsgBox "First Record", vbInformation, "Message"
        RsNAV.MoveFirst
        Exit Sub
    Else
        Call NAVData
    End If
End Sub

Private Sub PicLast_Click()
    On Error Resume Next
    RsNAV.MoveLast
    lblnav.Caption = Val(RsNAV.RecordCount)
    
    If RsNAV.EOF = True Then
        MsgBox "Last Record", vbInformation, "Message"
        RsNAV.MoveLast
    Else
        Call NAVData
    End If
    
End Sub

Private Sub PicNext_Click()
    On Error Resume Next
    RsNAV.MoveNext

    lblnav.Caption = Val(lblnav) + 1

    If RsNAV.EOF = True Then
        lblnav.Caption = Val(lblnav) - 1
        MsgBox "Last Record", vbInformation, "Message"
        RsNAV.MoveLast
    Else
        Call NAVData
    End If
End Sub

Private Sub PicPrev_Click()
    On Error Resume Next
    RsNAV.MovePrevious
    
    lblnav.Caption = Val(lblnav) - 1
    
    
    If RsNAV.BOF = True Then
        lblnav.Caption = Val(lblnav) + 1
        MsgBox "First Record", vbInformation, "Message"
        RsNAV.MoveFirst
    Else
        Call NAVData
    End If
    
End Sub

Public Sub NAVData()

'Getting data from Purchase
            
            txtId.Text = Val(RsNAV(0))
            Dtp1.Value = RsNAV(1)
            
'Getting data from TransactionDetail
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
        
       mSqlQry = "SELECT Accounts.AcTitle, Transaction_Detail.Descript, Transaction_Detail.DrAmount, Transaction_Detail.CrAmount, Accounts.AcId, Transaction_Detail.TranDetId, Transaction_Main.TransId" & _
        " FROM Transaction_Main INNER JOIN (Accounts INNER JOIN Transaction_Detail ON Accounts.AcId = Transaction_Detail.AcId) ON Transaction_Main.TransId = Transaction_Detail.TranId" & _
        " WHERE Transaction_Main.TransId = " & Val(txtId) & ""
        
        RS.Open mSqlQry, Con, adOpenStatic, adLockOptimistic
            
            If RS.RecordCount <= 0 Then
                Exit Sub
            End If
            
            fgTransaction.Rows = 2
            fgTransaction.Row = 1
            
                txtDebit.Text = ""
                txtCredit.Text = ""
            
            For rCount = 1 To RS.RecordCount
                fgTransaction.TextMatrix(fgTransaction.Row, 1) = RS(0)
                fgTransaction.TextMatrix(fgTransaction.Row, 2) = RS(1)
                fgTransaction.TextMatrix(fgTransaction.Row, 3) = RS(2)
                fgTransaction.TextMatrix(fgTransaction.Row, 4) = RS(3)
                fgTransaction.TextMatrix(fgTransaction.Row, 5) = RS(4)
                
                txtDebit.Text = Val(txtDebit) + Val(RS(2))
                txtCredit.Text = Val(txtCredit) + Val(RS(3))
                
                fgTransaction.Rows = fgTransaction.Rows + 1
                fgTransaction.Row = fgTransaction.Row + 1
                                
                RS.MoveNext
            Next
    
End Sub

Public Sub AutoId()
'Calling MaxNumber function to get Auto Id for the record
    MaxNumber "TransId", "Max_Codes"
    txtId.Text = Val(MaxNmbr)
End Sub

Public Sub FindRecord()
            mSqlQry = "Select * from Transaction_Main Where TransType = '" & "Voucher" & "' and TransId = " & Val(txtFindId) & ""
            
            Set RS = New ADODB.Recordset
            If RS.State = 1 Then RS.Close
                RS.Open mSqlQry, Con, adOpenStatic, adLockOptimistic
                    If RS.RecordCount <= 0 Then
                        MsgBox "No transaction found with this ID", vbCritical, "Message.."
                        txtFindId.SetFocus
                        Exit Sub
                    Else
                        ShowFoundData
                    End If
    
End Sub


Public Sub ShowFoundData()
'Getting data from TransactionMain
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
    
        RS.Open "Select * from Transaction_Main Where TransType = '" & "Voucher" & "' and TransId = " & Val(txtFindId) & "", Con, adOpenStatic, adLockOptimistic
            
            If RS.EOF = True Then
                Exit Sub
            End If
            
            txtId.Text = Val(RS(0))
            Dtp1.Value = RS(1)
        RS.Close
            
'Getting data from TransactionDetail
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
        
       mSqlQry = "SELECT Accounts.AcTitle, Transaction_Detail.Descript, Transaction_Detail.DrAmount, Transaction_Detail.CrAmount, Accounts.AcId, Transaction_Detail.TranDetId, Transaction_Main.TransId" & _
        " FROM Transaction_Main INNER JOIN (Accounts INNER JOIN Transaction_Detail ON Accounts.AcId = Transaction_Detail.AcId) ON Transaction_Main.TransId = Transaction_Detail.TranId" & _
        " WHERE Transaction_Main.TransId = " & Val(txtId) & ""

        
        RS.Open mSqlQry, Con, adOpenStatic, adLockOptimistic
            
            If RS.RecordCount <= 0 Then
                Exit Sub
            End If
            
            fgTransaction.Rows = 2
            fgTransaction.Row = 1
            
            For rCount = 1 To RS.RecordCount
                fgTransaction.TextMatrix(fgTransaction.Row, 1) = RS(0)
                fgTransaction.TextMatrix(fgTransaction.Row, 2) = RS(1)
                fgTransaction.TextMatrix(fgTransaction.Row, 3) = RS(2)
                fgTransaction.TextMatrix(fgTransaction.Row, 4) = RS(3)
                fgTransaction.TextMatrix(fgTransaction.Row, 5) = RS(4)
                
                txtDebit.Text = Val(txtDebit) + Val(RS(2))
                txtCredit.Text = Val(txtCredit) + Val(RS(3))
                
                fgTransaction.Rows = fgTransaction.Rows + 1
                fgTransaction.Row = fgTransaction.Row + 1
                                
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

