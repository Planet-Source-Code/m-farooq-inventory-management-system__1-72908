VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmRep 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7350
   ClientLeft      =   4410
   ClientTop       =   2325
   ClientWidth     =   8790
   ControlBox      =   0   'False
   ForeColor       =   &H00CCD7CC&
   Icon            =   "FrmReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox PicSrchGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   6210
      Left            =   2370
      ScaleHeight     =   6180
      ScaleWidth      =   5490
      TabIndex        =   28
      Top             =   585
      Visible         =   0   'False
      Width           =   5520
      Begin VB.TextBox TxtGrdSrch 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         Height          =   330
         Left            =   105
         TabIndex        =   29
         Top             =   750
         Width           =   5310
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MshSearch 
         Height          =   4935
         Left            =   90
         TabIndex        =   30
         Top             =   1185
         Width           =   5310
         _ExtentX        =   9366
         _ExtentY        =   8705
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
         TabIndex        =   32
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
         TabIndex        =   31
         Top             =   450
         Width           =   1635
      End
   End
   Begin VB.CommandButton CmdRun 
      BackColor       =   &H00DBF0EC&
      Caption         =   "&Run"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5295
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   6495
      Width           =   1215
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00DBF0EC&
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6765
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6495
      Width           =   1215
   End
   Begin VB.ListBox ListLov 
      Appearance      =   0  'Flat
      Height          =   1395
      ItemData        =   "FrmReport.frx":000C
      Left            =   5775
      List            =   "FrmReport.frx":000E
      TabIndex        =   25
      Top             =   4575
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.PictureBox PicLookup1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F0F0&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   4635
      ScaleHeight     =   705
      ScaleWidth      =   3705
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   5310
      Width           =   3735
      Begin VB.TextBox TxtLookupCode1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F3F3F3&
         DragIcon        =   "FrmReport.frx":0010
         Enabled         =   0   'False
         Height          =   285
         Left            =   150
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   300
         Width           =   1095
      End
      Begin VB.TextBox TxtLookup1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1230
         TabIndex        =   6
         Top             =   300
         Width           =   2325
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select List Of Value Column"
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
         Left            =   810
         TabIndex        =   23
         Top             =   90
         Width           =   2385
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   870
      Top             =   3570
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtDesc 
      Height          =   285
      Left            =   5865
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   6030
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.PictureBox PicLookup 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F0F0&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   4650
      ScaleHeight     =   705
      ScaleWidth      =   3705
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4290
      Width           =   3735
      Begin VB.TextBox TxtLookup 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1215
         TabIndex        =   5
         Top             =   270
         Width           =   2325
      End
      Begin VB.TextBox TxtLookupCode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F3F3F3&
         DragIcon        =   "FrmReport.frx":0452
         Enabled         =   0   'False
         Height          =   285
         Left            =   150
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label LblLookup 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select List Of Value Column"
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
         Left            =   810
         TabIndex        =   15
         Top             =   60
         Width           =   2385
      End
   End
   Begin VB.PictureBox PicDateBw 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F0F0&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   4650
      ScaleHeight     =   705
      ScaleWidth      =   3705
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3300
      Width           =   3735
      Begin MSComCtl2.DTPicker TxtFrom 
         Height          =   285
         Left            =   180
         TabIndex        =   3
         Top             =   300
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   65929219
         UpDown          =   -1  'True
         CurrentDate     =   39056
      End
      Begin MSComCtl2.DTPicker TxtTo 
         Height          =   285
         Left            =   2010
         TabIndex        =   4
         Top             =   300
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   65929219
         UpDown          =   -1  'True
         CurrentDate     =   39056
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Date"
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
         Left            =   2370
         TabIndex        =   10
         Top             =   90
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Date"
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
         Left            =   420
         TabIndex        =   9
         Top             =   90
         Width           =   885
      End
      Begin VB.Shape Shape1 
         Height          =   345
         Index           =   0
         Left            =   1980
         Top             =   270
         Width           =   1575
      End
      Begin VB.Shape Shape4 
         Height          =   345
         Left            =   150
         Top             =   270
         Width           =   1575
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   870
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReport.frx":0894
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReport.frx":0CE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReport.frx":1138
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReport.frx":1452
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReport.frx":1BA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReport.frx":22F6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TV 
      Height          =   6105
      Left            =   240
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   645
      WhatsThisHelpID =   10445
      Width           =   4230
      _ExtentX        =   7461
      _ExtentY        =   10769
      _Version        =   393217
      Indentation     =   353
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      SingleSel       =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox PicDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F0F0&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   4650
      ScaleHeight     =   705
      ScaleWidth      =   3705
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1290
      Width           =   3735
      Begin MSComCtl2.DTPicker TxtDate 
         Height          =   285
         Left            =   1110
         TabIndex        =   0
         Top             =   270
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   65929219
         UpDown          =   -1  'True
         CurrentDate     =   39056
      End
      Begin VB.Shape Shape1 
         Height          =   345
         Index           =   1
         Left            =   1080
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Left            =   1530
         TabIndex        =   17
         Top             =   60
         Width           =   420
      End
   End
   Begin VB.PictureBox PicCode 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F0F0&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   4650
      ScaleHeight     =   705
      ScaleWidth      =   3705
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2295
      Width           =   3735
      Begin VB.TextBox TxtToCode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F3F3F3&
         Height          =   285
         Left            =   2130
         TabIndex        =   2
         Top             =   300
         Width           =   1335
      End
      Begin VB.TextBox TxtCode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F3F3F3&
         Height          =   285
         Left            =   270
         TabIndex        =   1
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Code"
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
         Left            =   2340
         TabIndex        =   19
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Code"
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
         Left            =   390
         TabIndex        =   12
         Top             =   120
         Width           =   915
      End
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Report from Left and Provide the HIGHLIGHTED information"
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
      Left            =   1410
      TabIndex        =   34
      Top             =   7035
      Width           =   5730
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reports"
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
      Left            =   3855
      TabIndex        =   33
      Top             =   -15
      Width           =   1095
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Parameters"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   5700
      TabIndex        =   24
      Top             =   600
      Width           =   1845
   End
   Begin VB.Label LBLCode 
      Caption         =   "Label10"
      Height          =   315
      Left            =   2610
      TabIndex        =   20
      Top             =   4920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00404040&
      Height          =   5625
      Left            =   4530
      Top             =   600
      Width           =   4005
   End
   Begin VB.Shape Shape11 
      BorderColor     =   &H00404040&
      Height          =   6255
      Left            =   210
      Top             =   570
      Width           =   4275
   End
   Begin VB.Shape Shape13 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   465
      Left            =   4530
      Top             =   600
      Width           =   4005
   End
   Begin VB.Shape Shape12 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   345
      Left            =   4530
      Top             =   780
      Width           =   4005
   End
   Begin VB.Image Image1 
      Height          =   390
      Left            =   0
      Picture         =   "FrmReport.frx":2748
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8805
   End
End
Attribute VB_Name = "FrmRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nodbb As Node
Dim VRun As String
'------------------------------
Dim RsLov As ADODB.Recordset
Dim a As Single                 'Variable For Loop
Dim DrBal As Double             'Variable for Debit Balance
Dim CrBal As Double             'Variable for Credit Balance
Dim mDateQry As String          'Variable for storing BETWEEN DATES Query
Dim OpBal As Double             'Difference between Total DR and Total CR for Opening
Dim MaxId As Double             'Id for Ledger Report
Dim DtFrom As Date              'Date for COG (FROM)
Dim DtTo As Date                'Date for COG (TO)
'-------------------------------

Public Sub cmdExit_Click()
    If PicSrchGrid.Visible = True Then
        PicSrchGrid.Visible = False
        Exit Sub
    End If
    
    On Error Resume Next
        Unload Me
    On Error Resume Next
End Sub

Private Sub CMDexit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CmdExit.Font.Bold = True
End Sub
Private Sub CMdexit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
''    Shape1(6).BackColor = &H568C73
End Sub
Private Sub CMDexit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CmdExit.Font.Bold = False
End Sub

Private Sub CmdLovCancel2_Click()
    Select Case VRun
        Case "Accounts Status"
            TxtLookup1.SetFocus
       Case "Daily Collection List"
            TxtDate.SetFocus
    End Select
    PicSrchGrid.Visible = False
    If RsLov.State = 1 Then RsLov.Close
End Sub
Private Sub CmdLovOK2_Click()
    Select Case VRun
        Case "Accounts Ledger"
            If MshSearch.Rows <= 1 Then CmdLovCancel2_Click: Exit Sub
            TxtLookupCode.Text = MshSearch.TextMatrix(MshSearch.Row, 0)
            TxtLookup.Text = MshSearch.TextMatrix(MshSearch.Row, 1)
            
        Case "Account Wise Receipts"
            If MshSearch.Rows <= 1 Then CmdLovCancel2_Click: Exit Sub
            TxtLookupCode.Text = MshSearch.TextMatrix(MshSearch.Row, 1)
            TxtLookup.Text = MshSearch.TextMatrix(MshSearch.Row, 0)
        
        Case "Account Wise Payments"
            If MshSearch.Rows <= 1 Then CmdLovCancel2_Click: Exit Sub
            TxtLookupCode.Text = MshSearch.TextMatrix(MshSearch.Row, 1)
            TxtLookup.Text = MshSearch.TextMatrix(MshSearch.Row, 0)
            
        Case "Supplier Wise Purchase History"
            If MshSearch.Rows <= 1 Then CmdLovCancel2_Click: Exit Sub
            TxtLookupCode.Text = MshSearch.TextMatrix(MshSearch.Row, 0)
            TxtLookup.Text = MshSearch.TextMatrix(MshSearch.Row, 1)
            
        Case "Product Wise Purchase History"
            If MshSearch.Rows <= 1 Then CmdLovCancel2_Click: Exit Sub
            TxtLookupCode.Text = MshSearch.TextMatrix(MshSearch.Row, 0)
            TxtLookup.Text = MshSearch.TextMatrix(MshSearch.Row, 1)
         
        Case "Product Wise Sale History"
            If MshSearch.Rows <= 1 Then CmdLovCancel2_Click: Exit Sub
            TxtLookupCode.Text = MshSearch.TextMatrix(MshSearch.Row, 0)
            TxtLookup.Text = MshSearch.TextMatrix(MshSearch.Row, 1)
         
         Case "Customer Wise Sale History"
            If MshSearch.Rows <= 1 Then CmdLovCancel2_Click: Exit Sub
            TxtLookupCode.Text = MshSearch.TextMatrix(MshSearch.Row, 0)
            TxtLookup.Text = MshSearch.TextMatrix(MshSearch.Row, 1)
           
        Case "Month Wise Account Status"
            If MshSearch.Rows <= 1 Then CmdLovCancel2_Click: Exit Sub
            TxtLookupCode.Text = MshSearch.TextMatrix(MshSearch.Row, 1)
            TxtLookup.Text = MshSearch.TextMatrix(MshSearch.Row, 0)
            TxtLookup.SetFocus
        
        Case "Accounts Status"
            If MshSearch.Rows <= 1 Then CmdLovCancel2_Click: Exit Sub
            TxtLookupCode1.Text = MshSearch.TextMatrix(MshSearch.Row, 0)
            TxtLookup1.Text = MshSearch.TextMatrix(MshSearch.Row, 1)
        
        Case "Receivable"
            If MshSearch.Rows <= 1 Then CmdLovCancel2_Click: Exit Sub
            TxtLookupCode1.Text = MshSearch.TextMatrix(MshSearch.Row, 0)
            TxtLookup1.Text = MshSearch.TextMatrix(MshSearch.Row, 1)
        
        Case "Payable"
            If MshSearch.Rows <= 1 Then CmdLovCancel2_Click: Exit Sub
            TxtLookupCode1.Text = MshSearch.TextMatrix(MshSearch.Row, 0)
            TxtLookup1.Text = MshSearch.TextMatrix(MshSearch.Row, 1)
        
        
    End Select
    CmdLovCancel2_Click
End Sub
Private Sub CmdRun_Click()
    
    If PicDateBw.Enabled = True Then
        If TxtFrom.Value > TxtTo.Value Then
            MsgBox "From Date value must be less than To Date value", vbInformation, "Date error"
            TxtFrom.SetFocus
            Exit Sub
        End If
    End If
    
    
    DtFrom = TxtFrom.Value
    DtTo = TxtTo.Value
    
    
    mDateQry = " BETWEEN #" & DtFrom & "# And #" & DtTo & "#"
    
    
    Select Case VRun
        Case "Accounts Status"
            
            If TxtLookupCode1.Text = "" Then
                
                If TxtLookup.Text = "ALL" Then
                    StrQry = "Select AcId, AcTitle, AcType, Balance from ViewAccountStatus Where AcType NOT IN ('Product') order By AcId"
                    dataRepAccountStatus.Show vbModal
                    
                Else
                    StrQry = "Select AcId, AcTitle, AcType, Balance from ViewAccountStatus where AcType = '" & TxtLookup.Text & "' order By AcId"
                    dataRepAccountStatus.Show vbModal
                
                End If
            
            Else
                    StrQry = "Select AcId, AcTitle, AcType, Balance from ViewAccountStatus where AcType = '" & TxtLookup.Text & "' and AcId = " & Val(TxtLookupCode1) & " order By AcId"
                    dataRepAccountStatus.Show vbModal
            
            End If

                   
        Case "Accounts Ledger"
            Call AccLedger
        
        
        Case "Stock Status"
            Call ProductStock
        
        Case "Reorder Status"
            Call ReOrderStatus
            
        Case "Products To Reorder"
            Call ProductsToReorder
            
        
''        Case "Receivable"
''            If TxtLookupCode1.Text = "" Then
''
''                If TxtLookup.Text = "ALL" Then
''                    StrQry = "Select Code,Remarks from Accounts order By code"
''                    RepBalances.Sections("ReportHeader").Controls("LblRepType").Caption = "ALL BALANCE"
''                Else
''                    StrQry = "Select Code,Remarks from Accounts where AccType = '" & TxtLookup.Text & "' order By code"
''                    RepBalances.Sections("ReportHeader").Controls("LblRepType").Caption = UCase(TxtLookup.Text) & " BALANCE"
''                End If
''
''            Else
''                    StrQry = "Select Code,Remarks from Accounts where AccType = '" & TxtLookup.Text & "' and Code = " & Val(TxtLookupCode1) & " order By code"
''            End If
''
''            Call BalanceRep
        
        
        Case "Date Wise Purchase History"
            Call PurchaseHistoryByDate
        
        Case "Supplier Wise Purchase History"
            Call PurchaseHistoryBySupplier
        
        Case "Product Wise Purchase History"
            Call PurchaseHistoryByProduct
        
        Case "Date Wise Sale History"
            Call DateWiseSaleHistory
        
        Case "Customer Wise Sale History"
            Call CustomerWiseSaleHistory
        
        Case "Product Wise Sale History"
            Call SaleHistoryByProduct
        
        Case "Product Price List"
            Call ProductPriceList

        Case "Trial Balance"
            Call TrialBalanceData
            DataRepTrialBalance.Show vbModal
            
        Case "Balance Sheet"
            
            If TxtLookup.Text = "" Then
                MsgBox "Select Month using Ctrl+Space", vbCritical, "Message"
                TxtLookup.SetFocus
                Exit Sub
            End If
                
            If TxtLookup1.Text = "" Then
                MsgBox "Enter YEAR for Balance Sheet", vbCritical, "Message..."
                TxtLookup1.SetFocus
                Exit Sub
            End If
            
            Call MonthDates


''        Case "Stock Adjustment"
''            Call StockAdjustment
        
''        Case "Payable"
''            If TxtLookupCode1.Text = "" Then
''
''                If TxtLookup.Text = "ALL" Then
''                    StrQry = "Select Code,Remarks from Accounts order By code"
''                    RepBalances.Sections("ReportHeader").Controls("LblRepType").Caption = "ALL BALANCE"
''                Else
''                    StrQry = "Select Code,Remarks from Accounts where AccType = '" & TxtLookup.Text & "' order By code"
''                    RepBalances.Sections("ReportHeader").Controls("LblRepType").Caption = UCase(TxtLookup.Text) & " BALANCE"
''                End If
''
''            Else
''                    StrQry = "Select Code,Remarks from Accounts where AccType = '" & TxtLookup.Text & "' and Code = " & Val(TxtLookupCode1) & " order By code"
''            End If
''
''            Call BalanceRep
''
''        Case "Cost Of Goods Sold"
''
''            If TxtLookup.Text = "" Then
''                MsgBox "Select Month using Ctrl+Space", vbCritical, "Message"
''                TxtLookup.SetFocus
''                Exit Sub
''            End If
''
''            If TxtLookup1.Text = "" Then
''                MsgBox "Enter YEAR for COG", vbCritical, "Message..."
''                TxtLookup1.SetFocus
''                Exit Sub
''            End If
''
''            Call MonthDates

             
        Case "Income Statement"
                Call IncomeStatData
            
''        Case "Date Wise Receipts"
''            If DeRep.rsCashReceipts.State = 1 Then DeRep.rsCashReceipts.Close
''
''
''
''                DeRep.CashReceipts TxtFrom.Value, TxtTo.Value
''                RepCashReceipts.Sections("ReportHeader").Controls("LblDate").Caption = "From :  " & TxtFrom.Value & "    " & "To :  " & TxtTo.Value
''                Load RepCashReceipts
''                RepCashReceipts.Show 1
''
''        Case "Account Wise Receipts"
''            If DeRep.rsCashReceiptsByAccount.State = 1 Then DeRep.rsCashReceiptsByAccount.Close
''
''
''
''                DeRep.CashReceiptsByAccount TxtFrom.Value, TxtTo.Value, Val(TxtLookupCode)
''                RepCashReceiptsByAccount.Sections("ReportHeader").Controls("LblDate").Caption = "From :  " & TxtFrom.Value & "    " & "To :  " & TxtTo.Value
''                Load RepCashReceiptsByAccount
''                RepCashReceiptsByAccount.Show 1
''
''        Case "Date Wise Payments"
''            If DeRep.rsCashPayment.State = 1 Then DeRep.rsCashPayment.Close
''
''
''
''                DeRep.CashPayment TxtFrom.Value, TxtTo.Value
''                RepCashPayments.Sections("ReportHeader").Controls("LblDate").Caption = "From :  " & TxtFrom.Value & "    " & "To :  " & TxtTo.Value
''                Load RepCashPayments
''                RepCashPayments.Show 1
''
''        Case "Account Wise Payments"
''            If DeRep.rsCashPaymentByAccount.State = 1 Then DeRep.rsCashPaymentByAccount.Close
''
''
''
''                DeRep.CashPaymentByAccount TxtFrom.Value, TxtTo.Value, Val(TxtLookupCode)
''                RepCashReceiptsByAccount.Sections("ReportHeader").Controls("LblDate").Caption = "From :  " & TxtFrom.Value & "    " & "To :  " & TxtTo.Value
''                Load RepCashPaymentByAccount
''                RepCashPaymentByAccount.Show 1
''
''        Case "Department Vehicles"
''            If DeRep.rsVehicleList.State = 1 Then DeRep.rsVehicleList.Close
''
''
''
''                DeRep.VehicleList Val(TxtLookupCode)
''                RepVehicles.Sections("Section4").Controls("LblDept").Caption = "Department :  " & TxtLookup.Text
''                Load RepVehicles
''                RepVehicles.Show 1
''
''        Case "Issues"
''            If TxtLookup.Text = "" Then
''                MsgBox "Enter meter number", vbInformation, "Message"
''                TxtLookup.SetFocus
''                Exit Sub
''            End If
''
''            If Not IsNumeric(TxtLookup) Then
''                MsgBox "Enter digits only", vbInformation, "Message"
''                TxtLookup.SetFocus
''                Exit Sub
''            End If
''
''            If DeRep.rsMeterIssue.State = 1 Then DeRep.rsMeterIssue.Close
''
''
''
''            DeRep.MeterIssue TxtFrom.Value, TxtTo.Value, Val(TxtLookup)
''            RepMeterIssue.Sections("ReportHeader").Controls("LblDate").Caption = "From :  " & TxtFrom.Value & "     " & "To :  " & TxtTo.Value
''            Load RepMeterIssue
''            RepMeterIssue.Show 1
''
''        Case "Reading"
''            If TxtLookup.Text = "" Then
''                MsgBox "Enter meter number", vbInformation, "Message"
''                TxtLookup.SetFocus
''                Exit Sub
''            End If
''
''            If Not IsNumeric(TxtLookup) Then
''                MsgBox "Enter digits only", vbInformation, "Message"
''                TxtLookup.SetFocus
''                Exit Sub
''            End If
''
''            If DeRep.rsMeterReading.State = 1 Then DeRep.rsMeterReading.Close
''
''
''
''            DeRep.MeterReading TxtFrom.Value, TxtTo.Value, Val(TxtLookup)
''            RepMeterReading.Sections("ReportHeader").Controls("LblDate").Caption = "From :  " & TxtFrom.Value & "     " & "To :  " & TxtTo.Value
''            RepMeterReading.Sections("ReportHeader").Controls("LblMeterCode").Caption = "Meter Number : " & Val(TxtLookup)
''
''            Load RepMeterReading
''            RepMeterReading.Show 1
''
''        Case "Meter Stock"
''            Call MeterStock
''            If DeRep.rsMeterStock.State = 1 Then DeRep.rsMeterStock.Close
''
''
''
''                DeRep.MeterStock
''                DeRep.rsMeterStock.Requery
''                RepMeterStock.Sections("ReportHeader").Controls("LblDate").Caption = Format(Now, "dd/mmm/yyyy - HH:MM")
''
''                Load RepMeterStock
''                RepMeterStock.Show 1
''
''        Case "Code Wise Remaining"
''            If DeRep.rsReminderMain.State = 1 Then DeRep.rsReminderMain.Close
''
''             DeRep.ReminderMain TxtLookup.Text
''                Load RepReminderByCode
''                RepReminderByCode.Show 1
        
        
''            Call COGData
''            Call IncomeStatement
''            Call BalanceSheet
''            Load RepIncomeStatement
''            RepIncomeStatement.Show
        
    End Select
End Sub

Private Sub CmdRun_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CmdRun.Font.Bold = True
End Sub
Private Sub CmdRun_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
''    Shape1(7).BackColor = &H568C73
End Sub
Private Sub CmdRun_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CmdRun.Font.Bold = False
End Sub


Private Sub Command1_Click()

    dataRepIncomeStat.Show vbModal

'''''Dim NetSale As Single
'''''Dim BegInv As Single
'''''Dim Purchase As Single
'''''Dim AvblForSale As Single
'''''Dim EndingInv As Single
'''''Dim COGSold As Single
'''''Dim Gross As Single
'''''Dim OtherIncome As Single
'''''Dim Expense As Single
'''''Dim NetProfit As Single
'''''
'''''Dim mDateQry As String
'''''
'''''
'''''    NetSale = 0
'''''    BegInv = 0
'''''    Purchase = 0
'''''    AvblForSale = 0
'''''    EndingInv = 0
'''''    COGSold = 0
'''''    Gross = 0
'''''    OtherIncome = 0
'''''    Expense = 0
'''''    NetProfit = 0
'''''
'''''
'''''    mDateQry = " BETWEEN #" & FrmRep.TxtFrom.Value & "# And #" & FrmRep.TxtTo.Value & "#"
'''''
'''''
'''''' NET SALES OF THE DEFINED PERIOD
'''''    Set RS = New ADODB.Recordset
'''''    If RS.State = 1 Then RS.Close
'''''        RS.Open "Select Sum(TotalAmount) from Sale_Main where SaleDate" & mDateQry, Con, adOpenStatic, adLockOptimistic
'''''            If IsNull(RS(0)) Then
'''''                NetSale = 0
'''''            Else
'''''                NetSale = Val(RS(0))
'''''            End If
'''''
'''''' OPENING INVENTORY BEFORE DEFINED PERIOD
'''''    Dim ProductOpenings As Single
'''''    Dim OpQty As Single
'''''    Dim OpAvgPrice As Single
'''''    Dim BeforePeriod As String
'''''    Dim PurValueBefore As Single
'''''    Dim SalValueBefore As Single
'''''
'''''    BeforePeriod = " < #" & FrmRep.TxtFrom.Value & "#"
'''''
'''''    'Value of Product Openings
'''''        Set RS = New ADODB.Recordset
'''''        If RS.State = 1 Then RS.Close
'''''            RS.Open "Select Sum(Amount), Sum(Qty) from Product_Openings", Con, adOpenStatic, adLockOptimistic
'''''                If IsNull(RS(0)) Then
'''''                    ProductOpenings = 0
'''''                Else
'''''                    ProductOpenings = Val(RS(0))
'''''                End If
'''''
'''''                If IsNull(RS(1)) Then
'''''                    OpQty = 0
'''''                Else
'''''                    OpQty = Val(RS(1))
'''''                End If
'''''
'''''            On Error Resume Next
'''''                OpAvgPrice = Val(RS(0)) / Val(RS(1))
'''''            On Error Resume Next
'''''
'''''    'Value of Purchase before given period
'''''        Set RS = New ADODB.Recordset
'''''        If RS.State = 1 Then RS.Close
'''''            RS.Open "Select Sum(TotalAmount) from Purchase_Main where PurDate " & BeforePeriod, Con, adOpenStatic, adLockOptimistic
'''''                If IsNull(RS(0)) Then
'''''                    PurValueBefore = 0
'''''                Else
'''''                    PurValueBefore = Val(RS(0))
'''''                End If
'''''
'''''
'''''    'Value of Sale before given period
'''''        Set RS = New ADODB.Recordset
'''''        If RS.State = 1 Then RS.Close
'''''            RS.Open "Select Sum(TotalAmount) from Sale_Main where SaleDate " & BeforePeriod, Con, adOpenStatic, adLockOptimistic
'''''                If IsNull(RS(0)) Then
'''''                    SalValueBefore = 0
'''''                Else
'''''                    SalValueBefore = Val(RS(0))
'''''                End If
'''''
'''''
'''''    'Calculating Begining Inventory
'''''        BegInv = Val(ProductOpenings) + Val(PurValueBefore) - Val(SalValueBefore)
'''''
'''''
'''''' NET PURCHASES OF THE DEFINED PERIOD
'''''    Set RS = New ADODB.Recordset
'''''    If RS.State = 1 Then RS.Close
'''''        RS.Open "Select Sum(TotalAmount) from Purchase_Main where PurDate" & mDateQry, Con, adOpenStatic, adLockOptimistic
'''''            If IsNull(RS(0)) Then
'''''                Purchase = 0
'''''            Else
'''''                Purchase = Val(RS(0))
'''''            End If
'''''
'''''' CALCULATING VALUE OF GOODS AVAILABALE FOR SALE
'''''        AvblForSale = Val(BegInv) + Val(Purchase)
'''''
'''''
'''''' ENDING INVENTORY OF DEFINED PERIOD
'''''
'''''    Dim AfterPeriod As String
'''''    Dim PurValueAfter As Single
'''''    Dim SalValue As Single
'''''    Dim AvgCost As Single
'''''    Dim SaleQty As Single
'''''
'''''    AfterPeriod = "#" & FrmRep.TxtTo.Value & "#"
'''''
'''''    'Value of Product Openings
'''''        'Taken from Variable PRODUCTOPENINGS
'''''
'''''    'Value of Purchase till Last Date
'''''        Set RS = New ADODB.Recordset
'''''        If RS.State = 1 Then RS.Close
'''''            RS.Open "Select Sum(TotalAmount) from Purchase_Main where PurDate " & AfterPeriod, Con, adOpenStatic, adLockOptimistic
'''''                If IsNull(RS(0)) Then
'''''                    PurValueAfter = 0
'''''                Else
'''''                    PurValueAfter = Val(RS(0))
'''''                End If
'''''
'''''        'Avg Purchase Cost till last date
'''''            Set RS = New ADODB.Recordset
'''''            If RS.State = 1 Then RS.Close
'''''                RS.Open "SELECT Sum(Purchase_Detail.Qty) AS TotalQty, Purchase_Main.TotalAmount AS TotalAmount, TotalAmount/TotalQty AS Cost FROM Purchase_Detail INNER JOIN Purchase_Main ON Purchase_Detail.PurId = Purchase_Main.PurId GROUP BY Purchase_Main.TotalAmount, Purchase_Detail.Qty where Purchase_Main.PurDate " & AfterPeriod, Con, adOpenStatic, adLockOptimistic
'''''                    If Val(Cost) > 0 Then
'''''                        AvgCost = Val(Cost) + Val(OpAvgPrice) / 2
'''''                    Else
'''''                        AvgCost = Val(OpAvgPrice)
'''''                    End If
'''''
'''''
'''''    'Cost of Sale till Last Date
'''''
'''''        Set RS = New ADODB.Recordset
'''''        If RS.State = 1 Then RS.Close
'''''            RS.Open "SELECT Sale_Main.SaleDate, Sum(Sale_Detail.Qty) AS SumOfQty FROM Sale_Main INNER JOIN Sale_Detail ON Sale_Main.SaleId = Sale_Detail.SaleId GROUP BY Sale_Main.SaleDate, Sale_Detail.Qty HAVING Sale_Main.SaleDate <= " & AfterPeriod, Con, adOpenStatic, adLockOptimistic
'''''
'''''                If IsNull(RS(1)) Then
'''''                    SaleQty = 0
'''''                Else
'''''                    SaleQty = Val(RS(1))
'''''                    SalValue = Val(AvgCost) * Val(SaleQty)
'''''                End If
'''''
'''''    'Calculating Ending Inventory
'''''        EndingInv = Val(ProductOpenings) + Val(PurValueAfter) - Val(SalValue)
'''''        If Val(EndingInv) < 0 Then
'''''            EndingInv = 0
'''''        End If
'''''
'''''' COST OF GOODS SOLD
'''''        COGSold = Val(AvblForSale) - Val(EndingInv)
'''''
'''''' GROSS PROFIT
'''''        Gross = Val(NetSale) - Val(COGSold)
'''''


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then           'F5
        CmdRun_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        cmdExit_Click
    End If
End Sub

Private Sub Form_Load()
    Me.Left = Me.Left
    PicDateBw.Enabled = False
    PicCode.Enabled = False
    PicDate.Enabled = False
    PicLookup.Enabled = False
    PicLookup.BackColor = vbWhite
    PicDate.BackColor = vbWhite
    PicDateBw.BackColor = vbWhite
    PicCode.BackColor = vbWhite
    PicLookup1.Enabled = False
    PicLookup1.BackColor = vbWhite

    TxtDate.Day = Day(Date)
    TxtDate.Month = Month(Date)
    TxtDate.Year = Year(Date)
    TxtFrom.Day = Day(Date)
    TxtFrom.Month = Month(Date)
    TxtFrom.Year = Year(Date)
    TxtTo.Day = Day(Date)
    TxtTo.Month = Month(Date)
    TxtTo.Year = Year(Date)

    Set nodbb = TV.Nodes.Add(, , "Root", "Reports", 3) 'Report
    nodbb.Expanded = True
'        ==============[PURCHASE]===============
        TV.Nodes.Add "Root", tvwChild, "A", "Purchase", 6
            TV.Nodes.Add "A", tvwChild, "A1", "Date Wise Purchase History", 1
            TV.Nodes.Add "A", tvwChild, "A2", "Supplier Wise Purchase History", 1
            TV.Nodes.Add "A", tvwChild, "A3", "Product Wise Purchase History", 1
            TV.Nodes.Item("A").Bold = True
        
'        ==============[SALES]===============
        TV.Nodes.Add "Root", tvwChild, "B", "Sales", 6
            TV.Nodes.Add "B", tvwChild, "B2", "Date Wise Sale History", 1
            TV.Nodes.Add "B", tvwChild, "B3", "Customer Wise Sale History", 1
            TV.Nodes.Add "B", tvwChild, "B4", "Product Wise Sale History", 1
            TV.Nodes.Item("B").Bold = True
        
'        ==============[ACCOUNTS]===============
        TV.Nodes.Add "Root", tvwChild, "C", "Accounts", 6
            TV.Nodes.Add "C", tvwChild, "C1", "Accounts Status", 1
            TV.Nodes.Add "C", tvwChild, "C2", "Accounts Ledger", 1
            TV.Nodes.Item("C").Bold = True
     
'        ==============[STOCK]===============
        TV.Nodes.Add "Root", tvwChild, "D", "Stock/Product Reports", 6
            TV.Nodes.Add "D", tvwChild, "D1", "Stock Status", 1
            TV.Nodes.Add "D", tvwChild, "D2", "Product Price List", 1
            TV.Nodes.Add "D", tvwChild, "D3", "Reorder Status", 1
            TV.Nodes.Add "D", tvwChild, "D4", "Products To Reorder", 1
            TV.Nodes.Item("D").Bold = True
        
'        ==============[FINANCIALS]===============
        TV.Nodes.Add "Root", tvwChild, "E", "Financial Reports", 6
            TV.Nodes.Add "E", tvwChild, "E1", "Trial Balance", 1
            TV.Nodes.Add "E", tvwChild, "E2", "Income Statement", 1
            TV.Nodes.Item("E").Bold = True
        
'Assigning values to LISTLOV
    Call AccountType
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
''        Shape1(6).BackColor = &H96BEAE
''        Shape1(7).BackColor = &H96BEAE
End Sub

Private Sub MnuTitle_Click()

End Sub

Private Sub Form_Resize()
    Me.Left = Me.Left + 1300
End Sub

Private Sub ListLov_Click()
    TxtLookup.Text = ListLov.Text
    ListLov.Visible = False
    
    If VRun = "Accounts Status" Then
    
        If TxtLookup.Text = "ALL" Then
            PicLookup1.Enabled = False
            PicLookup1.BackColor = vbWhite
        Else
            PicLookup1.Enabled = True
            PicLookup1.BackColor = &HC0C0C0
            TxtLookup1.SetFocus
        End If
    
    End If
        
End Sub

Private Sub ListLov_KeyPress(KeyAscii As Integer)
        TxtLookup.Text = ListLov.Text
        ListLov.Visible = False
        
        If VRun = "Accounts Status" Then
        
            If TxtLookup.Text = "ALL" Then
                PicLookup1.Enabled = False
                PicLookup1.BackColor = vbWhite
                
            Else
            
            PicLookup1.Enabled = True
            PicLookup1.BackColor = &HC0C0C0
            TxtLookup1.SetFocus
    
            End If
        End If
        
        
    
End Sub

Private Sub MshSearch_DblClick()
    CmdLovOK2_Click
    TxtLookup.SetFocus
End Sub

Private Sub MshSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdLovOK2_Click
        TxtLookup.SetFocus
    End If
    If KeyAscii = 27 Then
        CmdLovCancel2_Click
    End If
    
    If KeyAscii = 8 Then
        If TxtGrdSrch.Text <> "" Then TxtGrdSrch.Text = Left$(TxtGrdSrch.Text, (Len(TxtGrdSrch.Text) - 1))
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
    Else
        TxtGrdSrch.Text = TxtGrdSrch.Text + Chr$(KeyAscii)
    End If
    
End Sub

Private Sub TV_NodeClick(ByVal Node As MSComctlLib.Node)
     '====================Purchase Report Start==========================
    If Node = "Reports" Then
        VRun = ""
        TxtLookup = ""
        TxtLookupCode = ""
        TxtLookup1 = ""
        TxtLookupCode1 = ""
        TxtCode = ""
        TxtToCode = ""
        PicDateBw.Enabled = False
        PicCode.Enabled = False
        PicDate.Enabled = False
        PicLookup.Enabled = False
        PicLookup1.Enabled = False
        PicLookup.BackColor = vbWhite
        PicLookup1.BackColor = vbWhite
        PicLookup.BackColor = vbWhite
        PicDate.BackColor = vbWhite
        PicDateBw.BackColor = vbWhite
        PicCode.BackColor = vbWhite

     '-----------------------------------------------------------
    '--------------Accounts Reports-------------------
    ElseIf Node = "Accounts Status" Then
        VRun = "Accounts Status"
        ListLov.Clear
        Call AccountType
        
        PicDateBw.Enabled = False
        PicCode.Enabled = False
        PicDate.Enabled = False
        PicLookup.Enabled = True
        PicLookup.BackColor = &HC0C0C0
        PicDate.BackColor = vbWhite
        PicDateBw.BackColor = vbWhite
        PicCode.BackColor = vbWhite
        TxtLookup.SetFocus
        PicLookup1.Enabled = True
        PicLookup1.BackColor = &HC0C0C0
    
    
    ElseIf Node = "Trial Balance" Then
        VRun = "Trial Balance"
        PicDateBw.Enabled = False
        PicCode.Enabled = False
        PicDate.Enabled = False
        PicLookup.Enabled = False
        PicLookup.BackColor = vbWhite
        PicDate.BackColor = vbWhite
        PicDateBw.BackColor = vbWhite
        PicCode.BackColor = vbWhite
        PicLookup1.Enabled = False
        PicLookup1.BackColor = vbWhite
    
    ElseIf Node = "Receivable" Then
        VRun = "Receivable"
        ListLov.Clear
        Call AccountType
        
        PicDateBw.Enabled = False
        PicCode.Enabled = False
        PicDate.Enabled = False
        PicLookup.Enabled = True
        PicLookup.BackColor = &HC0C0C0
        PicDate.BackColor = vbWhite
        PicDateBw.BackColor = vbWhite
        PicCode.BackColor = vbWhite
        PicLookup1.Enabled = True
        PicLookup1.BackColor = &HC0C0C0
        TxtLookup.SetFocus
        
    ElseIf Node = "Payable" Then
        VRun = "Payable"
        ListLov.Clear
        Call AccountType
        
        PicDateBw.Enabled = False
        PicCode.Enabled = False
        PicDate.Enabled = False
        PicLookup.Enabled = True
        PicLookup.BackColor = &HC0C0C0
        PicDate.BackColor = vbWhite
        PicDateBw.BackColor = vbWhite
        PicCode.BackColor = vbWhite
        PicLookup1.Enabled = True
        PicLookup1.BackColor = &HC0C0C0
        TxtLookup.SetFocus
    
    ElseIf Node = "Accounts Ledger" Then
        VRun = "Accounts Ledger"
        ListLov.Clear
        Call AccountType
        
        PicDateBw.Enabled = True
        PicDateBw.BackColor = &HC0C0C0
        PicCode.Enabled = False
        PicDate.Enabled = False
        PicLookup.Enabled = True
        PicLookup.BackColor = &HC0C0C0
        PicLookup1.BackColor = vbWhite
        PicLookup1.Enabled = False
        PicDate.BackColor = vbWhite
        TxtFrom.SetFocus
        
    ElseIf Node = "Income Statement" Then
        VRun = "Income Statement"
        PicDateBw.Enabled = True
        PicCode.Enabled = False
        PicDate.Enabled = False
        PicLookup.Enabled = False
        PicLookup.BackColor = vbWhite
        PicDate.BackColor = vbWhite
        PicDateBw.BackColor = &HC0C0C0
        PicCode.BackColor = vbWhite
        PicLookup1.Enabled = False
        PicLookup1.BackColor = vbWhite
        TxtFrom.SetFocus

    
    ElseIf Node = "Stock Status" Then
        VRun = "Stock Status"
        
        PicDateBw.Enabled = False
        PicCode.Enabled = False
        PicDate.Enabled = False
        PicLookup.Enabled = False
        PicLookup.BackColor = vbWhite
        PicDate.BackColor = vbWhite
        PicDateBw.BackColor = vbWhite
        PicCode.BackColor = vbWhite
        PicLookup1.Enabled = False
        PicLookup1.BackColor = vbWhite
    
    ElseIf Node = "Reorder Status" Then
        VRun = "Reorder Status"
        
        PicDateBw.Enabled = False
        PicCode.Enabled = False
        PicDate.Enabled = False
        PicLookup.Enabled = False
        PicLookup.BackColor = vbWhite
        PicDate.BackColor = vbWhite
        PicDateBw.BackColor = vbWhite
        PicCode.BackColor = vbWhite
        PicLookup1.Enabled = False
        PicLookup1.BackColor = vbWhite
    
    ElseIf Node = "Products To Reorder" Then
        VRun = "Products To Reorder"
        
        PicDateBw.Enabled = False
        PicCode.Enabled = False
        PicDate.Enabled = False
        PicLookup.Enabled = False
        PicLookup.BackColor = vbWhite
        PicDate.BackColor = vbWhite
        PicDateBw.BackColor = vbWhite
        PicCode.BackColor = vbWhite
        PicLookup1.Enabled = False
        PicLookup1.BackColor = vbWhite
    

    ElseIf Node = "Daily Cash Book" Then
        VRun = "Daily Cash Book"
        
        PicDateBw.Enabled = True
        PicCode.Enabled = False
        PicDate.Enabled = False
        PicLookup.Enabled = False
        PicLookup.BackColor = vbWhite
        PicDate.BackColor = vbWhite
        PicDateBw.BackColor = &HC0C0C0
        PicCode.BackColor = vbWhite
        TxtFrom.SetFocus
        PicLookup1.Enabled = False
        PicLookup1.BackColor = vbWhite
    
    
    ElseIf Node = "Date Wise Purchase History" Then
        VRun = "Date Wise Purchase History"
        
        PicDateBw.Enabled = True
        PicCode.Enabled = False
        PicDate.Enabled = False
        PicLookup.Enabled = False
        PicLookup.BackColor = vbWhite
        PicDate.BackColor = vbWhite
        PicDateBw.BackColor = &HC0C0C0
        PicCode.BackColor = vbWhite
        TxtFrom.SetFocus
        PicLookup1.Enabled = False
        PicLookup1.BackColor = vbWhite
    
    ElseIf Node = "Supplier Wise Purchase History" Then
        VRun = "Supplier Wise Purchase History"
        
        PicDateBw.Enabled = True
        PicCode.Enabled = False
        PicDate.Enabled = False
        PicLookup.Enabled = True
        PicLookup.BackColor = &HC0C0C0
        PicDate.BackColor = vbWhite
        PicDateBw.BackColor = &HC0C0C0
        PicCode.BackColor = vbWhite
        TxtFrom.SetFocus
        PicLookup1.Enabled = False
        PicLookup1.BackColor = vbWhite
    
    ElseIf Node = "Product Wise Purchase History" Then
        VRun = "Product Wise Purchase History"
        
        PicDateBw.Enabled = True
        PicCode.Enabled = False
        PicDate.Enabled = False
        PicLookup.Enabled = True
        PicLookup.BackColor = &HC0C0C0
        PicDate.BackColor = vbWhite
        PicDateBw.BackColor = &HC0C0C0
        PicCode.BackColor = vbWhite
        TxtFrom.SetFocus
        PicLookup1.Enabled = False
        PicLookup1.BackColor = vbWhite
    
    ElseIf Node = "Product Wise Sale History" Then
        VRun = "Product Wise Sale History"
        
        PicDateBw.Enabled = True
        PicCode.Enabled = False
        PicDate.Enabled = False
        PicLookup.Enabled = True
        PicLookup.BackColor = &HC0C0C0
        PicDate.BackColor = vbWhite
        PicDateBw.BackColor = &HC0C0C0
        PicCode.BackColor = vbWhite
        TxtFrom.SetFocus
        PicLookup1.Enabled = False
        PicLookup1.BackColor = vbWhite
    
    
    
    ElseIf Node = "Sales Invoice" Then
        VRun = "Sales Invoice"
        
        PicDateBw.Enabled = False
        PicCode.Enabled = True
        PicDate.Enabled = False
        PicLookup.Enabled = False
        PicLookup.BackColor = vbWhite
        PicDate.BackColor = vbWhite
        PicDateBw.BackColor = vbWhite
        PicCode.BackColor = &HC0C0C0
        PicLookup1.Enabled = False
        PicLookup1.BackColor = vbWhite
        TxtCode.SetFocus
        
    ElseIf Node = "Date Wise Sale History" Then
        VRun = "Date Wise Sale History"
        
        PicDateBw.Enabled = True
        PicCode.Enabled = False
        PicDate.Enabled = False
        PicLookup.Enabled = False
        PicLookup.BackColor = vbWhite
        PicDate.BackColor = vbWhite
        PicDateBw.BackColor = &HC0C0C0
        PicCode.BackColor = vbWhite
        TxtFrom.SetFocus
        PicLookup1.Enabled = False
        PicLookup1.BackColor = vbWhite
    
    ElseIf Node = "Customer Wise Sale History" Then
        VRun = "Customer Wise Sale History"
        
        PicDateBw.Enabled = True
        PicCode.Enabled = False
        PicDate.Enabled = False
        PicLookup.Enabled = True
        PicLookup.BackColor = &HC0C0C0
        PicDate.BackColor = vbWhite
        PicDateBw.BackColor = &HC0C0C0
        PicCode.BackColor = vbWhite
        TxtFrom.SetFocus
        PicLookup1.Enabled = False
        PicLookup1.BackColor = vbWhite
    
    
    ElseIf Node = "Vehicle Wise Sale History" Then
        VRun = "Vehicle Wise Sale History"
        
        PicDateBw.Enabled = True
        PicCode.Enabled = False
        PicDate.Enabled = False
        PicLookup.Enabled = True
        PicLookup.BackColor = &HC0C0C0
        PicDate.BackColor = vbWhite
        PicDateBw.BackColor = &HC0C0C0
        PicCode.BackColor = vbWhite
        TxtFrom.SetFocus
        PicLookup1.Enabled = False
        PicLookup1.BackColor = vbWhite
    
    
    ElseIf Node = "Product Price List" Then
        VRun = "Product Price List"
        
        PicDateBw.Enabled = False
        PicDateBw.BackColor = vbWhite
        PicCode.Enabled = False
        PicDate.Enabled = False
        PicLookup.Enabled = False
        PicLookup.BackColor = vbWhite
        PicLookup1.BackColor = vbWhite
        PicLookup1.Enabled = False
        PicDate.BackColor = vbWhite
    
    ElseIf Node = "Stock Adjustment" Then
        VRun = "Stock Adjustment"
        
        PicDateBw.Enabled = True
        PicDateBw.BackColor = &HC0C0C0
        PicCode.Enabled = False
        PicDate.Enabled = False
        PicLookup.Enabled = True
        PicLookup.BackColor = &HC0C0C0
        PicLookup1.BackColor = vbWhite
        PicLookup1.Enabled = False
        PicDate.BackColor = vbWhite
        TxtFrom.SetFocus
    
'RECEIPTS
    ElseIf Node = "Date Wise Receipts" Then
        VRun = "Date Wise Receipts"
        
        PicDateBw.Enabled = True
        PicDateBw.BackColor = &HC0C0C0
        PicCode.Enabled = False
        PicDate.Enabled = False
        PicLookup.Enabled = False
        PicLookup.BackColor = vbWhite
        PicLookup1.BackColor = vbWhite
        PicLookup1.Enabled = False
        PicDate.BackColor = vbWhite
        TxtFrom.SetFocus
    
    ElseIf Node = "Account Wise Receipts" Then
        VRun = "Account Wise Receipts"
        
        PicDateBw.Enabled = True
        PicDateBw.BackColor = &HC0C0C0
        PicCode.Enabled = False
        PicDate.Enabled = False
        PicLookup.Enabled = True
        PicLookup.BackColor = &HC0C0C0
        PicLookup1.BackColor = vbWhite
        PicLookup1.Enabled = False
        PicDate.BackColor = vbWhite
        TxtFrom.SetFocus
    
'PAYMENTS
    ElseIf Node = "Date Wise Payments" Then
        VRun = "Date Wise Payments"
        
        PicDateBw.Enabled = True
        PicDateBw.BackColor = &HC0C0C0
        PicCode.Enabled = False
        PicDate.Enabled = False
        PicLookup.Enabled = False
        PicLookup.BackColor = vbWhite
        PicLookup1.BackColor = vbWhite
        PicLookup1.Enabled = False
        PicDate.BackColor = vbWhite
        TxtFrom.SetFocus
    
    ElseIf Node = "Account Wise Payments" Then
        VRun = "Account Wise Payments"
        
        PicDateBw.Enabled = True
        PicDateBw.BackColor = &HC0C0C0
        PicCode.Enabled = False
        PicDate.Enabled = False
        PicLookup.Enabled = True
        PicLookup.BackColor = &HC0C0C0
        PicLookup1.BackColor = vbWhite
        PicLookup1.Enabled = False
        PicDate.BackColor = vbWhite
        TxtFrom.SetFocus
    
'VEHICLES
    ElseIf Node = "Department Vehicles" Then
        VRun = "Department Vehicles"
        
        PicDateBw.Enabled = False
        PicDateBw.BackColor = vbWhite
        PicCode.Enabled = False
        PicDate.Enabled = False
        PicLookup.Enabled = True
        PicLookup.BackColor = &HC0C0C0
        PicLookup1.BackColor = vbWhite
        PicLookup1.Enabled = False
        PicDate.BackColor = vbWhite
        TxtLookup.SetFocus

'METER
    ElseIf Node = "Issues" Then
        VRun = "Issues"
        
        PicDateBw.Enabled = True
        PicDateBw.BackColor = &HC0C0C0
        PicCode.Enabled = False
        PicDate.Enabled = False
        PicLookup.Enabled = True
        PicLookup.BackColor = &HC0C0C0
        PicLookup1.BackColor = vbWhite
        PicLookup1.Enabled = False
        PicDate.BackColor = vbWhite
        TxtFrom.SetFocus

    ElseIf Node = "Reading" Then
        VRun = "Reading"
        
        PicDateBw.Enabled = True
        PicDateBw.BackColor = &HC0C0C0
        PicCode.Enabled = False
        PicDate.Enabled = False
        PicLookup.Enabled = True
        PicLookup.BackColor = &HC0C0C0
        PicLookup1.BackColor = vbWhite
        PicLookup1.Enabled = False
        PicDate.BackColor = vbWhite
        TxtFrom.SetFocus

    ElseIf Node = "Meter Stock" Then
        VRun = "Meter Stock"
        
        PicDateBw.Enabled = False
        PicDateBw.BackColor = vbWhite
        PicCode.Enabled = False
        PicDate.Enabled = False
        PicLookup.Enabled = False
        PicLookup.BackColor = vbWhite
        PicLookup1.BackColor = vbWhite
        PicLookup1.Enabled = False
        PicDate.BackColor = vbWhite


    ElseIf Node = "Code Wise Remaining" Then
        VRun = "Code Wise Remaining"
        
        PicDateBw.Enabled = False
        PicDateBw.BackColor = vbWhite
        PicCode.Enabled = False
        PicDate.Enabled = False
        PicLookup.Enabled = True
        PicLookup.BackColor = &HC0C0C0
        PicLookup1.BackColor = vbWhite
        PicLookup1.Enabled = False
        PicDate.BackColor = vbWhite
        TxtLookup.SetFocus

    
    
    
    
    End If
    TxtLookup = ""
    TxtLookupCode = ""
    TxtLookup1 = ""
    TxtLookupCode1 = ""
    TxtCode = ""
    TxtToCode = ""
End Sub

Private Sub TxtCode_GotFocus()
'     SelectAll TxtLookup
End Sub

Private Sub TxtCode_KeyPress(KeyAscii As Integer)
    Cng KeyAscii
End Sub

Private Sub TxtFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        TxtTo.SetFocus
    End If
    
End Sub

Private Sub TxtGrdSrch_Change()
    Call SearchRecord
End Sub

Private Sub TxtLookup_Change()
    If TxtLookup.Text = "" Then
        TxtLookupCode.Text = ""
    End If
End Sub

Private Sub TxtLookup_GotFocus()
''    SelectAll TxtLookup
End Sub
Private Sub TxtLookup_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 32 And Shift = 2 Then
        Select Case VRun
            Case "Accounts Status"
                ListLov.Clear
                ListLov.Visible = True
                ListLov.SetFocus
                Call AccountType
                
            Case "Accounts Ledger"
                PicSrchGrid.Visible = True
                TxtGrdSrch.Text = ""
                ListLov.Clear
                Call AccountsData
                
            Case "Account Wise Receipts"
                PicSrchGrid.Visible = True
                TxtGrdSrch.Text = ""
                ListLov.Clear
                Call AccountsData
            
            Case "Account Wise Payments"
                PicSrchGrid.Visible = True
                TxtGrdSrch.Text = ""
                ListLov.Clear
                Call AccountsData
                
            Case "Department Vehicles"
                PicSrchGrid.Visible = True
                TxtGrdSrch.Text = ""
                ListLov.Clear
                Call AccountsData
            
                
            Case "Receivable"
                ListLov.Clear
                ListLov.Visible = True
                Call AccountType
                
            Case "Payable"
                ListLov.Clear
                ListLov.Visible = True
                Call AccountType
                
            Case "Income Statement"
                ListLov.Visible = True
                Call MonthNames
                ListLov.SetFocus
                
            Case "Supplier Wise Purchase History"
                PicSrchGrid.Visible = True
                TxtGrdSrch.Text = ""
                ListLov.Clear
                Call AccountsData
        
            Case "Product Wise Purchase History"
                PicSrchGrid.Visible = True
                TxtGrdSrch.Text = ""
                ListLov.Clear
                Call AccountsData
        
            Case "Product Wise Sale History"
                PicSrchGrid.Visible = True
                TxtGrdSrch.Text = ""
                ListLov.Clear
                Call AccountsData
            
            Case "Customer Wise Sale History"
                PicSrchGrid.Visible = True
                TxtGrdSrch.Text = ""
                ListLov.Clear
                Call AccountsData
        
                
        End Select
    End If
End Sub

''Public Sub StockStatus()
''
''
''        Connect.Cn.Execute "delete from RepStockStat"
''        DataEnvironment2.ProductsStatus Val(TxtLookupCode)
''        For a = 1 To DataEnvironment2.rsProductsStatus.RecordCount
''            If RsDTrans.State = 1 Then RsDTrans.Close
''            RsDTrans.Open "Select sum(purchasedetail.qtys)from purchasedetail inner join purchaseheader on purchasedetail.purcode=purchaseheader.purcode where productCode= " & Val(DataEnvironment2.rsProductsStatus("ProductCode")) & " and purchaseheader.purdate< '" & Format(TxtFrom.Value, "DD-MMM-YY") & " 00:00:00.000'", Connect.Cn, adOpenStatic, adLockReadOnly
''            If IsNull(RsDTrans(0)) = False Then
''                DrSum = RsDTrans(0)
''            Else
''                DrSum = 0
''            End If
''            RsDTrans.Close
''            If RsDTrans.State = 1 Then RsDTrans.Close
''            RsDTrans.Open "Select sum(salesdetail.qty)from salesdetail inner join salesheader on salesdetail.salcode=salesheader.salcode where productCode= " & Val(DataEnvironment2.rsProductsStatus("ProductCode")) & " and salesheader.saldate< '" & Format(TxtFrom.Value, "DD-MMM-YY") & " 00:00:00.000'", Connect.Cn, adOpenStatic, adLockReadOnly
''            If IsNull(RsDTrans(0)) = False Then
''                CrSum = RsDTrans(0)
''            Else
''                CrSum = 0
''            End If
''            RsDTrans.Close
''            OpBal = DrSum - CrSum
''            RsDTrans.Open "Select sum(purchasedetail.qtys),sum(purchasedetail.bonus) from purchasedetail inner join purchaseheader on purchasedetail.purcode=purchaseheader.purcode where purchasedetail.productCode= " & Val(DataEnvironment2.rsProductsStatus("ProductCode")) & " and purchaseheader.purdate between '" & Format(TxtFrom.Value, "DD-MMM-YY") & " 00:00:00.000'" & " And '" & Format(TxtTo.Value, "DD-MMM-YY") & " 00:00:00.000'", Connect.Cn, adOpenStatic, adLockReadOnly
''            If IsNull(RsDTrans(0)) = False Then
''                Rec = RsDTrans(0)
''                TBonus = RsDTrans(1)
''            Else
''                Rec = 0
''                TBonus = 0
''            End If
''            RsDTrans.Close
''            RsDTrans.Open "Select sum(salesdetail.qty)from salesdetail inner join salesheader on salesdetail.salcode=salesheader.salcode where salesdetail.productCode= " & Val(DataEnvironment2.rsProductsStatus("ProductCode")) & " and salesheader.saldate between '" & Format(TxtFrom.Value, "DD-MMM-YY") & " 00:00:00.000'" & " And '" & Format(TxtTo.Value, "DD-MMM-YY") & " 00:00:00.000'", Connect.Cn, adOpenStatic, adLockReadOnly
''            If IsNull(RsDTrans(0)) = False Then
''                TSale = RsDTrans(0)
''            Else
''                TSale = 0
''            End If
''            RsDTrans.Close
''            If RsDTrans.State = 1 Then RsDTrans.Close
''            RsDTrans.Open "Select sum(salesdetail.qty)from salesdetail inner join salesheader on salesdetail.salcode=salesheader.salcode where productCode= " & Val(DataEnvironment2.rsProductsStatus("ProductCode")) & " and salesheader.saldate= '" & Format(TxtTo.Value, "DD-MMM-YY") & " 00:00:00.000'", Connect.Cn, adOpenStatic, adLockReadOnly
''            If IsNull(RsDTrans(0)) = False Then
''                DSale = RsDTrans(0)
''            Else
''                DSale = 0
''            End If
''            RsDTrans.Close
''            '**************** insert into RepStockStat
''            VSQL = "INSERT INTO RepStockStat VALUES ("
''            VSQL = VSQL & a & ",'" & DataEnvironment2.rsProductsStatus("ProductName")
''            VSQL = VSQL & "','" & DataEnvironment2.rsProductsStatus("Packing")
''            VSQL = VSQL & "'," & DataEnvironment2.rsProductsStatus("PrincipalCode")
''            VSQL = VSQL & "," & DataEnvironment2.rsProductsStatus("Rate")
''            VSQL = VSQL & "," & OpBal
''            VSQL = VSQL & "," & Rec
''            VSQL = VSQL & "," & OpBal + Rec
''            VSQL = VSQL & "," & DSale
''            VSQL = VSQL & "," & TSale
''            VSQL = VSQL & "," & TSale * DataEnvironment2.rsProductsStatus("Rate")
''            VSQL = VSQL & "," & DataEnvironment2.rsProductsStatus("Claim")
''            VSQL = VSQL & "," & DataEnvironment2.rsProductsStatus("Bonus")
''            VSQL = VSQL & "," & DataEnvironment2.rsProductsStatus("CurBal")
''            VSQL = VSQL & "," & DataEnvironment2.rsProductsStatus("TotalAmount")
''            VSQL = VSQL & ",'" & TxtTo.Value
''            VSQL = VSQL & "'," & TBonus & ")"
''            Cn.Execute VSQL
'''            MsgBox DataEnvironment2.rsProductsStatus(1) & ",  " & OpBal
''            DataEnvironment2.rsProductsStatus.MoveNext
''        Next
''        DataEnvironment2.rsProductsStatus.Close
''End Sub
Private Sub TxtLookup_KeyPress(KeyAscii As Integer)
''    MMisFuncion.SpaceNotAllow KeyAscii, TxtLookup
End Sub

Private Sub TxtLookup_LostFocus()
    If TxtLookup.Text = "ALL" Then
        TxtLookup1.Text = ""
    End If
End Sub

Private Sub TxtLookup1_Change()
    If TxtLookup1.Text = "" Then
        TxtLookupCode1.Text = ""
    End If
    
End Sub

Private Sub TxtLookup1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 32 And Shift = 2 Then
        Select Case VRun
            
            Case "Accounts Status"
                PicSrchGrid.Visible = True
                Call AccountsData
                MshSearch.SetFocus
        
        End Select
    End If
End Sub

Private Sub TxtLookup1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub TxtTo_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        TxtLookup.SetFocus
        On Error Resume Next
    
    End If
End Sub

Private Sub TxttoCode_GotFocus()
'     SelectAll TxtLookup
End Sub
Public Sub BalanceRep() 'FAROOQ
''''''Deleting old records
'''''
'''''    Set RsLov = New ADODB.Recordset
'''''    If RsLov.State = 1 Then RsLov.Close
'''''        RsLov.Open "Delete * from TblRep", Db, adOpenStatic, adLockOptimistic
'''''
'''''    Db.Execute "Delete from TblRep"
'''''
'''''    RepBalances.Refresh
'''''
''''''Opening Accounts Table
'''''    Set RsAcc = New ADODB.Recordset
'''''    If RsAcc.State = 1 Then RsAcc.Close
'''''        RsAcc.Open StrQry, Db, adOpenStatic, adLockPessimistic
'''''        Dim Vc1 As Single
'''''        For Vc1 = 1 To RsAcc.RecordCount
'''''
''''''Opening record set for Debit Transactions
'''''            Set RsDrt = New ADODB.Recordset
'''''                If RsDrt.State = 1 Then RsDrt.Close
'''''                    RsDrt.Open "Select Sum(Amount) from DRTGLT where AccountCode = " & RsAcc.Fields("Code"), Db, adOpenStatic, adLockPessimistic
'''''                        If IsNull(RsDrt(0)) Then
'''''                            DrBal = 0
'''''                        Else
'''''                            DrBal = Val(RsDrt(0))
'''''                        End If
'''''
''''''Opening record set for Credit Transactions
'''''            Set RsCrt = New ADODB.Recordset
'''''                If RsCrt.State = 1 Then RsCrt.Close
'''''                    RsCrt.Open "Select Sum(Amount) from CRTGLT where AccountCode = " & RsAcc.Fields("Code"), Db, adOpenStatic, adLockPessimistic
'''''                        If IsNull(RsCrt(0)) Then
'''''                            CrBal = 0
'''''                        Else
'''''                            CrBal = Val(RsCrt(0))
'''''                        End If
'''''
'''''
''''''Generating Difference between Total Debit and Total Credit
'''''    Dim VBal As Double
'''''    VBal = Val(DrBal) - Val(CrBal)
'''''
''''''Inserting data into TblRep
'''''    Db.Execute "INSERT INTO TblRep (Code, Descrip, TotalDr, TotalCr, Balance, DrCr) VALUES (" & Val(RsAcc("Code")) & ", '" & RsAcc("Name") & "', " & Val(DrBal) & ", " & Val(CrBal) & ", " & Val(VBal) & ", '" & "-" & "' )"
'''''
'''''''''Updating Record to know that is the balacne DR or CR
''''''''        If Val(VBal) > 0 Then
''''''''            Db.Execute "UPDATE TblRep SET DrCr = '" & "Dr" & "' WHERE TotalDr > 0"
''''''''        ElseIf Val(VBal) < 0 Then
''''''''            Db.Execute "UPDATE TblRep SET DrCr = '" & "Cr" & "' WHERE TotalCr > 0"
''''''''        Else
''''''''            Db.Execute "UPDATE TblRep SET DrCr = '" & "--" & "' WHERE BALANCE = 0"
''''''''        End If
'''''
'''''            DrBal = 0
'''''            CrBal = 0
'''''            VBal = 0
'''''
'''''
'''''            RsAcc.MoveNext
'''''        Next
'''''
'''''    On Error Resume Next
'''''
'''''        If DeRep.rsBalances.State = 1 Then DeRep.rsBalances.Close
'''''
'''''
'''''        DeRep.Balances
'''''
'''''''        DeRep.rsBalances.Requery
'''''
'''''    On Error Resume Next

End Sub
Public Sub AccountType()
    Set RsLov = New ADODB.Recordset
    If RsLov.State = 1 Then RsLov.Close
        RsLov.Open "SELECT DISTINCT AcType FROM Accounts Where AcType not in ('Product')", Con, adOpenStatic, adLockReadOnly
            For a = 1 To RsLov.RecordCount
                ListLov.AddItem RsLov(0)
                RsLov.MoveNext
            Next
            ListLov.AddItem "ALL"
End Sub


''Public Sub TBal()
'''Deleting old records
''    Cn.Execute "Delete from TblTrial"
''
'''Opening Accounts Table
''    Set RsAcc = New ADODB.Recordset
''    If RsAcc.State = 1 Then RsAcc.Close
''        RsAcc.Open "Select Code,Remarks from Accounts", Cn, adOpenStatic, adLockPessimistic
''
''        Dim Vc1 As Single
''        For Vc1 = 1 To RsAcc.RecordCount
''
'''Opening record set for Debit Transactions
''            Set RsDrt = New ADODB.Recordset
''                If RsDrt.State = 1 Then RsDrt.Close
''                    RsDrt.Open "Select Sum(Amount) from DRt where AccountCode = " & RsAcc.Fields("Code"), Cn, adOpenStatic, adLockPessimistic
''                        If IsNull(RsDrt(0)) Then
''                            DrBal = 0
''                        Else
''                            DrBal = RsDrt(0)
''                        End If
''
'''Opening record set for Credit Transactions
''            Set RsCrt = New ADODB.Recordset
''                If RsCrt.State = 1 Then RsCrt.Close
''                    RsCrt.Open "Select Sum(Amount) from CRt where AccountCode = " & RsAcc.Fields("Code"), Cn, adOpenStatic, adLockPessimistic
''                        If IsNull(RsCrt(0)) Then
''                            CrBal = 0
''                        Else
''                            CrBal = RsCrt(0)
''                        End If
''
''
'''Generating Difference between Total Debit and Total Credit for each account
''    Dim VBal As Double
''    VBal = Val(DrBal) - Val(CrBal)
''
''        Cn.Execute "INSERT INTO TblTrial (Code, Descrip, Debit, Credit,Id) VALUES (" & Val(RsAcc("Code")) & ", '" & RsAcc("Remarks") & "',0,0, " & Val(Vc1) & ")"
''
'''Updating Record to know that is the balacne DR or CR
''        If Val(VBal) > 0 Then
''            Cn.Execute "UPDATE TblTrial SET Debit = " & Val(VBal) & " where Id = " & Val(Vc1) & ""
''        ElseIf Val(VBal) < 0 Then
''            Cn.Execute "UPDATE TblTrial SET Credit = " & Abs(Val(VBal)) & " where Id = " & Val(Vc1) & ""
''        End If
''
''            VBal = 0
''            RsAcc.MoveNext
''        Next
''
'''Getting the difference between Total Debit and Total Credit
''    Set RsLov = New ADODB.Recordset
''    If RsLov.State = 1 Then RsLov.Close
''        RsLov.Open "Select sum(Debit),sum(Credit) from TblTrial", Cn, adOpenStatic, adLockReadOnly
''            VBal = Val(RsLov(0)) - Val(RsLov(1))
''            If Val(VBal) > 0 Then
''                RepTrial.Sections("Section5").Controls("LblDif").Caption = "Dr. Diff: " & Format(Val(VBal), "##,#.00")
''            ElseIf Val(VBal) < 0 Then
''                RepTrial.Sections("Section5").Controls("LblDif").Caption = "Cr. Diff: " & Format(Val(VBal), "##,#.00")
''            Else
''                RepTrial.Sections("Section5").Controls("LblDif").Caption = Format(0, "##,#.00")
''            End If
''
'''Showing Report
''    If De1.rsRepTrial.State = 1 Then De1.rsRepTrial.Close
''    RepTrial.Refresh
''    Load RepTrial
''    RepTrial.Show 1
''End Sub

Public Sub AccountsData()
    
        Select Case VRun
            
            Case "Accounts Status"
                SetSearchGrid
                FillGridAccounts
                MshSearch.Col = 0
                MshSearch.Row = 1
                MshSearch.SetFocus
        
            Case "Accounts Ledger"
                SetSearchGrid
                FillGridAccounts
                MshSearch.Col = 0
                MshSearch.Row = 1
                MshSearch.SetFocus

            Case "Account Wise Receipts"
                SetSearchGrid
                FillGridAccounts
                MshSearch.Col = 0
                MshSearch.Row = 1
                MshSearch.SetFocus

            Case "Account Wise Payments"
                SetSearchGrid
                FillGridAccounts
                MshSearch.Col = 0
                MshSearch.Row = 1
                MshSearch.SetFocus

            Case "Supplier Wise Purchase History"
                SetSearchGrid
                FillGridAccounts
                MshSearch.Col = 0
                MshSearch.Row = 1
                MshSearch.SetFocus

            Case "Product Wise Purchase History"
                SetSearchGrid
                FillGridAccounts
                MshSearch.Col = 0
                MshSearch.Row = 1
                MshSearch.SetFocus

            Case "Customer Wise Sale History"
                SetSearchGrid
                FillGridAccounts
                MshSearch.Col = 0
                MshSearch.Row = 1
                MshSearch.SetFocus
        
            Case "Product Wise Sale History"
                SetSearchGrid
                FillGridAccounts
                MshSearch.Col = 0
                MshSearch.Row = 1
                MshSearch.SetFocus
        
            
            Case "Month Wise Account Status"
                Set RsLov = New ADODB.Recordset
                If RsLov.State = 1 Then RsLov.Close
                    RsLov.Open "SELECT Code,Name FROM Accounts ORDER BY Remarks", Con, adOpenStatic, adLockReadOnly
                        Set MshSearch.DataSource = RsLov
                        MshSearch.ColWidth(0) = 1000
                        MshSearch.ColWidth(1) = 4000
                        MshSearch.SetFocus
        
            Case "Receivable"
                Set RsLov = New ADODB.Recordset
                If RsLov.State = 1 Then RsLov.Close
                    RsLov.Open "Select Code,Name from Accounts Where AccType = '" & ListLov & "' order By Remarks", Con, adOpenStatic, adLockReadOnly
                        Set MshSearch.DataSource = RsLov
                        PicSrchGrid.Visible = True
                        MshSearch.ColWidth(0) = 1000
                        MshSearch.ColWidth(1) = 4000
                        MshSearch.SetFocus
        
            Case "Payable"
                Set RsLov = New ADODB.Recordset
                If RsLov.State = 1 Then RsLov.Close
                    RsLov.Open "Select Code,Name from Accounts Where AccType = '" & ListLov & "' order By Remarks", Con, adOpenStatic, adLockReadOnly
                        Set MshSearch.DataSource = RsLov
                        PicSrchGrid.Visible = True
                        MshSearch.ColWidth(0) = 1000
                        MshSearch.ColWidth(1) = 4000
                        MshSearch.SetFocus
        
        
        End Select
End Sub

Public Sub AccLedger()
Dim DtFrom As Date
Dim DtTo As Date
Dim mBalance As Single
Dim DrOpBal As Single
Dim CrOpBal As Single
Dim DrCr As String
Dim RunningBal As Double

    DtFrom = TxtFrom.Value
    DtTo = TxtTo.Value

    mDateQry = " BETWEEN #" & DtFrom & "# And #" & DtTo & "#"

'Delete Data from RepAcLedger
    Con.Execute "Delete * from RepAcLedger"

'Getting Opening Balance for the selected Account
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
        RS.Open "Select Sum(DrAmount),Sum(CrAmount) from ViewAccountLedger where TransDate < #" & DtFrom & "# and AcID = " & Val(TxtLookupCode) & "", Con, adOpenStatic, adLockOptimistic
'If result is NULL
            If IsNull(RS(0)) Or IsNull(RS(1)) Then
                mBalance = 0
            Else
                mBalance = Val(RS(0)) - Val(RS(1))
            End If
            
'Finding DR/CR/Nill balance
            If Val(mBalance) > 0 Then
                DrOpBal = Val(mBalance)
                CrOpBal = 0
                DrCr = "DR"
            ElseIf Val(mBalance) < 0 Then
                CrOpBal = Abs(Val(mBalance))
                DrOpBal = 0
                DrCr = "CR"
            Else
                DrOpBal = 0
                CrOpBal = 0
                DrCr = "NILL"
            End If

'Saving Opening Balance in REPACLEDGER
    Con.Execute "Insert into RepAcLedger Values( " & Val(TxtLookupCode) & ", '" & Format(DtFrom - 1, "dd/mm/yyyy") & "', '" & "Last Balance (C/F)" & "', " & Val(DrOpBal) & ", " & Val(CrOpBal) & ", " & Abs(Val(mBalance)) & ", '" & DrCr & "') "
    
'Fetching Data into RepAcLedger
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
        RS.Open "Select * from ViewAccountLedger where AcId = " & Val(TxtLookupCode) & " and  TransDate " & mDateQry, Con, adOpenStatic, adLockOptimistic
        RS.Requery
        
            While Not RS.EOF = True
                RunningBal = Val(mBalance) + Val(RS(4)) - Val(RS(5))
                mBalance = RunningBal
                
    'Finding DR/CR/Nill balance
            If Val(RunningBal) > 0 Then
                DrCr = "DR"
            ElseIf Val(RunningBal) < 0 Then
                DrCr = "CR"
            Else
                DrCr = "NILL"
            End If

'                                                                AcID                         Date                         Desc                DR                  CR              balance                     drcr
                Con.Execute "Insert into RepAcLedger Values( " & Val(RS(0)) & ", '" & Format(RS(2), "dd/mm/yyyy") & "', '" & RS(3) & "', " & Val(RS(4)) & ", " & Val(RS(5)) & ", " & Val(RunningBal) & ", '" & DrCr & "') "
                RS.MoveNext
            Wend
            
    daraRepAcLedger.Show vbModal

End Sub

Public Sub MonthNames()
'List of Months for LOV
        ListLov.Clear
        
        ListLov.AddItem "January"
        ListLov.AddItem "February"
        ListLov.AddItem "March"
        ListLov.AddItem "April"
        ListLov.AddItem "May"
        ListLov.AddItem "June"
        ListLov.AddItem "July"
        ListLov.AddItem "August"
        ListLov.AddItem "September"
        ListLov.AddItem "October"
        ListLov.AddItem "November"
        ListLov.AddItem "December"
End Sub

Public Sub MonthDates()
'Days of Monthts for Calcutaing COG
    If TxtLookup.Text = "January" Then
        DtFrom = "01-Jan-" & TxtLookup1.Text
        DtTo = "31-Jan-" & TxtLookup1.Text
    
    ElseIf TxtLookup.Text = "February" Then
        DtFrom = "01-Feb-" & TxtLookup1.Text
        DtTo = "28-Feb-" & TxtLookup1.Text
    
    ElseIf TxtLookup.Text = "March" Then
        DtFrom = "01-Mar-" & TxtLookup1.Text
        DtTo = "31-Mar-" & TxtLookup1.Text
    
    ElseIf TxtLookup.Text = "April" Then
        DtFrom = "01-Apr-" & TxtLookup1.Text
        DtTo = "30-Apr-" & TxtLookup1.Text
    
    ElseIf TxtLookup.Text = "May" Then
        DtFrom = "01-May-" & TxtLookup1.Text
        DtTo = "31-May-" & TxtLookup1.Text
    
    ElseIf TxtLookup.Text = "June" Then
        DtFrom = "01-Jun-" & TxtLookup1.Text
        DtTo = "30-Jun-" & TxtLookup1.Text

    ElseIf TxtLookup.Text = "July" Then
        DtFrom = "01-Jul-" & TxtLookup1.Text
        DtTo = "31-Jul-" & TxtLookup1.Text

    ElseIf TxtLookup.Text = "August" Then
        DtFrom = "01-Aug-" & TxtLookup1.Text
        DtTo = "31-Aug-" & TxtLookup1.Text

    ElseIf TxtLookup.Text = "September" Then
        DtFrom = "01-Sep-" & TxtLookup1.Text
        DtTo = "30-Sep-" & TxtLookup1.Text

    ElseIf TxtLookup.Text = "October" Then
        DtFrom = "01-Oct-" & TxtLookup1.Text
        DtTo = "31-Oct-" & TxtLookup1.Text

    ElseIf TxtLookup.Text = "November" Then
        DtFrom = "01-Nov-" & TxtLookup1.Text
        DtTo = "30-Nov-" & TxtLookup1.Text

    ElseIf TxtLookup.Text = "December" Then
        DtFrom = "01-Dec-" & TxtLookup1.Text
        DtTo = "31-Dec-" & TxtLookup1.Text

    End If
    
End Sub


''Public Sub IncomeStatement()
''    TxtFrom.Value = DtFrom
''    TxtTo.Value = DtTo
''
''    MDateQry = " BETWEEN '" & TxtFrom.Value & "' And '" & TxtTo.Value & "'"
''
''
'''Sales during the given period
''    Dim mTotalSales As Double
''
''    Set RsLov = New ADODB.Recordset
''    If RsLov.State = 1 Then RsLov.Close
''        RsLov.Open "SELECT SUM(SaleDetail.TotalAmount) FROM Sale INNER JOIN SaleDetail ON Sale.Code = SaleDetail.Code INNER JOIN Product ON SaleDetail.ProductCode = Product.Code WHERE Product.Type IN ('Billets/Ingots') and Sale.TDate" & MDateQry, Cn, adOpenStatic, adLockReadOnly
''            If IsNull(RsLov(0)) Then
''                mTotalSales = 0
''            Else
''                mTotalSales = Val(RsLov(0))
''            End If
''
''        RepIncomeStatement.Sections("Section2").Controls("LblTotalSale").Caption = Val(mTotalSales)
''
'''Sales Return during the period
''    Dim mSaleRet As Double
''
''    Set RsLov = New ADODB.Recordset
''    If RsLov.State = 1 Then RsLov.Close
''        RsLov.Open "SELECT SUM(SaleDetailR.TotalAmount)FROM SaleR INNER JOIN SaleDetailR ON SaleR.Code = SaleDetailR.Code INNER JOIN Product ON SaleDetailR.ProductCode = Product.Code WHERE Product.Type IN ('Billets/Ingots') and SaleR.TDate" & MDateQry, Cn, adOpenStatic, adLockReadOnly
''            If IsNull(RsLov(0)) Then
''                mSaleRet = 0
''            Else
''                mSaleRet = Val(RsLov(0))
''            End If
''
''        RepIncomeStatement.Sections("Section2").Controls("LblSaleRet").Caption = Val(mSaleRet)
''
'''Sales Discounts during the period
''    Dim mSaleDiscount As Double
''
''    Set RsLov = New ADODB.Recordset
''    If RsLov.State = 1 Then RsLov.Close
''        RsLov.Open "SELECT SUM(Drt.Amount) FROM GLT INNER JOIN DRT ON GLT.Code = DRT.Code INNER JOIN Accounts ON DRT.AccountCode = Accounts.Code WHERE (Accounts.AccType IN ('Discount Allowed')) and Glt.Tdate" & MDateQry, Cn, adOpenStatic, adLockReadOnly
''            If IsNull(RsLov(0)) Then
''                mSaleDiscount = 0
''            Else
''                mSaleDiscount = Val(RsLov(0))
''            End If
''
''        RepIncomeStatement.Sections("Section2").Controls("LblDiscount").Caption = Val(mSaleDiscount)
''
'''Calculating NetSales
''    Dim mNetSale As Double
''
''        mNetSale = Val(mTotalSales) - Val(mSaleRet) + Val(mSaleDiscount)
''        RepIncomeStatement.Sections("Section2").Controls("LblNetSale").Caption = Val(mNetSale)
''
'''Gross Profit
''    Dim mGrossProfit As Double
''        mGrossProfit = Val(mNetSale) - Val(mCOG)
''        RepIncomeStatement.Sections("Section2").Controls("LblGross").Caption = Val(mGrossProfit)
''
''
'''Operating Expense Report
''
''        StrQry = "SELECT Code, Remarks FROM Accounts WHERE AccType IN ('Opeating Expense') AND AccType NOT LIKE '%Tax%'"
''
'''========================Copying Data to TblRep===========================
''
'''Deleting old records
''    Cn.Execute "Delete from TblRep"
''    RepBalances.Refresh
''
'''Opening Accounts Table
''    Set RsAcc = New ADODB.Recordset
''    If RsAcc.State = 1 Then RsAcc.Close
''        RsAcc.Open StrQry, Cn, adOpenStatic, adLockPessimistic
''        Dim Vc1 As Single
''        For Vc1 = 1 To RsAcc.RecordCount
''
'''Opening record set for Debit Transactions
''            Set RsDrt = New ADODB.Recordset
''                If RsDrt.State = 1 Then RsDrt.Close
''                    RsDrt.Open "SELECT SUM(DRT.Amount) FROM GLT INNER JOIN DRT ON GLT.Code = DRT.Code where DRT.AccountCode = " & RsAcc.Fields("Code") & " and GLT.Tdate" & MDateQry, Cn, adOpenStatic, adLockPessimistic
''                        If IsNull(RsDrt(0)) Then
''                            DrBal = 0
''                        Else
''                            DrBal = Val(RsDrt(0))
''                        End If
''
'''Opening record set for Credit Transactions
''            Set RsCrt = New ADODB.Recordset
''                If RsCrt.State = 1 Then RsCrt.Close
''                    RsCrt.Open "SELECT SUM(CRT.Amount) FROM GLT INNER JOIN CRT ON GLT.Code = CRT.Code where CRT.AccountCode = " & RsAcc.Fields("Code") & " and GLT.TDate" & MDateQry, Cn, adOpenStatic, adLockPessimistic
''                        If IsNull(RsCrt(0)) Then
''                            CrBal = 0
''                        Else
''                            CrBal = Val(RsCrt(0))
''                        End If
''
''
'''Generating Difference between Total Debit and Total Credit
''    Dim VBal As Double
''    VBal = Val(DrBal) - Val(CrBal)
''
'''Inserting data into TblRep
''    Cn.Execute "INSERT INTO TblRep (Code, Descrip,TotalDr, TotalCr, Balance, DrCr) VALUES (" & Val(RsAcc("Code")) & ", '" & RsAcc("Remarks") & "', " & Val(DrBal) & ", " & Val(CrBal) & ", " & Abs(Val(VBal)) & ", '" & "-" & "' )"
''
'''Updating Record to know that is the balacne DR or CR
''        If Val(VBal) > 0 Then
''            Cn.Execute "UPDATE TblRep SET DrCr = '" & "Dr" & "' WHERE TotalDr > 0"
''        ElseIf Val(VBal) < 0 Then
''            Cn.Execute "UPDATE TblRep SET DrCr = '" & "Cr" & "' WHERE TotalCr > 0"
''        Else
''            Cn.Execute "UPDATE TblRep SET DrCr = '" & "--" & "' WHERE BALANCE = 0"
''        End If
''
''            DrBal = 0
''            CrBal = 0
''            VBal = 0
''
''
''            RsAcc.MoveNext
''        Next
''
''    Dim mTotalOperating As Double
''
''        Set RsLov = New ADODB.Recordset
''        If RsLov.State = 1 Then RsLov.Close
''            RsLov.Open "Select Sum(Balance) from TblRep", Cn, adOpenStatic, adLockReadOnly
''                If IsNull(RsLov(0)) Then
''                    mTotalOperating = 0
''                Else
''                    mTotalOperating = Val(RsLov(0))
''                End If
''
''    If De1.rsBalances.State = 1 Then De1.rsBalances.Requery
''
'''========================END Copying Data to TblRep===========================
''
''        RepIncomeStatement.Sections("Section5").Controls("LblTotalOperating").Caption = Val(mTotalOperating)
'''Net Income
''    Dim mNetIncome As Double
''        mNetIncome = Val(mGrossProfit) - Val(mTotalOperating)
''        RepIncomeStatement.Sections("Section5").Controls("LblNetIncome").Caption = Val(mNetIncome)
''
'''Other Income
''    Dim mOtherIncome As Double
''        Set RsLov = New ADODB.Recordset
''        If RsLov.State = 1 Then RsLov.Close
''            RsLov.Open "SELECT SUM(CRT.Amount) FROM GLT INNER JOIN CRT ON GLT.Code = CRT.Code INNER JOIN Accounts ON CRT.AccountCode = Accounts.Code WHERE Accounts.AccType = 'Other Income' and GLT.Tdate" & MDateQry, Cn, adOpenStatic, adLockReadOnly
''                If IsNull(RsLov(0)) Then
''                    mOtherIncome = 0
''                Else
''                    mOtherIncome = Val(RsLov(0))
''                End If
''
''        RepIncomeStatement.Sections("Section5").Controls("LblOtherIncome").Caption = Val(mOtherIncome)
'''Net Income before tax
''    Dim NetBeforeTax As Double
''        NetBeforeTax = Val(mNetIncome) + Val(mOtherIncome)
''        RepIncomeStatement.Sections("Section5").Controls("LblIncomeBeforeTax").Caption = Val(NetBeforeTax)
''
'''Finding Taxes Amount
''    Dim mTaxes As Double
''        Set RsLov = New ADODB.Recordset
''        If RsLov.State = 1 Then RsLov.Close
''            RsLov.Open "SELECT SUM(DRT.Amount) FROM GLT INNER JOIN DRT ON GLT.Code = DRT.Code INNER JOIN Accounts ON DRT.AccountCode = Accounts.Code WHERE Accounts.AccType like '%Tax%' and GLT.Tdate" & MDateQry, Cn, adOpenStatic, adLockReadOnly
''                If IsNull(RsLov(0)) Then
''                    mTaxes = 0
''                Else
''                    mTaxes = Val(RsLov(0))
''                End If
''
''        RepIncomeStatement.Sections("Section5").Controls("LblTaxes").Caption = Val(mTaxes)
''
'''Net Income After Taxes
''    Dim NetAfterTax As Double
''
''        NetAfterTax = Val(NetBeforeTax) - Val(mTaxes)
''        RepIncomeStatement.Sections("Section5").Controls("LblIncomeAfterTax").Caption = Val(NetAfterTax)
''
''End Sub

''Public Sub BalanceSheet()
'''Deleting Old Records
''    Cn.Execute "Delete from TblTrialCat"
''
'''Categorized Trial Balance
''    Set RsAcc = New ADODB.Recordset
''    If RsAcc.State = 1 Then RsAcc.Close
''        RsAcc.Open "Select MainHead,Code,Remarks from Accounts Where MainHead in (1,2,3)", Cn, adOpenStatic, adLockReadOnly
''            Dim Vc1 As Single
''            For Vc1 = 1 To RsAcc.RecordCount
''                Set RsAcType = New ADODB.Recordset
''                    If RsAcType.State = 1 Then RsAcType.Close
''                        RsAcType.Open "Select Code from AccountTypes where Code = " & Val(RsAcc("MainHead")) & "", Cn, adOpenStatic, adLockReadOnly
''
'''Opening record set for Debit Transactions
''                            Set RsDrt = New ADODB.Recordset
''                                If RsDrt.State = 1 Then RsDrt.Close
''                                    RsDrt.Open "Select Sum(Amount) from DRt where AccountCode = " & RsAcc("Code"), Cn, adOpenStatic, adLockPessimistic
''                                        If IsNull(RsDrt(0)) Then
''                                            DrBal = 0
''                                        Else
''                                            DrBal = RsDrt(0)
''                                        End If
''
'''Opening record set for Credit Transactions
''                            Set RsCrt = New ADODB.Recordset
''                                If RsCrt.State = 1 Then RsCrt.Close
''                                    RsCrt.Open "Select Sum(Amount) from CRt where AccountCode = " & RsAcc("Code"), Cn, adOpenStatic, adLockPessimistic
''                                        If IsNull(RsCrt(0)) Then
''                                            CrBal = 0
''                                        Else
''                                            CrBal = RsCrt(0)
''                                        End If
''
''
'''Generating Difference between Total Debit and Total Credit for each account
''    Dim VBal As Double
''    VBal = Val(DrBal) - Val(CrBal)
''
'''Inserting Values into TblTrialCat
''        Cn.Execute "INSERT INTO TblTrialCat (MainHead,Code, Descrip, Debit, Credit,Id) VALUES (" & Val(RsAcType("Code")) & ", " & Val(RsAcc("Code")) & ", '" & RsAcc("Remarks") & "',0,0, " & Val(Vc1) & ")"
''
'''Updating Record to know that is the balacne DR or CR
''        If Val(VBal) > 0 Then
''            Cn.Execute "UPDATE TblTrialCat SET Debit = " & Val(VBal) & " where Id = " & Val(Vc1) & ""
''        ElseIf Val(VBal) < 0 Then
''            Cn.Execute "UPDATE TblTrialCat SET Credit = " & Abs(Val(VBal)) & " where Id = " & Val(Vc1) & ""
''        End If
''
''            VBal = 0
''            RsAcc.MoveNext
''        Next
''
'''Showing Report
''    If De1.rsTrialCatMain.State = 1 Then De1.rsTrialCatMain.Close
''    RepTrialCat.Refresh
''    Load RepTrialCat
''    RepTrialCat.Show 1
''
''End Sub
Public Sub ExactRecords()
    Dim ExactRowCount As Integer
    
    If PicSrchGrid.Visible = True Then
    
        MshSearch.Rows = 2
        MshSearch.Row = 1
        
        ExactRowCount = 0
        
        If TxtGrdSrch.Text = "" Then
            SQLQry = "Select Code, Name, AccType from Accounts where Disable = 0"
        Else
            SQLQry = "Select Code, Name, AccType from Accounts Where Disable = 0 And Name Like '" & TxtGrdSrch.Text & "%'"
        End If
        
        Set RsSrch = New ADODB.Recordset
        If RsSrch.State = 1 Then RsSrch.Close
            RsSrch.Open SQLQry, dB, adOpenStatic, adLockReadOnly
                
                If RsSrch.RecordCount <= 0 Then
                    MshSearch.Rows = 1
                    Exit Sub
                End If
                
                For ExactRowCount = 1 To RsSrch.RecordCount
                    
                    If ExactRowCount >= MshSearch.Rows - 1 Then
                        MshSearch.Rows = MshSearch.Rows + 1
                        MshSearch.Row = MshSearch.Row + 1
                    End If
                    
                    MshSearch.TextMatrix(ExactRowCount, 0) = RsSrch(1)
                    MshSearch.TextMatrix(ExactRowCount, 1) = RsSrch(0)
                    MshSearch.TextMatrix(ExactRowCount, 2) = RsSrch(2)
                    
                    
                    RsSrch.MoveNext
                Next
                
                MshSearch.Col = 0
                MshSearch.Row = 1
                MshSearch.SetFocus
                
                MshSearch.ColAlignment(1) = 3
    End If
End Sub

Public Sub ProductStock()
    Dim mProdId     As Integer
    Dim mProdName   As String
    Dim mOpQty      As Single
    Dim mPurQty     As Single
    Dim mSalQty     As Single
    Dim mReOrderQty As Single
    Dim mCounter    As Integer
    Dim AvblStock   As Single
    
   
    Con.Execute "Delete from RepStockStatus"

'=====================Checking Available Stock=======================

        Set RsMisc = New ADODB.Recordset
        If RsMisc.State = 1 Then RsMisc.Close
            RsMisc.Open "Select AcId,AcTitle,ReOrderPoint from Accounts where AcType = 'Product'", Con, adOpenStatic, adLockOptimistic

                 
        For mCounter = 1 To RsMisc.RecordCount
    
                    mProdId = Val(RsMisc(0))
                    mProdName = RsMisc(1)
                    mReOrderQty = Val(RsMisc(2))
        'Opening Stock
                Set RS = New ADODB.Recordset
                If RS.State = 1 Then RS.Close
                    
                    RS.Open "Select * from Product_Openings where ProdId = " & Val(mProdId) & "", Con, adOpenStatic, adLockOptimistic
                        If RS.RecordCount <= 0 Then
                            GoTo CheckPurchase
                        Else
                            mOpQty = Val(RS(1))
                        End If
                    RS.Close
                    Set RS = Nothing
                    
CheckPurchase:
        'Purchases
                Set RS = New ADODB.Recordset
                If RS.State = 1 Then RS.Close
                    RS.Open "Select Sum(Qty) from Purchase_Detail where ProdId = " & Val(mProdId) & "", Con, adOpenStatic, adLockOptimistic
                        If IsNull(RS(0)) Then
                            mPurQty = 0
                            GoTo CheckSales
                        Else
                            mPurQty = Val(RS(0))
                        End If
                    RS.Close
                    Set RS = Nothing
                        
CheckSales:
        'Sales
                Set RS = New ADODB.Recordset
                If RS.State = 1 Then RS.Close
                    RS.Open "Select Sum(Qty) from Sale_Detail where ProdId = " & Val(mProdId) & "", Con, adOpenStatic, adLockOptimistic
                        If IsNull(RS(0)) Then
                            mSalQty = 0
                            GoTo MyStock
                        Else
                            mSalQty = Val(RS(0))
                        End If
                        
MyStock:
        'Available Stock
                AvblStock = Val(mOpQty) + Val(mPurQty) - Val(mSalQty)
                
        'Fetching Data into RepStockStatus
                Con.Execute "Insert Into RepStockStatus Values(" & Val(mProdId) & ", '" & mProdName & "', " & Val(AvblStock) & ", " & Val(mReOrderQty) & ")"
        
        RsMisc.MoveNext
        Next
        
        
    dataRepStockStatus.Show vbModal


End Sub


Public Sub PurchaseHistoryByDate()
    dataRepPurchaseByDate.Show vbModal
''SELECT Purchase_Main.PurId, Purchase_Main.PurDate, Accounts.AcId, Accounts.AcTitle, Purchase_Main.TotalAmount FROM (Purchase_Main INNER JOIN Purchase_Detail ON Purchase_Main.PurId = Purchase_Detail.PurId) INNER JOIN Accounts ON Purchase_Detail.AcId = Accounts.AcId;

End Sub

Public Sub PurchaseHistoryBySupplier()
    dataRepPurchaseBySupplier.Show vbModal
End Sub

Public Sub DateWiseSaleHistory()
    dataRepSaleByDate.Show vbModal
End Sub

Public Sub CustomerWiseSaleHistory()
    dataRepSaleByCustomer.Show vbModal
End Sub

Public Sub ProductPriceList()
    dataRepPriceList.Show vbModal
End Sub

Public Sub StockAdjustment()
    If DeRep.rsStockAdjustment.State = 1 Then DeRep.rsStockAdjustment.Close
        
        WaitMode
        
        DeRep.StockAdjustment TxtLookup.Text, TxtFrom.Value, TxtTo.Value
        
        RepStockAdjustment.Sections("Section4").Controls("LblDate").Caption = "From: " & TxtFrom.Value & "     " & "To: " & TxtTo.Value
        RepStockAdjustment.Sections("Section4").Controls("LblAdjType").Caption = UCase(TxtLookup.Text)
        
        Load RepStockAdjustment
        RepStockAdjustment.Refresh
        
        RepStockAdjustment.Show 1
End Sub
Public Sub VehData()
    Dim TotVeh As Integer
        
    Set RsLov = New ADODB.Recordset
    If RsLov.State = 1 Then RsLov.Close
        RsLov.Open "Select VehNo from VehicleReg", dB, adOpenStatic, adLockOptimistic
            If RsLov.RecordCount <= 0 Then
                Exit Sub
            End If
            
            For TotVeh = 1 To RsLov.RecordCount
                ListLov.AddItem RsLov(0)
                
                RsLov.MoveNext
            Next
            
            
End Sub

Public Sub VehicleWiseSaleHistory()

    If DeRep.rsInvoicePrint.State = 1 Then DeRep.rsInvoicePrint.Close
        DeRep.InvoicePrint TxtLookup.Text, TxtFrom.Value, TxtTo.Value
        
        RepInvoicePrint.Sections("ReportHeader").Controls("LblDate").Caption = "From: " & TxtFrom.Value & "     " & "To: " & TxtTo.Value
        RepInvoicePrint.Sections("ReportHeader").Controls("LblVeh").Caption = UCase(TxtLookup.Text)
        RepInvoicePrint.Sections("ReportHeader").Controls("LblMs").Caption = DeptName
        
        Load RepInvoicePrint
        RepInvoicePrint.Refresh
        
        RepInvoicePrint.Show 1
End Sub

Public Sub DeptData()
'Getting Dept Code for the selected Veh
    Set RsLov = New ADODB.Recordset
    If RsLov.State = 1 Then RsLov.Close
        RsLov.Open "Select DeptCode From VehicleReg Where VehNo = '" & TxtLookup.Text & "'", dB, adOpenStatic, adLockOptimistic
            
'Getting Dept Name for the Selected Veh
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
        RS.Open "Select Name from Customer Where Code = " & Val(RsLov(0)) & "", dB, adOpenStatic, adLockOptimistic
            DeptName = RS(0)
End Sub

Public Sub MeterStock()
    Dim Vc1 As Integer
    Dim mOpStock    As Single
    Dim mIssue      As Single
    Dim mAvbl       As Single
    Dim mConsume    As Single
    Dim mDiff       As Single
    
    
    dB.Execute "Delete from TblMeterStock"
    
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
        RS.Open "Select Code from MeterInfo", dB, adOpenStatic, adLockOptimistic
            
            If RS.EOF = True Then
                MsgBox "No data found", vbInformation, "Message"
                Exit Sub
            End If
            
    For Vc1 = 1 To RS.RecordCount
        
'Op Stock
        Set RsLov = New ADODB.Recordset
        If RsLov.State = 1 Then RsLov.Close
            RsLov.Open "Select OpStock from MeterInfo Where Code = " & Val(Vc1) & "", dB, adOpenForwardOnly, adLockReadOnly
                If IsNull(RsLov(0)) Then
                    mOpStock = 0
                Else
                    mOpStock = Val(RsLov(0))
                End If
                
'Issue
        Set RsLov = New ADODB.Recordset
        If RsLov.State = 1 Then RsLov.Close
            RsLov.Open "Select Sum(Qty) from MeterIssue Where MeterCode = " & Val(Vc1) & "", dB, adOpenForwardOnly, adLockReadOnly
                If IsNull(RsLov(0)) Then
                    mIssue = 0
                Else
                    mIssue = Val(RsLov(0))
                End If
                

        mAvbl = Val(mOpStock) + Val(mIssue)

'Consumption
        Set RsLov = New ADODB.Recordset
        If RsLov.State = 1 Then RsLov.Close
            RsLov.Open "Select Sum(ClosingReading)-Sum(OpReading)+Sum(TestQty) from MeterReading Where MeterCode = " & Val(Vc1) & "", dB, adOpenForwardOnly, adLockReadOnly
                If IsNull(RsLov(0)) Then
                    mConsume = 0
                Else
                    mConsume = Val(RsLov(0))
                End If
                
        mDiff = Val(mAvbl) - Val(mConsume)
        
        dB.Execute "INSERT INTO TblMeterStock(MeterCode,OpStock,Issue,Avbl,Consume,Diff) VALUES(" & Val(Vc1) & ", " & Val(mOpStock) & ", " & Val(mIssue) & ", " & Val(mAvbl) & ", " & Val(mConsume) & ", " & Val(mDiff) & " )"
                    
        RS.MoveNext
    Next
End Sub

Private Sub TxtToCode_KeyPress(KeyAscii As Integer)
    On Error Resume Next
        Cng KeyAscii
    On Error Resume Next
End Sub

Public Sub Pause(ByVal Delay As Single)
    Dim X As Single
    X = Timer + Delay                  ' Add a delay to the current time
    Do While X > Timer                 ' and waits for the current time
        DoEvents                       ' to catch up.
    Loop
End Sub

Public Sub CashBookData()
    Dim CashAccountCode     As Integer
    Dim DrCash              As Single
    Dim CrCash              As Single
    Dim LastCash            As Single
    Dim CrSale              As Single
    Dim CashSale            As Single
    Dim Receipts            As Single
    Dim CrPurchase          As Single
    Dim CashPurchase        As Single
    Dim Payments            As Single
    Dim Parties             As Single
    Dim CreditTotal         As Single
    Dim DebitTotal          As Single
    
'=============== Getting Last Day Cash (before the given date)
        
    Set RsLov = New ADODB.Recordset
        If RsLov.State = 1 Then RsLov.Close
            RsLov.Open "Select Code from Accounts Where AccType In ('Cash')", dB, adOpenStatic, adLockOptimistic
                
                If RsLov.EOF = True Then
                    MsgBox "Cash Account not found", vbInformation, "Message"
                    Exit Sub
                Else
                    CashAccountCode = Val(RsLov(0))
                End If
                                
    Set RsLov = New ADODB.Recordset
        If RsLov.State = 1 Then RsLov.Close
            RsLov.Open "Select Sum(Amount) from DrtGlt Where AccountCode = " & Val(CashAccountCode) & " And Tdate < #" & TxtFrom.Value & "#", dB, adOpenStatic, adLockOptimistic
                If IsNull(RsLov(0)) Then
                    DrCash = 0
                Else
                    DrCash = Val(RsLov(0))
                End If
                
    Set RsLov = New ADODB.Recordset
        If RsLov.State = 1 Then RS.Close
            RsLov.Open "Select Sum(Amount) from CrtGlt Where AccountCode = " & Val(CashAccountCode) & " And Tdate < #" & TxtFrom.Value & "#", dB, adOpenStatic, adLockOptimistic
                 
                If IsNull(RsLov(0)) Then
                    CrCash = 0
                Else
                    CrCash = Val(RsLov(0))
                End If
    
    LastCash = Val(DrCash) - Val(CrCash)
    

'=============== Getting Today's Credit Sale
    Set RsLov = New ADODB.Recordset
        If RsLov.State = 1 Then RsLov.Close
            RsLov.Open "Select Sum(TotalAmount) from Sale Where SaleType = '" & "Credit" & "' And Tdate = #" & TxtFrom.Value & "#", dB, adOpenStatic, adLockOptimistic
                
                If IsNull(RsLov(0)) Then
                    CrSale = 0
                Else
                    CrSale = Val(RsLov(0))
                End If
            
'=============== Getting Today's Cash Sale
    Set RsLov = New ADODB.Recordset
        If RsLov.State = 1 Then RsLov.Close
            RsLov.Open "Select Sum(TotalAmount) from Sale Where SaleType = '" & "Cash" & "' And Tdate = #" & TxtFrom.Value & "#", dB, adOpenStatic, adLockOptimistic
                
                If IsNull(RsLov(0)) Then
                    CashSale = 0
                Else
                    CashSale = Val(RsLov(0))
                End If

'=============== Getting Today's Receipts
    Set RsLov = New ADODB.Recordset
        If RsLov.State = 1 Then RsLov.Close
            RsLov.Open "Select Sum(Amount) from DrtGlt Where AccountCode = " & Val(CashAccountCode) & " And Tdate = #" & TxtFrom.Value & "# And TType Not In ('SAL')", dB, adOpenStatic, adLockOptimistic
                If IsNull(RsLov(0)) Then
                    Receipts = 0
                Else
                    Receipts = Val(RsLov(0))
                End If

'--------------------- DEBIT SIDE

'=============== Getting Today's Credit Purchase
    Set RsLov = New ADODB.Recordset
        If RsLov.State = 1 Then RsLov.Close
            RsLov.Open "Select Sum(TotalAmount) from Purchase Where PurchaseType = '" & "Credit" & "' And Tdate = #" & TxtFrom.Value & "#", dB, adOpenStatic, adLockOptimistic
                
                If IsNull(RsLov(0)) Then
                    CrPurchase = 0
                Else
                    CrPurchase = Val(RsLov(0))
                End If
            
'=============== Getting Today's Cash Purchase
    Set RsLov = New ADODB.Recordset
        If RsLov.State = 1 Then RsLov.Close
            RsLov.Open "Select Sum(TotalAmount) from Purchase Where PurchaseType = '" & "Cash" & "' And Tdate = #" & TxtFrom.Value & "#", dB, adOpenStatic, adLockOptimistic
                
                If IsNull(RsLov(0)) Then
                    CashPurchase = 0
                Else
                    CashPurchase = Val(RsLov(0))
                End If

'=============== Getting Today's Payments
    Set RsLov = New ADODB.Recordset
        If RsLov.State = 1 Then RsLov.Close
            RsLov.Open "Select Sum(Amount) from CrtGlt Where AccountCode = " & Val(CashAccountCode) & " And Tdate = #" & TxtFrom.Value & "# And TType Not In ('PUR')", dB, adOpenStatic, adLockOptimistic
                If IsNull(RsLov(0)) Then
                    Payments = 0
                Else
                    Payments = Val(RsLov(0))
                End If

'=============== Getting Today's Credit Sale (To balance on both sides with the caption name {Parties})
    Set RsLov = New ADODB.Recordset
        If RsLov.State = 1 Then RsLov.Close
            RsLov.Open "Select Sum(TotalAmount) from Sale Where SaleType = '" & "Credit" & "' And Tdate = #" & TxtFrom.Value & "#", dB, adOpenStatic, adLockOptimistic
                
                If IsNull(RsLov(0)) Then
                    Parties = 0
                Else
                    Parties = Val(RsLov(0))
                End If

    RepCashBook.Sections("Section4").Controls("LblDate").Caption = Format(TxtFrom, "dd/mm/yyyy")

'Showing Totals In the Report
'CREDIT SIDE
    RepCashBook.Sections("Section2").Controls("LblLastCash").Caption = Val(LastCash)
    RepCashBook.Sections("Section2").Controls("LblCreditSale").Caption = Val(CrSale)
    RepCashBook.Sections("Section2").Controls("LblCashSale").Caption = Val(CashSale)
    RepCashBook.Sections("Section2").Controls("LblCashReceipts").Caption = Val(Receipts)
    
'DEBIT SIDE
    RepCashBook.Sections("Section2").Controls("LblCreditPurchase").Caption = Val(CrPurchase)
    RepCashBook.Sections("Section2").Controls("LblCashPurchase").Caption = Val(CashPurchase)
    RepCashBook.Sections("Section2").Controls("LblAllPayment").Caption = Val(Payments)
    RepCashBook.Sections("Section2").Controls("LblParties").Caption = Val(Parties)
    
        
'BALANCING CREDIT SIDE TOTAL FOR CREDIT PURCHASE
    RepCashBook.Sections("Section2").Controls("LblCreditParties").Caption = Val(CrPurchase)
        
        

'CREDIT SIDE TOTAL
    CreditTotal = Val(LastCash) + Val(CrSale) + Val(CashSale) + Val(Receipts) + Val(CrPurchase)
    RepCashBook.Sections("Section2").Controls("LblCreditTotal").Caption = Val(CreditTotal)

'DEBIT SIDE TOTAL
    DebitTotal = Val(CrPurchase) + Val(CashPurchase) + Val(Payments) + Val(Parties)
    RepCashBook.Sections("Section2").Controls("LblDebitTotal").Caption = Val(DebitTotal)


'BALANCE TOTAL
    RepCashBook.Sections("Section2").Controls("LblTotal").Caption = Val(CreditTotal) - Val(DebitTotal)
            
        
    RepCashBook.Show 1
End Sub

Public Sub SetSearchGrid()
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
    Select Case VRun
        Case "Accounts Status"
            SQLQry = "Select AcId, AcTitle, AcType from ViewHeadWise where AcType = '" & TxtLookup.Text & "'"
    
        Case "Accounts Ledger"
            SQLQry = "Select AcId, AcTitle, AcType from ViewHeadWise where AcType NOT in ('Product')"
    
        Case "Supplier Wise Purchase History"
            SQLQry = "Select AcId, AcTitle, AcType from ViewHeadWise where AcType  in ('Customer','Supplier')"
    
        Case "Product Wise Purchase History"
            SQLQry = "Select AcId, AcTitle, AcType from ViewHeadWise where AcType  in ('Product')"
    
        Case "Customer Wise Sale History"
            SQLQry = "Select AcId, AcTitle, AcType from ViewHeadWise where AcType  in ('Customer','Supplier')"
        
        Case "Product Wise Sale History"
            SQLQry = "Select AcId, AcTitle, AcType from ViewHeadWise where AcType  in ('Product')"
    
    
    End Select
    
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
        
        RS.Open SQLQry, Con, adOpenStatic, adLockReadOnly
            If RS.RecordCount <= 0 Then
                MsgBox "No data found", vbInformation, "Message"
                Exit Sub
            End If
        Set MshSearch.DataSource = RS
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

Public Sub ReOrderStatus()
    dataRepReOrderPoint.Show vbModal
End Sub

Public Sub PurchaseHistoryByProduct()
    dataRepPurchaseByProduct.Show vbModal
End Sub

Public Sub SaleHistoryByProduct()
    dataRepSaleByProduct.Show vbModal
End Sub

Public Sub ProductsToReorder()
    dataRepProductToReorder.Show vbModal
End Sub
Public Sub TrialBalanceData()
Dim AcId As Single
Dim AcTitle As String
Dim SumofDR As Single
Dim SumofCr As Single
Dim Balance As Single
Dim DrBal As Single
Dim CrBal As Single

    Con.Execute "Delete * from RepTrialBalance"
    
'Copying data from ViewTrialBalance to RepTrialBalance
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
        RS.Open "Select * from ViewTrialBalance", Con, adOpenStatic, adLockOptimistic
            While Not RS.EOF = True
                AcId = Val(RS(0))
                AcTitle = RS(1)
                SumofDR = Val(RS(2))
                SumofCr = Val(RS(3))
                Balance = Val(RS(2)) - Val(RS(3))
                
                If Balance > 0 Then
                    DrBal = Abs(Balance)
                    CrBal = 0
                ElseIf Balance < 0 Then
                    CrBal = Abs(Balance)
                    DrBal = 0
                Else
                    DrBal = 0
                    CrBal = 0
                End If
                
                Con.Execute "Insert into RepTrialBalance(Id,Title,Dr,Cr) Values (" & Val(AcId) & ", '" & AcTitle & "', " & Val(DrBal) & ", " & Val(CrBal) & ")"
                RS.MoveNext
                
            Wend
            
            
'Getting Total of both (Debit / Credit Side)
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
        RS.Open "Select Sum(Dr),Sum(Cr) from RepTrialBalance", Con, adOpenStatic, adLockOptimistic
            
            On Error Resume Next
            
            TrialDr = Val(RS(0))
            TrialCr = Val(RS(1))
            

'=============================================REPORT===============================================
            
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close

    RS.Open "Select * from RepTrialBalance", Con, adOpenStatic, adLockOptimistic

        With DataRepTrialBalance
            Set .DataSource = RS
            .DataMember = RS.DataMember

            .Sections("Section1").Controls("Text1").DataField = "Id"
            .Sections("Section1").Controls("Text2").DataField = "Title"
            .Sections("Section1").Controls("Text3").DataField = "DR"
            .Sections("Section1").Controls("Text4").DataField = "CR"
        End With
            
            

End Sub

Public Sub IncomeStatData()
    Dim NetSale                         As Single
    Dim COG                             As Single
    Dim ProfitB4OtherIncome             As Single
    Dim OtherIncome                     As Single
    Dim Expense                         As Single
    Dim NetProfit                       As Single

    mDateQry = " BETWEEN #" & FrmRep.TxtFrom.Value & "# And #" & FrmRep.TxtTo.Value & "#"


' NET SALES OF THE DEFINED PERIOD
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
        RS.Open "Select Sum(TotalAmount) from Sale_Main where SaleDate" & mDateQry, Con, adOpenStatic, adLockOptimistic
            If IsNull(RS(0)) Then
                NetSale = 0
            Else
                NetSale = Val(RS(0))
            End If

    Con.Execute "Update RepIncomeStat Set Amount = " & Val(NetSale) & " where Id =1"


' COST OF GOODS SOLD
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
        RS.Open "Select Sum(Qty*PurAvg) from ProfitLoss where SaleDate" & mDateQry, Con, adOpenStatic, adLockOptimistic
            If IsNull(RS(0)) Then
                COG = 0
            Else
                COG = Val(RS(0))
            End If

    Con.Execute "Update RepIncomeStat Set Amount = " & Val(COG) & " where Id =2"


' GROSS PROFIT BEFORE OTHER INCOME
    ProfitB4OtherIncome = Val(NetSale) - Val(COG)

    Con.Execute "Update RepIncomeStat Set Amount = " & Val(ProfitB4OtherIncome) & " where Id =3"

 
' GET OTHER INCOME
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
        RS.Open "Select Sum(OtherIncome) from ViewOtherIncome where TransDate " & mDateQry, Con, adOpenStatic, adLockOptimistic
            If IsNull(RS(0)) Then
                OtherIncome = 0
            Else
                OtherIncome = Val(RS(0))
            End If

    Con.Execute "Update RepIncomeStat Set Amount = " & Val(OtherIncome) & " where Id =4"

' GET GROSS PROFIT
    Gross = Val(ProfitB4OtherIncome) + Val(OtherIncome)

    Con.Execute "Update RepIncomeStat Set Amount = " & Val(Gross) & " where Id =5"


' GET EXPENSE
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
        RS.Open "Select Sum(TotalExp) from ViewExpense where TransDate " & mDateQry, Con, adOpenStatic, adLockOptimistic
            If IsNull(RS(0)) Then
                Expense = 0
            Else
                Expense = Val(RS(0))
            End If

    Con.Execute "Update RepIncomeStat Set Amount = " & Val(Expense) & " where Id =6"

' NET PROFIT
    NetProfit = Val(Gross) - Val(Expense)
    Con.Execute "Update RepIncomeStat Set Amount = " & Val(NetProfit) & " where Id =7"
    
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
        RS.Open "Select * from RepIncomeStat", Con, adOpenStatic, adLockReadOnly
            
    With dataRepIncomeStat
        
        Set .DataSource = RS
        .DataMember = RS.DataMember
        
        .Sections("Section2").Controls("lblNetSale").Caption = Val(NetSale)
        .Sections("Section2").Controls("lblCOG").Caption = Val(COG)
        .Sections("Section2").Controls("lblGross").Caption = Val(ProfitB4OtherIncome)
        .Sections("Section2").Controls("lblOtherIncome").Caption = Val(OtherIncome)
        .Sections("Section2").Controls("lblExpense").Caption = Val(Expense)
        .Sections("Section2").Controls("lblNetProfit").Caption = Val(NetProfit)

        .Sections("Section4").Controls("lblPeriod").Caption = "From " & FrmRep.TxtFrom.Value & "  To " & FrmRep.TxtTo.Value
    
        .Show vbModal
    
    End With
    
    

End Sub
